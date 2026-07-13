using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.Data.Sqlite;

namespace EmailReviewViewer.App.Data;

public static class EmailNdjsonImporter
{
    private const int BatchSize = 1000;

    public static async Task<EmailImportResult> ImportAsync(
        string inputPath,
        string databasePath,
        int? expectedCount = null,
        CancellationToken cancellationToken = default)
    {
        inputPath = Path.GetFullPath(inputPath);
        databasePath = Path.GetFullPath(databasePath);
        if (!File.Exists(inputPath))
            throw new FileNotFoundException("NDJSON input was not found.", inputPath);
        if (expectedCount < 0)
            throw new ArgumentOutOfRangeException(nameof(expectedCount));

        Directory.CreateDirectory(Path.GetDirectoryName(databasePath)!);
        var temporaryPath = databasePath + ".importing";
        DeleteDatabaseFiles(temporaryPath);

        var inputCount = 0;
        try
        {
            var repository = new EmailRepository(temporaryPath);
            await repository.EnsureCreatedAsync(cancellationToken);
            var batch = new List<EmailMessage>(BatchSize);
            using var reader = new StreamReader(inputPath, new UTF8Encoding(false, true));
            while (await reader.ReadLineAsync(cancellationToken) is { } line)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                var record = JsonSerializer.Deserialize<ImportRecord>(line, JsonOptions)
                    ?? throw new InvalidDataException($"NDJSON line {inputCount + 1} is null.");
                inputCount++;
                batch.Add(ToMessage(record, inputCount));
                if (batch.Count == BatchSize)
                {
                    await repository.InsertBatchAsync(batch, cancellationToken);
                    batch.Clear();
                }
            }
            if (batch.Count > 0)
                await repository.InsertBatchAsync(batch, cancellationToken);

            var importedCount = (int)(await repository.SearchPageAsync(
                new EmailQuery(Limit: 1), cancellationToken)).TotalCount;
            if (expectedCount.HasValue && importedCount != expectedCount.Value)
                throw new InvalidDataException(
                    $"Imported row count {importedCount} does not match expected count {expectedCount.Value}.");

            SqliteConnection.ClearAllPools();
            File.Move(temporaryPath, databasePath, true);
            return new EmailImportResult(inputCount, importedCount, databasePath);
        }
        catch
        {
            SqliteConnection.ClearAllPools();
            DeleteDatabaseFiles(temporaryPath);
            throw;
        }
    }

    private static EmailMessage ToMessage(ImportRecord record, long id)
    {
        var entryId = Clean(record.EntryId);
        if (entryId.Length == 0)
        {
            var identity = string.Join("\n",
                Clean(record.FolderPath), Clean(record.SenderAddress), Clean(record.Subject),
                record.SentUtc?.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture) ?? "",
                record.ReceivedUtc?.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture) ?? "",
                Clean(record.BodyText));
            entryId = "FALLBACK-" + Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(identity)));
        }

        var body = record.BodyText ?? "";
        var preview = Clean(record.Preview);
        if (preview.Length == 0)
            preview = body.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries)
                .Select(value => value.Trim()).FirstOrDefault(value => value.Length > 0) ?? "";
        if (preview.Length > 500)
            preview = preview[..500];

        return new EmailMessage
        {
            Id = id,
            FolderPath = Clean(record.FolderPath),
            SenderName = Clean(record.SenderName),
            SenderAddress = Clean(record.SenderAddress),
            ToRecipients = Clean(record.ToRecipients),
            CcRecipients = Clean(record.CcRecipients),
            Subject = Clean(record.Subject),
            SentUtc = record.SentUtc?.UtcDateTime,
            ReceivedUtc = record.ReceivedUtc?.UtcDateTime,
            Preview = preview,
            BodyText = body,
            MessageClass = Clean(record.MessageClass),
            EntryId = entryId,
            ConversationId = Clean(record.ConversationId),
            ConversationTopic = Clean(record.ConversationTopic)
        };
    }

    private static void DeleteDatabaseFiles(string path)
    {
        foreach (var candidate in new[] { path, path + "-wal", path + "-shm" })
            if (File.Exists(candidate))
                File.Delete(candidate);
    }

    private static string Clean(string? value) => value?.Trim() ?? "";

    private sealed record ImportRecord(
        string? FolderPath,
        string? SenderName,
        string? SenderAddress,
        string? ToRecipients,
        string? CcRecipients,
        string? Subject,
        DateTimeOffset? SentUtc,
        DateTimeOffset? ReceivedUtc,
        string? Preview,
        string? BodyText,
        string? MessageClass,
        string? EntryId,
        string? ConversationId,
        string? ConversationTopic);

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };
}

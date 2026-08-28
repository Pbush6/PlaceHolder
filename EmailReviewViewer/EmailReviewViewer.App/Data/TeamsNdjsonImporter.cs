using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.Data.Sqlite;

namespace EmailReviewViewer.App.Data;

public static class TeamsNdjsonImporter
{
    private const int BatchSize = 1000;

    public static async Task<TeamsImportResult> ImportAsync(
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
            var repository = new TeamsRepository(temporaryPath);
            await repository.EnsureCreatedAsync(cancellationToken);
            var batch = new List<TeamsMessage>(BatchSize);
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

            var importedCount = (int)await repository.CountMessagesAsync(cancellationToken);
            if (expectedCount.HasValue && importedCount != expectedCount.Value)
                throw new InvalidDataException(
                    $"Imported row count {importedCount} does not match expected count {expectedCount.Value}.");

            SqliteConnection.ClearAllPools();
            File.Move(temporaryPath, databasePath, true);
            return new TeamsImportResult(inputCount, importedCount, databasePath);
        }
        catch
        {
            SqliteConnection.ClearAllPools();
            DeleteDatabaseFiles(temporaryPath);
            throw;
        }
    }

    private static TeamsMessage ToMessage(ImportRecord record, long id)
    {
        var entryId = Clean(record.EntryId);
        if (entryId.Length == 0)
        {
            var identity = string.Join("\n",
                Clean(record.FolderPath), Clean(record.SenderAddress), Clean(record.ConversationKey),
                Clean(record.Subject),
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

        return new TeamsMessage
        {
            Id = id,
            FolderPath = Clean(record.FolderPath),
            SenderName = Clean(record.SenderName),
            SenderAddress = Clean(record.SenderAddress),
            SenderDisplay = FirstNonEmpty(Clean(record.SenderDisplay), Clean(record.SenderName), "(unknown sender)"),
            Participants = Clean(record.Participants),
            ConversationKey = FirstNonEmpty(Clean(record.ConversationKey), $"row-{id}"),
            ConversationTitle = FirstNonEmpty(Clean(record.ConversationTitle), Clean(record.Subject), "(no subject)"),
            Subject = Clean(record.Subject),
            SentUtc = record.SentUtc?.UtcDateTime,
            ReceivedUtc = record.ReceivedUtc?.UtcDateTime,
            Preview = preview,
            BodyText = body,
            MessageClass = Clean(record.MessageClass),
            EntryId = entryId,
            AttachmentsText = Clean(record.AttachmentsText),
            ToRecipients = Clean(record.ToRecipients),
            CcRecipients = Clean(record.CcRecipients)
        };
    }

    private static void DeleteDatabaseFiles(string path)
    {
        foreach (var candidate in new[] { path, path + "-wal", path + "-shm" })
            if (File.Exists(candidate))
                File.Delete(candidate);
    }

    private static string Clean(string? value) => value?.Trim() ?? "";

    private static string FirstNonEmpty(params string[] values) =>
        values.FirstOrDefault(value => value.Length > 0) ?? "";

    private sealed record ImportRecord(
        string? FolderPath,
        string? SenderName,
        string? SenderAddress,
        string? SenderDisplay,
        string? Participants,
        string? ConversationKey,
        string? ConversationTitle,
        string? Subject,
        DateTimeOffset? SentUtc,
        DateTimeOffset? ReceivedUtc,
        string? Preview,
        string? BodyText,
        string? MessageClass,
        string? EntryId,
        string? AttachmentsText,
        string? ToRecipients,
        string? CcRecipients);

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };
}

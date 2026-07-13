using Microsoft.Data.Sqlite;

namespace EmailReviewViewer.App.Data;

public sealed class EmailDatabaseStoreSwitcher : IDisposable
{
    public EmailRepository? Repository { get; private set; }
    public string? DatabasePath => Repository?.DatabasePath;

    public async Task SwitchAsync(string databasePath, CancellationToken cancellationToken = default)
    {
        var fullPath = Path.GetFullPath(databasePath);
        await ValidateAsync(fullPath, cancellationToken);
        var next = new EmailRepository(fullPath);
        var previous = Repository;
        Repository = next;
        previous?.Dispose();
    }

    public void Dispose()
    {
        Repository?.Dispose();
        Repository = null;
    }

    private static async Task ValidateAsync(string databasePath, CancellationToken cancellationToken)
    {
        if (!File.Exists(databasePath))
            throw new InvalidDataException($"Database file not found:{Environment.NewLine}{databasePath}");

        try
        {
            var connectionString = new SqliteConnectionStringBuilder
            {
                DataSource = databasePath,
                Mode = SqliteOpenMode.ReadOnly
            }.ToString();
            await using var connection = new SqliteConnection(connectionString);
            await connection.OpenAsync(cancellationToken);

            await using (var integrity = connection.CreateCommand())
            {
                integrity.CommandText = "PRAGMA quick_check;";
                var result = Convert.ToString(await integrity.ExecuteScalarAsync(cancellationToken));
                if (!string.Equals(result, "ok", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("The selected database is corrupt or failed SQLite validation.");
            }

            var objects = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            await using (var command = connection.CreateCommand())
            {
                command.CommandText = """
                    SELECT name
                    FROM sqlite_master
                    WHERE name IN ('EmailMessages', 'EmailMessagesFts');
                    """;
                await using var reader = await command.ExecuteReaderAsync(cancellationToken);
                while (await reader.ReadAsync(cancellationToken))
                    objects.Add(reader.GetString(0));
            }

            var requiredColumns = new HashSet<string>(
            [
                "Id", "FolderPath", "SenderName", "SenderAddress", "ToRecipients", "CcRecipients",
                "Subject", "SentUtc", "ReceivedUtc", "Preview", "BodyText", "MessageClass", "EntryId",
                "ConversationId", "ConversationTopic"
            ], StringComparer.OrdinalIgnoreCase);
            await using (var command = connection.CreateCommand())
            {
                command.CommandText = "PRAGMA table_info(EmailMessages);";
                await using var reader = await command.ExecuteReaderAsync(cancellationToken);
                while (await reader.ReadAsync(cancellationToken))
                    requiredColumns.Remove(reader.GetString(1));
            }

            if (!objects.Contains("EmailMessages") ||
                !objects.Contains("EmailMessagesFts") ||
                requiredColumns.Count > 0)
            {
                throw new InvalidDataException(
                    "The selected file is not a compatible Email Review Viewer database.");
            }
        }
        catch (InvalidDataException)
        {
            throw;
        }
        catch (Exception exception) when (exception is SqliteException or IOException or UnauthorizedAccessException)
        {
            throw new InvalidDataException(
                $"The selected file is not a valid, readable Email Review Viewer database. {exception.Message}",
                exception);
        }
    }
}

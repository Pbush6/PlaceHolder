using System.Globalization;
using System.Text;
using Microsoft.Data.Sqlite;

namespace EmailReviewViewer.App.Data;

public sealed class EmailRepository(string databasePath)
{
    private readonly string _connectionString = new SqliteConnectionStringBuilder
    {
        DataSource = Path.GetFullPath(databasePath),
        Mode = SqliteOpenMode.ReadWriteCreate,
        Cache = SqliteCacheMode.Shared
    }.ToString();

    public string DatabasePath { get; } = Path.GetFullPath(databasePath);

    public async Task EnsureCreatedAsync(CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(DatabasePath)!);
        await using var connection = await OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        command.CommandText = """
            PRAGMA journal_mode = WAL;
            PRAGMA synchronous = NORMAL;
            CREATE TABLE IF NOT EXISTS EmailMessages (
                Id INTEGER PRIMARY KEY,
                FolderPath TEXT NOT NULL,
                SenderName TEXT NOT NULL,
                SenderAddress TEXT NOT NULL,
                ToRecipients TEXT NOT NULL,
                CcRecipients TEXT NOT NULL,
                Subject TEXT NOT NULL,
                SentUtc TEXT NULL,
                ReceivedUtc TEXT NULL,
                Preview TEXT NOT NULL,
                BodyText TEXT NOT NULL,
                MessageClass TEXT NOT NULL,
                EntryId TEXT NOT NULL,
                ConversationId TEXT NOT NULL DEFAULT '',
                ConversationTopic TEXT NOT NULL DEFAULT ''
            );
            CREATE INDEX IF NOT EXISTS IX_EmailMessages_ReceivedUtc ON EmailMessages(ReceivedUtc DESC);
            CREATE INDEX IF NOT EXISTS IX_EmailMessages_SentUtc ON EmailMessages(SentUtc DESC);
            CREATE INDEX IF NOT EXISTS IX_EmailMessages_SenderAddress ON EmailMessages(SenderAddress);
            CREATE UNIQUE INDEX IF NOT EXISTS UX_EmailMessages_EntryId ON EmailMessages(EntryId);
            CREATE INDEX IF NOT EXISTS IX_EmailMessages_FolderPath ON EmailMessages(FolderPath);
            CREATE VIRTUAL TABLE IF NOT EXISTS EmailMessagesFts USING fts5(
                Subject, BodyText, SenderName, SenderAddress, ToRecipients, CcRecipients,
                content='EmailMessages', content_rowid='Id', tokenize='unicode61'
            );
            CREATE TRIGGER IF NOT EXISTS EmailMessages_ai AFTER INSERT ON EmailMessages BEGIN
                INSERT INTO EmailMessagesFts(
                    rowid, Subject, BodyText, SenderName, SenderAddress, ToRecipients, CcRecipients
                ) VALUES (
                    new.Id, new.Subject, new.BodyText, new.SenderName, new.SenderAddress,
                    new.ToRecipients, new.CcRecipients
                );
            END;
            CREATE TRIGGER IF NOT EXISTS EmailMessages_ad AFTER DELETE ON EmailMessages BEGIN
                INSERT INTO EmailMessagesFts(
                    EmailMessagesFts, rowid, Subject, BodyText, SenderName, SenderAddress,
                    ToRecipients, CcRecipients
                ) VALUES (
                    'delete', old.Id, old.Subject, old.BodyText, old.SenderName, old.SenderAddress,
                    old.ToRecipients, old.CcRecipients
                );
            END;
            CREATE TRIGGER IF NOT EXISTS EmailMessages_au AFTER UPDATE ON EmailMessages BEGIN
                INSERT INTO EmailMessagesFts(
                    EmailMessagesFts, rowid, Subject, BodyText, SenderName, SenderAddress,
                    ToRecipients, CcRecipients
                ) VALUES (
                    'delete', old.Id, old.Subject, old.BodyText, old.SenderName, old.SenderAddress,
                    old.ToRecipients, old.CcRecipients
                );
                INSERT INTO EmailMessagesFts(
                    rowid, Subject, BodyText, SenderName, SenderAddress, ToRecipients, CcRecipients
                ) VALUES (
                    new.Id, new.Subject, new.BodyText, new.SenderName, new.SenderAddress,
                    new.ToRecipients, new.CcRecipients
                );
            END;
            """;
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task InsertBatchAsync(
        IEnumerable<EmailMessage> messages,
        CancellationToken cancellationToken = default)
    {
        await using var connection = await OpenAsync(cancellationToken);
        await using var transaction = await connection.BeginTransactionAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        command.Transaction = (SqliteTransaction)transaction;
        command.CommandText = """
            INSERT INTO EmailMessages (
                Id, FolderPath, SenderName, SenderAddress, ToRecipients, CcRecipients,
                Subject, SentUtc, ReceivedUtc, Preview, BodyText, MessageClass, EntryId,
                ConversationId, ConversationTopic
            ) VALUES (
                $id, $folder, $senderName, $senderAddress, $to, $cc,
                $subject, $sent, $received, $preview, $body, $class, $entryId,
                $conversationId, $conversationTopic
            )
            ON CONFLICT(EntryId) DO UPDATE SET
                FolderPath=excluded.FolderPath,
                SenderName=excluded.SenderName,
                SenderAddress=excluded.SenderAddress,
                ToRecipients=excluded.ToRecipients,
                CcRecipients=excluded.CcRecipients,
                Subject=excluded.Subject,
                SentUtc=excluded.SentUtc,
                ReceivedUtc=excluded.ReceivedUtc,
                Preview=excluded.Preview,
                BodyText=excluded.BodyText,
                MessageClass=excluded.MessageClass,
                ConversationId=excluded.ConversationId,
                ConversationTopic=excluded.ConversationTopic;
            """;

        var parameters = new Dictionary<string, SqliteParameter>();
        foreach (var name in new[]
                 {
                     "$id", "$folder", "$senderName", "$senderAddress", "$to", "$cc", "$subject",
                     "$sent", "$received", "$preview", "$body", "$class", "$entryId",
                     "$conversationId", "$conversationTopic"
                 })
        {
            parameters[name] = command.Parameters.Add(name, SqliteType.Text);
        }

        foreach (var message in messages)
        {
            parameters["$id"].Value = message.Id;
            parameters["$folder"].Value = message.FolderPath;
            parameters["$senderName"].Value = message.SenderName;
            parameters["$senderAddress"].Value = message.SenderAddress;
            parameters["$to"].Value = message.ToRecipients;
            parameters["$cc"].Value = message.CcRecipients;
            parameters["$subject"].Value = message.Subject;
            parameters["$sent"].Value = DbDate(message.SentUtc);
            parameters["$received"].Value = DbDate(message.ReceivedUtc);
            parameters["$preview"].Value = message.Preview;
            parameters["$body"].Value = message.BodyText;
            parameters["$class"].Value = message.MessageClass;
            parameters["$entryId"].Value = message.EntryId;
            parameters["$conversationId"].Value = message.ConversationId;
            parameters["$conversationTopic"].Value = message.ConversationTopic;
            await command.ExecuteNonQueryAsync(cancellationToken);
        }

        await transaction.CommitAsync(cancellationToken);
    }

    public async Task<EmailPage> SearchPageAsync(
        EmailQuery query,
        CancellationToken cancellationToken = default)
    {
        var limit = Math.Clamp(query.Limit, 1, 1000);
        var offset = Math.Max(query.Offset, 0);
        await using var connection = await OpenAsync(cancellationToken);
        var (fromSql, whereSql) = BuildFilterSql(query);

        await using var countCommand = connection.CreateCommand();
        countCommand.CommandText = $"SELECT COUNT(*) FROM EmailMessages m {fromSql} {whereSql};";
        BindFilters(countCommand, query);
        var totalCount = (long)(await countCommand.ExecuteScalarAsync(cancellationToken) ?? 0L);

        await using var pageCommand = connection.CreateCommand();
        pageCommand.CommandText = $"""
            SELECT m.Id, COALESCE(m.ReceivedUtc, m.SentUtc), m.SenderName, m.SenderAddress,
                   m.ToRecipients, m.Subject, m.Preview
            FROM EmailMessages m
            {fromSql}
            {whereSql}
            ORDER BY COALESCE(m.ReceivedUtc, m.SentUtc) DESC, m.Id DESC
            LIMIT $limit OFFSET $offset;
            """;
        BindFilters(pageCommand, query);
        pageCommand.Parameters.AddWithValue("$limit", limit);
        pageCommand.Parameters.AddWithValue("$offset", offset);

        var items = new List<EmailListItem>(limit);
        await using var reader = await pageCommand.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            items.Add(new EmailListItem(
                reader.GetInt64(0),
                ReadDate(reader, 1),
                reader.GetString(2),
                reader.GetString(3),
                reader.GetString(4),
                reader.GetString(5),
                reader.GetString(6)));
        }

        return new EmailPage(items, totalCount);
    }

    public async Task<IReadOnlyList<FolderCount>> GetFolderCountsAsync(
        CancellationToken cancellationToken = default)
    {
        await using var connection = await OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT FolderPath, COUNT(*)
            FROM EmailMessages
            GROUP BY FolderPath
            ORDER BY FolderPath COLLATE NOCASE;
            """;

        var folders = new List<FolderCount>();
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
            folders.Add(new FolderCount(reader.GetString(0), reader.GetInt64(1)));
        return folders;
    }

    public async Task<EmailMessage?> GetByIdAsync(long id, CancellationToken cancellationToken = default)
    {
        await using var connection = await OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT Id, FolderPath, SenderName, SenderAddress, ToRecipients, CcRecipients,
                   Subject, SentUtc, ReceivedUtc, Preview, BodyText, MessageClass, EntryId,
                   ConversationId, ConversationTopic
            FROM EmailMessages
            WHERE Id = $id;
            """;
        command.Parameters.AddWithValue("$id", id);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken))
            return null;

        return new EmailMessage
        {
            Id = reader.GetInt64(0),
            FolderPath = reader.GetString(1),
            SenderName = reader.GetString(2),
            SenderAddress = reader.GetString(3),
            ToRecipients = reader.GetString(4),
            CcRecipients = reader.GetString(5),
            Subject = reader.GetString(6),
            SentUtc = ReadDate(reader, 7),
            ReceivedUtc = ReadDate(reader, 8),
            Preview = reader.GetString(9),
            BodyText = reader.GetString(10),
            MessageClass = reader.GetString(11),
            EntryId = reader.GetString(12),
            ConversationId = reader.GetString(13),
            ConversationTopic = reader.GetString(14)
        };
    }

    private async Task<SqliteConnection> OpenAsync(CancellationToken cancellationToken)
    {
        var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        command.CommandText = "PRAGMA busy_timeout = 5000;";
        await command.ExecuteNonQueryAsync(cancellationToken);
        return connection;
    }

    private static (string FromSql, string WhereSql) BuildFilterSql(EmailQuery query)
    {
        var clauses = new List<string>();
        var hasKeyword = !string.IsNullOrWhiteSpace(query.Keyword);
        if (hasKeyword)
            clauses.Add("EmailMessagesFts MATCH $keyword");
        if (query.FromUtc.HasValue)
            clauses.Add("COALESCE(m.ReceivedUtc, m.SentUtc) >= $fromUtc");
        if (query.ToUtc.HasValue)
            clauses.Add("COALESCE(m.ReceivedUtc, m.SentUtc) <= $toUtc");
        if (!string.IsNullOrWhiteSpace(query.Party))
        {
            clauses.Add("""
                (m.SenderName LIKE $party ESCAPE '\' OR m.SenderAddress LIKE $party ESCAPE '\'
                 OR m.ToRecipients LIKE $party ESCAPE '\' OR m.CcRecipients LIKE $party ESCAPE '\')
                """);
        }
        var folderPaths = SelectedFolderPaths(query);
        if (folderPaths.Count > 0)
            clauses.Add($"m.FolderPath IN ({string.Join(", ", Enumerable.Range(0, folderPaths.Count).Select(index => $"$folder{index}"))})");

        var from = hasKeyword ? "JOIN EmailMessagesFts ON EmailMessagesFts.rowid = m.Id" : "";
        var where = clauses.Count == 0 ? "" : "WHERE " + string.Join(" AND ", clauses);
        return (from, where);
    }

    private static void BindFilters(SqliteCommand command, EmailQuery query)
    {
        if (!string.IsNullOrWhiteSpace(query.Keyword))
            command.Parameters.AddWithValue("$keyword", BuildFtsQuery(query.Keyword));
        if (query.FromUtc.HasValue)
            command.Parameters.AddWithValue("$fromUtc", DbDate(query.FromUtc));
        if (query.ToUtc.HasValue)
            command.Parameters.AddWithValue("$toUtc", DbDate(query.ToUtc));
        if (!string.IsNullOrWhiteSpace(query.Party))
            command.Parameters.AddWithValue("$party", $"%{EscapeLike(query.Party.Trim())}%");
        var folderPaths = SelectedFolderPaths(query);
        for (var index = 0; index < folderPaths.Count; index++)
            command.Parameters.AddWithValue($"$folder{index}", folderPaths[index]);
    }

    private static IReadOnlyList<string> SelectedFolderPaths(EmailQuery query) =>
        query.FolderPaths?
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Distinct(StringComparer.Ordinal)
            .ToArray() ?? [];

    private static string BuildFtsQuery(string input)
    {
        var terms = input.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        return string.Join(" AND ", terms.Select(term => $"\"{term.Replace("\"", "\"\"")}\""));
    }

    private static string EscapeLike(string input) =>
        input.Replace(@"\", @"\\").Replace("%", @"\%").Replace("_", @"\_");

    private static object DbDate(DateTime? value) =>
        value.HasValue ? value.Value.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture) : DBNull.Value;

    private static DateTime? ReadDate(SqliteDataReader reader, int ordinal) =>
        reader.IsDBNull(ordinal)
            ? null
            : DateTime.Parse(reader.GetString(ordinal), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
}

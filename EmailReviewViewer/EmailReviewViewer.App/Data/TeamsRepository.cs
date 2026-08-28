using System.Globalization;
using Microsoft.Data.Sqlite;

namespace EmailReviewViewer.App.Data;

public sealed class TeamsRepository(string databasePath) : IDisposable
{
    private readonly string _connectionString = new SqliteConnectionStringBuilder
    {
        DataSource = Path.GetFullPath(databasePath),
        Mode = SqliteOpenMode.ReadWriteCreate,
        Cache = SqliteCacheMode.Shared
    }.ToString();

    public string DatabasePath { get; } = Path.GetFullPath(databasePath);
    private bool _disposed;

    public async Task EnsureCreatedAsync(CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(DatabasePath)!);
        await using var connection = await OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        command.CommandText = """
            PRAGMA journal_mode = WAL;
            PRAGMA synchronous = NORMAL;
            CREATE TABLE IF NOT EXISTS TeamsMessages (
                Id INTEGER PRIMARY KEY,
                FolderPath TEXT NOT NULL,
                SenderName TEXT NOT NULL,
                SenderAddress TEXT NOT NULL,
                SenderDisplay TEXT NOT NULL,
                Participants TEXT NOT NULL,
                ConversationKey TEXT NOT NULL,
                ConversationTitle TEXT NOT NULL,
                Subject TEXT NOT NULL,
                SentUtc TEXT NULL,
                ReceivedUtc TEXT NULL,
                Preview TEXT NOT NULL,
                BodyText TEXT NOT NULL,
                MessageClass TEXT NOT NULL,
                EntryId TEXT NOT NULL,
                AttachmentsText TEXT NOT NULL DEFAULT '',
                ToRecipients TEXT NOT NULL DEFAULT '',
                CcRecipients TEXT NOT NULL DEFAULT ''
            );
            CREATE INDEX IF NOT EXISTS IX_TeamsMessages_ReceivedUtc ON TeamsMessages(ReceivedUtc DESC);
            CREATE INDEX IF NOT EXISTS IX_TeamsMessages_ConversationKey ON TeamsMessages(ConversationKey);
            CREATE UNIQUE INDEX IF NOT EXISTS UX_TeamsMessages_EntryId ON TeamsMessages(EntryId);
            CREATE INDEX IF NOT EXISTS IX_TeamsMessages_FolderPath ON TeamsMessages(FolderPath);
            CREATE VIRTUAL TABLE IF NOT EXISTS TeamsMessagesFts USING fts5(
                BodyText, SenderDisplay, Participants, ConversationTitle, Subject,
                content='TeamsMessages', content_rowid='Id', tokenize='unicode61'
            );
            CREATE TRIGGER IF NOT EXISTS TeamsMessages_ai AFTER INSERT ON TeamsMessages BEGIN
                INSERT INTO TeamsMessagesFts(
                    rowid, BodyText, SenderDisplay, Participants, ConversationTitle, Subject
                ) VALUES (
                    new.Id, new.BodyText, new.SenderDisplay, new.Participants,
                    new.ConversationTitle, new.Subject
                );
            END;
            CREATE TRIGGER IF NOT EXISTS TeamsMessages_ad AFTER DELETE ON TeamsMessages BEGIN
                INSERT INTO TeamsMessagesFts(
                    TeamsMessagesFts, rowid, BodyText, SenderDisplay, Participants,
                    ConversationTitle, Subject
                ) VALUES (
                    'delete', old.Id, old.BodyText, old.SenderDisplay, old.Participants,
                    old.ConversationTitle, old.Subject
                );
            END;
            CREATE TRIGGER IF NOT EXISTS TeamsMessages_au AFTER UPDATE ON TeamsMessages BEGIN
                INSERT INTO TeamsMessagesFts(
                    TeamsMessagesFts, rowid, BodyText, SenderDisplay, Participants,
                    ConversationTitle, Subject
                ) VALUES (
                    'delete', old.Id, old.BodyText, old.SenderDisplay, old.Participants,
                    old.ConversationTitle, old.Subject
                );
                INSERT INTO TeamsMessagesFts(
                    rowid, BodyText, SenderDisplay, Participants, ConversationTitle, Subject
                ) VALUES (
                    new.Id, new.BodyText, new.SenderDisplay, new.Participants,
                    new.ConversationTitle, new.Subject
                );
            END;
            """;
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task InsertBatchAsync(
        IEnumerable<TeamsMessage> messages,
        CancellationToken cancellationToken = default)
    {
        await using var connection = await OpenAsync(cancellationToken);
        await using var transaction = await connection.BeginTransactionAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        command.Transaction = (SqliteTransaction)transaction;
        command.CommandText = """
            INSERT INTO TeamsMessages (
                Id, FolderPath, SenderName, SenderAddress, SenderDisplay, Participants,
                ConversationKey, ConversationTitle, Subject, SentUtc, ReceivedUtc,
                Preview, BodyText, MessageClass, EntryId, AttachmentsText, ToRecipients, CcRecipients
            ) VALUES (
                $id, $folder, $senderName, $senderAddress, $senderDisplay, $participants,
                $conversationKey, $conversationTitle, $subject, $sent, $received,
                $preview, $body, $class, $entryId, $attachments, $toRecipients, $ccRecipients
            )
            ON CONFLICT(EntryId) DO UPDATE SET
                FolderPath=excluded.FolderPath,
                SenderName=excluded.SenderName,
                SenderAddress=excluded.SenderAddress,
                SenderDisplay=excluded.SenderDisplay,
                Participants=excluded.Participants,
                ConversationKey=excluded.ConversationKey,
                ConversationTitle=excluded.ConversationTitle,
                Subject=excluded.Subject,
                SentUtc=excluded.SentUtc,
                ReceivedUtc=excluded.ReceivedUtc,
                Preview=excluded.Preview,
                BodyText=excluded.BodyText,
                MessageClass=excluded.MessageClass,
                AttachmentsText=excluded.AttachmentsText,
                ToRecipients=excluded.ToRecipients,
                CcRecipients=excluded.CcRecipients;
            """;

        var parameters = new Dictionary<string, SqliteParameter>();
        foreach (var name in new[]
                 {
                     "$id", "$folder", "$senderName", "$senderAddress", "$senderDisplay", "$participants",
                     "$conversationKey", "$conversationTitle", "$subject", "$sent", "$received",
                     "$preview", "$body", "$class", "$entryId", "$attachments", "$toRecipients", "$ccRecipients"
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
            parameters["$senderDisplay"].Value = message.SenderDisplay;
            parameters["$participants"].Value = message.Participants;
            parameters["$conversationKey"].Value = message.ConversationKey;
            parameters["$conversationTitle"].Value = message.ConversationTitle;
            parameters["$subject"].Value = message.Subject;
            parameters["$sent"].Value = DbDate(message.SentUtc);
            parameters["$received"].Value = DbDate(message.ReceivedUtc);
            parameters["$preview"].Value = message.Preview;
            parameters["$body"].Value = message.BodyText;
            parameters["$class"].Value = message.MessageClass;
            parameters["$entryId"].Value = message.EntryId;
            parameters["$attachments"].Value = message.AttachmentsText;
            parameters["$toRecipients"].Value = message.ToRecipients;
            parameters["$ccRecipients"].Value = message.CcRecipients;
            await command.ExecuteNonQueryAsync(cancellationToken);
        }

        await transaction.CommitAsync(cancellationToken);
    }

    public async Task<long> CountMessagesAsync(CancellationToken cancellationToken = default)
    {
        await using var connection = await OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT COUNT(*) FROM TeamsMessages;";
        return (long)(await command.ExecuteScalarAsync(cancellationToken) ?? 0L);
    }

    public async Task<TeamsConversationPage> SearchConversationsAsync(
        TeamsQuery query,
        CancellationToken cancellationToken = default)
    {
        var limit = Math.Clamp(query.Limit, 1, 1000);
        var offset = Math.Max(query.Offset, 0);
        var selected = query.SelectedParticipants?
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.Ordinal)
            .ToArray() ?? [];
        await using var connection = await OpenAsync(cancellationToken);
        var (fromSql, whereSql) = BuildFilterSql(query with { Keyword = query.Keyword, Party = null });

        await using var pageCommand = connection.CreateCommand();
        pageCommand.CommandText = $"""
            SELECT m.ConversationKey,
                   MAX(m.ConversationTitle),
                   MAX(m.Participants),
                   m.SenderDisplay,
                   COUNT(*),
                   MAX(COALESCE(m.ReceivedUtc, m.SentUtc))
            FROM TeamsMessages m
            {fromSql}
            {whereSql}
            GROUP BY m.ConversationKey, m.SenderDisplay
            """;
        BindFilters(pageCommand, query with { Party = null });

        var grouped = new Dictionary<string, ConversationAggregate>(StringComparer.Ordinal);
        await using var reader = await pageCommand.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            var key = reader.GetString(0);
            var rowParticipants = reader.IsDBNull(2) ? "" : reader.GetString(2);
            if (!grouped.TryGetValue(key, out var aggregate))
            {
                aggregate = new ConversationAggregate(
                    key,
                    reader.IsDBNull(1) ? "" : reader.GetString(1),
                    rowParticipants);
                grouped[key] = aggregate;
            }
            else
                aggregate.AddParticipants(rowParticipants);
            var sender = reader.IsDBNull(3) ? "" : reader.GetString(3);
            var count = reader.GetInt64(4);
            aggregate.AddSender(sender, count, ReadDate(reader, 5));
        }

        var matched = new List<TeamsConversationListItem>(grouped.Count);
        long visibleMessages = 0;
        foreach (var aggregate in grouped.Values)
        {
            var participants = TeamsParticipantMatching.Split(aggregate.Participants);
            if (!TeamsParticipantMatching.ConversationMatches(query.MatchMode, selected, participants))
                continue;
            var visible = aggregate.VisibleCount(query.MatchMode, selected);
            if (visible <= 0)
                continue;
            matched.Add(new TeamsConversationListItem(
                aggregate.ConversationKey,
                aggregate.LastUtc,
                aggregate.ConversationTitle,
                aggregate.Participants,
                visible,
                aggregate.TotalCount));
            visibleMessages += visible;
        }

        IEnumerable<TeamsConversationListItem> ordered = query.NewestFirst
            ? matched.OrderByDescending(item => item.LastUtc).ThenBy(item => item.ConversationKey, StringComparer.Ordinal)
            : matched.OrderBy(item => item.LastUtc).ThenBy(item => item.ConversationKey, StringComparer.Ordinal);
        var page = ordered.Skip(offset).Take(limit).ToList();
        return new TeamsConversationPage(page, matched.Count, visibleMessages);
    }

    public async Task<IReadOnlyList<TeamsMessage>> GetConversationMessagesAsync(
        string conversationKey,
        CancellationToken cancellationToken = default)
    {
        await using var connection = await OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT Id, FolderPath, SenderName, SenderAddress, SenderDisplay, Participants,
                   ConversationKey, ConversationTitle, Subject, SentUtc, ReceivedUtc,
                   Preview, BodyText, MessageClass, EntryId, AttachmentsText, ToRecipients, CcRecipients
            FROM TeamsMessages
            WHERE ConversationKey = $key
            ORDER BY COALESCE(ReceivedUtc, SentUtc) ASC, Id ASC;
            """;
        command.Parameters.AddWithValue("$key", conversationKey);
        var messages = new List<TeamsMessage>();
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
            messages.Add(ReadMessage(reader));
        return messages;
    }

    public async Task<IReadOnlyList<FolderCount>> GetFolderCountsAsync(
        CancellationToken cancellationToken = default)
    {
        await using var connection = await OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT FolderPath, COUNT(*)
            FROM TeamsMessages
            GROUP BY FolderPath
            ORDER BY FolderPath COLLATE NOCASE;
            """;

        var folders = new List<FolderCount>();
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
            folders.Add(new FolderCount(reader.GetString(0), reader.GetInt64(1)));
        return folders;
    }

    public async Task<TeamsReportSummary> GetReportSummaryAsync(CancellationToken cancellationToken = default)
    {
        var folders = await GetFolderCountsAsync(cancellationToken);
        var messageCount = await CountMessagesAsync(cancellationToken);
        await using var connection = await OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT DISTINCT Participants FROM TeamsMessages;";
        var names = new HashSet<string>(StringComparer.Ordinal);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            foreach (var name in TeamsParticipantMatching.Split(reader.IsDBNull(0) ? "" : reader.GetString(0)))
                names.Add(name);
        }

        var people = names.Where(TeamsParticipantMatching.IsLikelyPersonName).OrderBy(name => name, StringComparer.OrdinalIgnoreCase).ToArray();
        var other = names.Where(name => !TeamsParticipantMatching.IsLikelyPersonName(name)).OrderBy(name => name, StringComparer.OrdinalIgnoreCase).ToArray();
        return new TeamsReportSummary(messageCount, names.Count, folders.Count, people, other);
    }

    private async Task<SqliteConnection> OpenAsync(CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        command.CommandText = "PRAGMA busy_timeout = 5000;";
        await command.ExecuteNonQueryAsync(cancellationToken);
        return connection;
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        using var connection = new SqliteConnection(_connectionString);
        SqliteConnection.ClearPool(connection);
    }

    private static TeamsMessage ReadMessage(SqliteDataReader reader) => new()
    {
        Id = reader.GetInt64(0),
        FolderPath = reader.GetString(1),
        SenderName = reader.GetString(2),
        SenderAddress = reader.GetString(3),
        SenderDisplay = reader.GetString(4),
        Participants = reader.GetString(5),
        ConversationKey = reader.GetString(6),
        ConversationTitle = reader.GetString(7),
        Subject = reader.GetString(8),
        SentUtc = ReadDate(reader, 9),
        ReceivedUtc = ReadDate(reader, 10),
        Preview = reader.GetString(11),
        BodyText = reader.GetString(12),
        MessageClass = reader.GetString(13),
        EntryId = reader.GetString(14),
        AttachmentsText = reader.GetString(15),
        ToRecipients = reader.FieldCount > 16 && !reader.IsDBNull(16) ? reader.GetString(16) : "",
        CcRecipients = reader.FieldCount > 17 && !reader.IsDBNull(17) ? reader.GetString(17) : ""
    };

    private sealed class ConversationAggregate
    {
        public string ConversationKey { get; }
        public string ConversationTitle { get; }
        public DateTime? LastUtc { get; private set; }
        public long TotalCount { get; private set; }
        private readonly List<string> _participantNames = [];
        private readonly HashSet<string> _participantSet = new(StringComparer.Ordinal);
        private readonly Dictionary<string, long> _senderCounts = new(StringComparer.Ordinal);

        public string Participants => string.Join("||", _participantNames);

        public ConversationAggregate(string conversationKey, string conversationTitle, string participants)
        {
            ConversationKey = conversationKey;
            ConversationTitle = conversationTitle;
            AddParticipants(participants);
        }

        public void AddParticipants(string value)
        {
            foreach (var name in TeamsParticipantMatching.Split(value))
            {
                if (_participantSet.Add(name))
                    _participantNames.Add(name);
            }
        }

        public void AddSender(string sender, long count, DateTime? lastUtc)
        {
            TotalCount += count;
            _senderCounts[sender] = _senderCounts.GetValueOrDefault(sender) + count;
            if (lastUtc.HasValue && (LastUtc is null || lastUtc > LastUtc))
                LastUtc = lastUtc;
        }

        public long VisibleCount(TeamsParticipantMatchMode mode, IReadOnlyList<string> selected)
        {
            if (selected.Count == 0 || mode != TeamsParticipantMatchMode.MessagesFromSelected)
                return TotalCount;
            return _senderCounts
                .Where(pair => selected.Contains(pair.Key, StringComparer.Ordinal))
                .Sum(pair => pair.Value);
        }
    }

    private static (string FromSql, string WhereSql) BuildFilterSql(TeamsQuery query)
    {
        var clauses = new List<string>();
        var hasKeyword = !string.IsNullOrWhiteSpace(query.Keyword);
        if (hasKeyword)
        {
            clauses.Add("""
                m.ConversationKey IN (
                    SELECT keyed.ConversationKey
                    FROM TeamsMessages keyed
                    JOIN TeamsMessagesFts ON TeamsMessagesFts.rowid = keyed.Id
                    WHERE TeamsMessagesFts MATCH $keyword
                )
                """);
        }
        if (query.FromUtc.HasValue)
            clauses.Add("COALESCE(m.ReceivedUtc, m.SentUtc) >= $fromUtc");
        if (query.ToUtc.HasValue)
            clauses.Add("COALESCE(m.ReceivedUtc, m.SentUtc) <= $toUtc");
        if (!string.IsNullOrWhiteSpace(query.Party))
        {
            clauses.Add("""
                (m.SenderDisplay LIKE $party ESCAPE '\' OR m.SenderName LIKE $party ESCAPE '\'
                 OR m.SenderAddress LIKE $party ESCAPE '\' OR m.Participants LIKE $party ESCAPE '\')
                """);
        }
        var folderPaths = SelectedFolderPaths(query);
        if (folderPaths.Count > 0)
            clauses.Add($"m.FolderPath IN ({string.Join(", ", Enumerable.Range(0, folderPaths.Count).Select(index => $"$folder{index}"))})");

        var from = "";
        var where = clauses.Count == 0 ? "" : "WHERE " + string.Join(" AND ", clauses);
        return (from, where);
    }

    private static void BindFilters(SqliteCommand command, TeamsQuery query)
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

    private static IReadOnlyList<string> SelectedFolderPaths(TeamsQuery query) =>
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

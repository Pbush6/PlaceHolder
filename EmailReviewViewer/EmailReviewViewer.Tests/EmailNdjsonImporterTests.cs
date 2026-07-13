using System.Text;
using EmailReviewViewer.App.Data;
using Microsoft.Data.Sqlite;

namespace EmailReviewViewer.Tests;

public sealed class EmailNdjsonImporterTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), $"email-import-{Guid.NewGuid():N}");

    [Fact]
    public async Task Import_preserves_unicode_newlines_quotes_and_builds_fts_and_folder_counts()
    {
        Directory.CreateDirectory(_directory);
        var input = Path.Combine(_directory, "messages.ndjson");
        var database = Path.Combine(_directory, "messages.db");
        await File.WriteAllTextAsync(input,
            """
            {"FolderPath":"Mailbox\\Inbox","SenderName":"Zoë","SenderAddress":"zoe@example.com","ToRecipients":"Renée <renee@example.com>","CcRecipients":"","Subject":"Quoted \"subject\"","SentUtc":"2026-07-13T15:00:00Z","ReceivedUtc":"2026-07-13T15:01:00Z","Preview":"First line","BodyText":"First line\nSecond phoenix line","MessageClass":"IPM.Note","EntryId":"entry-1","ConversationId":"conversation-1","ConversationTopic":"Unicode topic"}
            {"FolderPath":"Mailbox\\Archive","SenderName":"李雷","SenderAddress":"li@example.com","ToRecipients":"zoe@example.com","CcRecipients":"","Subject":"Archive","SentUtc":null,"ReceivedUtc":"2026-07-12T12:00:00-05:00","Preview":"Body","BodyText":"Body","MessageClass":"IPM.Note","EntryId":"entry-2","ConversationId":"","ConversationTopic":""}
            """ + Environment.NewLine, new UTF8Encoding(false));

        var result = await EmailNdjsonImporter.ImportAsync(input, database, expectedCount: 2);
        var repository = new EmailRepository(database);
        var search = await repository.SearchPageAsync(new EmailQuery(Keyword: "phoenix"));
        var folders = await repository.GetFolderCountsAsync();
        var detail = await repository.GetByIdAsync(search.Items.Single().Id);

        Assert.Equal(2, result.ImportedCount);
        Assert.Equal(2, result.InputCount);
        Assert.Equal([new FolderCount(@"Mailbox\Archive", 1), new FolderCount(@"Mailbox\Inbox", 1)], folders);
        Assert.NotNull(detail);
        Assert.Equal("First line\nSecond phoenix line", detail.BodyText);
        Assert.Equal("conversation-1", detail.ConversationId);
        Assert.Equal("Unicode topic", detail.ConversationTopic);
    }

    [Fact]
    public async Task Import_deduplicates_repeated_entry_ids_deterministically()
    {
        Directory.CreateDirectory(_directory);
        var input = Path.Combine(_directory, "duplicates.ndjson");
        var database = Path.Combine(_directory, "duplicates.db");
        await File.WriteAllTextAsync(input,
            Record("same-entry", "First") + Environment.NewLine +
            Record("same-entry", "Replacement") + Environment.NewLine);

        var result = await EmailNdjsonImporter.ImportAsync(input, database);
        var repository = new EmailRepository(database);
        var page = await repository.SearchPageAsync(new EmailQuery());

        Assert.Equal(2, result.InputCount);
        Assert.Equal(1, result.ImportedCount);
        Assert.Single(page.Items);
        Assert.Equal("Replacement", page.Items[0].Subject);
    }

    [Fact]
    public async Task Malformed_input_fails_without_replacing_existing_database()
    {
        Directory.CreateDirectory(_directory);
        var input = Path.Combine(_directory, "malformed.ndjson");
        var database = Path.Combine(_directory, "existing.db");
        var existing = new EmailRepository(database);
        await existing.EnsureCreatedAsync();
        await existing.InsertBatchAsync([Message(1, "Existing", "existing-entry")]);
        await File.WriteAllTextAsync(input, Record("new-entry", "New") + Environment.NewLine + "{bad json");

        await Assert.ThrowsAnyAsync<Exception>(() => EmailNdjsonImporter.ImportAsync(input, database, expectedCount: 2));
        SqliteConnection.ClearAllPools();
        var page = await existing.SearchPageAsync(new EmailQuery());

        Assert.Single(page.Items);
        Assert.Equal("Existing", page.Items[0].Subject);
        Assert.False(File.Exists(database + ".importing"));
    }

    [Fact]
    public async Task Expected_count_mismatch_fails_atomically()
    {
        Directory.CreateDirectory(_directory);
        var input = Path.Combine(_directory, "count.ndjson");
        var database = Path.Combine(_directory, "count.db");
        await File.WriteAllTextAsync(input, Record("one", "One") + Environment.NewLine);

        await Assert.ThrowsAsync<InvalidDataException>(
            () => EmailNdjsonImporter.ImportAsync(input, database, expectedCount: 2));

        Assert.False(File.Exists(database));
        Assert.False(File.Exists(database + ".importing"));
    }

    private static string Record(string entryId, string subject) =>
        $$"""{"FolderPath":"Mailbox\\Inbox","SenderName":"Sender","SenderAddress":"sender@example.com","ToRecipients":"to@example.com","CcRecipients":"","Subject":"{{subject}}","SentUtc":"2026-07-13T15:00:00Z","ReceivedUtc":"2026-07-13T15:01:00Z","Preview":"Preview","BodyText":"Body","MessageClass":"IPM.Note","EntryId":"{{entryId}}","ConversationId":"","ConversationTopic":""}""";

    private static EmailMessage Message(long id, string subject, string entryId) => new()
    {
        Id = id,
        FolderPath = @"Mailbox\Inbox",
        Subject = subject,
        EntryId = entryId
    };

    public void Dispose()
    {
        SqliteConnection.ClearAllPools();
        if (Directory.Exists(_directory))
            Directory.Delete(_directory, true);
    }
}

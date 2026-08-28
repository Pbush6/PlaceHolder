using System.Text;
using EmailReviewViewer.App.Data;
using Microsoft.Data.Sqlite;

namespace EmailReviewViewer.Tests;

public sealed class TeamsNdjsonImporterTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), $"teams-import-{Guid.NewGuid():N}");

    [Fact]
    public async Task Import_builds_fts_and_preserves_conversation_key_and_attachments()
    {
        Directory.CreateDirectory(_directory);
        var input = Path.Combine(_directory, "messages.ndjson");
        var database = Path.Combine(_directory, "messages.db");
        await File.WriteAllTextAsync(input,
            """
            {"FolderPath":"SamplePst\\TeamsMessagesData","SenderName":"Torey Page","SenderAddress":"torey@example.com","SenderDisplay":"Torey Page","Participants":"Linda Artley||Torey Page","ConversationKey":"chat-1","ConversationTitle":"Sample Teams chat","Subject":"Sample Teams chat","SentUtc":"2024-01-01T09:00:00Z","ReceivedUtc":"2024-01-01T09:00:05Z","Preview":"Hello Linda.","BodyText":"Hello Linda. phoenix","MessageClass":"IPM.Note","EntryId":"entry-1","AttachmentsText":""}
            {"FolderPath":"SamplePst\\TeamsMessagesData","SenderName":"Linda Artley","SenderAddress":"linda@example.com","SenderDisplay":"Linda Artley","Participants":"Linda Artley||Torey Page","ConversationKey":"chat-1","ConversationTitle":"Sample Teams chat","Subject":"Sample Teams chat","SentUtc":"2024-01-01T09:01:00Z","ReceivedUtc":"2024-01-01T09:01:05Z","Preview":"Thanks.","BodyText":"Thanks.","MessageClass":"IPM.Note","EntryId":"entry-2","AttachmentsText":"sample.pdf"}
            """ + Environment.NewLine, new UTF8Encoding(false));

        var result = await TeamsNdjsonImporter.ImportAsync(input, database, expectedCount: 2);
        var repository = new TeamsRepository(database);
        var conversations = await repository.SearchConversationsAsync(new TeamsQuery(Keyword: "phoenix"));
        var messages = await repository.GetConversationMessagesAsync("chat-1");

        Assert.Equal(2, result.ImportedCount);
        Assert.Equal(1, conversations.TotalCount);
        Assert.Equal(2, messages.Count);
        Assert.Equal("sample.pdf", messages[1].AttachmentsText);
    }

    [Fact]
    public async Task Expected_count_mismatch_fails_atomically()
    {
        Directory.CreateDirectory(_directory);
        var input = Path.Combine(_directory, "count.ndjson");
        var database = Path.Combine(_directory, "count.db");
        await File.WriteAllTextAsync(input,
            """{"FolderPath":"F","SenderName":"A","SenderAddress":"a@example.com","SenderDisplay":"A","Participants":"A","ConversationKey":"k","ConversationTitle":"T","Subject":"T","SentUtc":"2024-01-01T09:00:00Z","ReceivedUtc":"2024-01-01T09:00:05Z","Preview":"P","BodyText":"B","MessageClass":"IPM.Note","EntryId":"e1","AttachmentsText":""}"""
            + Environment.NewLine);

        await Assert.ThrowsAsync<InvalidDataException>(
            () => TeamsNdjsonImporter.ImportAsync(input, database, expectedCount: 2));

        Assert.False(File.Exists(database));
        Assert.False(File.Exists(database + ".importing"));
    }

    public void Dispose()
    {
        SqliteConnection.ClearAllPools();
        if (Directory.Exists(_directory))
            Directory.Delete(_directory, true);
    }
}

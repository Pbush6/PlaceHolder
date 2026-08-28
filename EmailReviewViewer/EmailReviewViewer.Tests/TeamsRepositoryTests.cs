using EmailReviewViewer.App.Data;
using Microsoft.Data.Sqlite;

namespace EmailReviewViewer.Tests;

public sealed class TeamsRepositoryTests : IDisposable
{
    private readonly string _databasePath = Path.Combine(Path.GetTempPath(), $"teams-viewer-{Guid.NewGuid():N}.db");

    [Fact]
    public async Task Conversation_page_groups_messages_and_does_not_load_bodies()
    {
        using var repository = await CreateSeededRepositoryAsync();
        var page = await repository.SearchConversationsAsync(new TeamsQuery(Limit: 50));

        Assert.Equal(2, page.TotalCount);
        Assert.Equal(2, page.Items.Count);
        var chat = Assert.Single(page.Items, item => item.ConversationTitle == "Sample Teams chat");
        Assert.Equal(2, chat.MessageCount);
        Assert.Equal(2, chat.TotalMessageCount);
        Assert.Contains("Linda Artley", chat.Participants);
        Assert.Contains("Torey Page", chat.Participants);
        Assert.False(string.IsNullOrWhiteSpace(chat.ConversationKey));
    }

    [Fact]
    public async Task Load_conversation_returns_chronological_messages_for_one_thread()
    {
        using var repository = await CreateSeededRepositoryAsync();
        var page = await repository.SearchConversationsAsync(new TeamsQuery(Limit: 50));
        var chat = Assert.Single(page.Items, item => item.ConversationTitle == "Sample Teams chat");

        var messages = await repository.GetConversationMessagesAsync(chat.ConversationKey);

        Assert.Equal(2, messages.Count);
        Assert.Equal("Torey Page", messages[0].SenderDisplay);
        Assert.Equal("Hello Linda.", messages[0].BodyText);
        Assert.Equal("Linda Artley", messages[1].SenderDisplay);
        Assert.Contains("sample.pdf", messages[1].AttachmentsText);
    }

    [Fact]
    public async Task Keyword_filter_matches_fts_body_without_returning_other_conversations()
    {
        using var repository = await CreateSeededRepositoryAsync();
        var page = await repository.SearchConversationsAsync(new TeamsQuery(Keyword: "phoenix", Limit: 50));

        Assert.Equal(1, page.TotalCount);
        Assert.Equal("Budget chat", Assert.Single(page.Items).ConversationTitle);
    }

    [Fact]
    public async Task Messages_from_selected_people_hides_other_senders_but_keeps_the_conversation()
    {
        using var repository = await CreateSeededRepositoryAsync();
        var page = await repository.SearchConversationsAsync(new TeamsQuery(
            SelectedParticipants: ["Torey Page"],
            MatchMode: TeamsParticipantMatchMode.MessagesFromSelected,
            Limit: 50));

        var chat = Assert.Single(page.Items, item => item.ConversationTitle == "Sample Teams chat");
        Assert.Equal(1, chat.MessageCount);
        Assert.Equal(2, chat.TotalMessageCount);
        Assert.Equal(1, page.TotalCount);
        Assert.Equal(1, page.VisibleMessageCount);
    }

    [Fact]
    public async Task Exact_selected_people_only_keeps_conversations_with_the_same_participant_set()
    {
        using var repository = await CreateSeededRepositoryAsync();
        var page = await repository.SearchConversationsAsync(new TeamsQuery(
            SelectedParticipants: ["Linda Artley", "Torey Page"],
            MatchMode: TeamsParticipantMatchMode.ExactAll,
            Limit: 50));

        Assert.Equal(2, page.TotalCount);
    }

    [Fact]
    public void Likely_person_names_match_the_Teams_HTML_rules()
    {
        Assert.True(TeamsParticipantMatching.IsLikelyPersonName("Linda Artley"));
        Assert.False(TeamsParticipantMatching.IsLikelyPersonName("Fireflies.ai Notetaker"));
        Assert.False(TeamsParticipantMatching.IsLikelyPersonName("meeting-bot"));
    }

    [Fact]
    public async Task Folder_counts_group_messages()
    {
        using var repository = await CreateSeededRepositoryAsync();
        var folders = await repository.GetFolderCountsAsync();

        Assert.Equal(
        [
            new FolderCount(@"SamplePst\Inbox", 1),
            new FolderCount(@"SamplePst\TeamsMessagesData", 2)
        ], folders);
    }

    [Fact]
    public async Task Count_messages_returns_row_count_for_import_validation()
    {
        using var repository = await CreateSeededRepositoryAsync();
        Assert.Equal(3, await repository.CountMessagesAsync());
    }

    private async Task<TeamsRepository> CreateSeededRepositoryAsync()
    {
        var repository = new TeamsRepository(_databasePath);
        await repository.EnsureCreatedAsync();
        await repository.InsertBatchAsync(
        [
            Message(1, "chat-1", "Sample Teams chat", "Torey Page", "Hello Linda.",
                @"SamplePst\TeamsMessagesData", "Linda Artley||Torey Page", new DateTime(2024, 1, 1, 9, 0, 0, DateTimeKind.Utc), ""),
            Message(2, "chat-1", "Sample Teams chat", "Linda Artley", "Thanks.",
                @"SamplePst\TeamsMessagesData", "Linda Artley||Torey Page", new DateTime(2024, 1, 1, 9, 1, 0, DateTimeKind.Utc), "sample.pdf"),
            Message(3, "chat-2", "Budget chat", "Linda Artley", "The phoenix rollout is funded.",
                @"SamplePst\Inbox", "Linda Artley||Torey Page", new DateTime(2024, 1, 1, 10, 0, 0, DateTimeKind.Utc), "")
        ]);
        return repository;
    }

    private static TeamsMessage Message(
        long id,
        string conversationKey,
        string conversationTitle,
        string senderDisplay,
        string body,
        string folderPath,
        string participants,
        DateTime receivedUtc,
        string attachments) => new()
        {
            Id = id,
            FolderPath = folderPath,
            SenderName = senderDisplay,
            SenderAddress = senderDisplay.Replace(" ", ".").ToLowerInvariant() + "@example.com",
            SenderDisplay = senderDisplay,
            Participants = participants,
            ConversationKey = conversationKey,
            ConversationTitle = conversationTitle,
            Subject = conversationTitle,
            SentUtc = receivedUtc.AddSeconds(-5),
            ReceivedUtc = receivedUtc,
            Preview = body,
            BodyText = body,
            MessageClass = "IPM.Note",
            EntryId = $"entry-{id}",
            AttachmentsText = attachments
        };

    public void Dispose()
    {
        SqliteConnection.ClearAllPools();
        foreach (var candidate in new[] { _databasePath, _databasePath + "-wal", _databasePath + "-shm" })
            if (File.Exists(candidate))
                File.Delete(candidate);
    }
}

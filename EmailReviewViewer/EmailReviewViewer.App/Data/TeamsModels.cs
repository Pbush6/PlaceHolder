namespace EmailReviewViewer.App.Data;

public sealed class TeamsMessage
{
    public long Id { get; init; }
    public string FolderPath { get; init; } = "";
    public string SenderName { get; init; } = "";
    public string SenderAddress { get; init; } = "";
    public string SenderDisplay { get; init; } = "";
    public string Participants { get; init; } = "";
    public string ConversationKey { get; init; } = "";
    public string ConversationTitle { get; init; } = "";
    public string Subject { get; init; } = "";
    public DateTime? SentUtc { get; init; }
    public DateTime? ReceivedUtc { get; init; }
    public string Preview { get; init; } = "";
    public string BodyText { get; init; } = "";
    public string MessageClass { get; init; } = "";
    public string EntryId { get; init; } = "";
    public string AttachmentsText { get; init; } = "";
    public string ToRecipients { get; init; } = "";
    public string CcRecipients { get; init; } = "";
}

public sealed record TeamsConversationListItem(
    string ConversationKey,
    DateTime? LastUtc,
    string ConversationTitle,
    string Participants,
    long MessageCount,
    long TotalMessageCount);

public sealed record TeamsQuery(
    DateTime? FromUtc = null,
    DateTime? ToUtc = null,
    string? Party = null,
    string? Keyword = null,
    int Limit = 200,
    int Offset = 0,
    IReadOnlyList<string>? FolderPaths = null,
    IReadOnlyList<string>? SelectedParticipants = null,
    TeamsParticipantMatchMode MatchMode = TeamsParticipantMatchMode.MessagesFromSelected,
    bool NewestFirst = true);

public sealed record TeamsConversationPage(
    IReadOnlyList<TeamsConversationListItem> Items,
    long TotalCount,
    long VisibleMessageCount);

public enum TeamsParticipantMatchMode
{
    MessagesFromSelected,
    Involving,
    ExactAll,
    ExactPeopleOnly
}

public sealed record TeamsReportSummary(
    long MessageCount,
    int PeopleCount,
    int FolderCount,
    IReadOnlyList<string> People,
    IReadOnlyList<string> OtherNames);

public sealed record TeamsImportResult(int InputCount, int ImportedCount, string DatabasePath);

public enum ReportDatabaseKind
{
    Email,
    Teams
}

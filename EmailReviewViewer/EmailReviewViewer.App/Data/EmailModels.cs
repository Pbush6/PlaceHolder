namespace EmailReviewViewer.App.Data;

public sealed class EmailMessage
{
    public long Id { get; init; }
    public string FolderPath { get; init; } = "";
    public string SenderName { get; init; } = "";
    public string SenderAddress { get; init; } = "";
    public string ToRecipients { get; init; } = "";
    public string CcRecipients { get; init; } = "";
    public string Subject { get; init; } = "";
    public DateTime? SentUtc { get; init; }
    public DateTime? ReceivedUtc { get; init; }
    public string Preview { get; init; } = "";
    public string BodyText { get; init; } = "";
    public string MessageClass { get; init; } = "";
    public string EntryId { get; init; } = "";
    public string ConversationId { get; init; } = "";
    public string ConversationTopic { get; init; } = "";
}

public sealed record EmailListItem(
    long Id,
    DateTime? DateUtc,
    string SenderName,
    string SenderAddress,
    string ToRecipients,
    string Subject,
    string Preview);

public sealed record EmailQuery(
    DateTime? FromUtc = null,
    DateTime? ToUtc = null,
    string? Party = null,
    string? Keyword = null,
    int Limit = 200,
    int Offset = 0,
    IReadOnlyList<string>? FolderPaths = null);

public sealed record EmailPage(IReadOnlyList<EmailListItem> Items, long TotalCount);

public sealed record FolderCount(string FolderPath, long Count);

public sealed record EmailImportResult(int InputCount, int ImportedCount, string DatabasePath);

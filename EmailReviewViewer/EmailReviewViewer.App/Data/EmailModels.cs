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

public enum EmailSortColumn
{
    Date,
    From,
    To,
    Subject
}

public enum EmailSortDirection
{
    Ascending,
    Descending
}

public static class EmailSortSql
{
    public static bool TryMapColumnKey(string? key, out EmailSortColumn column)
    {
        switch (key)
        {
            case nameof(EmailListItem.DateUtc):
                column = EmailSortColumn.Date;
                return true;
            case nameof(EmailListItem.SenderName):
                column = EmailSortColumn.From;
                return true;
            case nameof(EmailListItem.ToRecipients):
                column = EmailSortColumn.To;
                return true;
            case nameof(EmailListItem.Subject):
                column = EmailSortColumn.Subject;
                return true;
            default:
                column = default;
                return false;
        }
    }

    public static string BuildOrderBy(EmailSortColumn column, EmailSortDirection direction)
    {
        var expression = column switch
        {
            EmailSortColumn.Date => "COALESCE(m.ReceivedUtc, m.SentUtc)",
            EmailSortColumn.From => "m.SenderName COLLATE NOCASE",
            EmailSortColumn.To => "m.ToRecipients COLLATE NOCASE",
            EmailSortColumn.Subject => "m.Subject COLLATE NOCASE",
            _ => throw new ArgumentOutOfRangeException(nameof(column), column, "Unsupported email sort column.")
        };
        var keyword = direction switch
        {
            EmailSortDirection.Ascending => "ASC",
            EmailSortDirection.Descending => "DESC",
            _ => throw new ArgumentOutOfRangeException(nameof(direction), direction, "Unsupported sort direction.")
        };
        return $"{expression} {keyword}, m.Id ASC";
    }
}

public sealed record EmailQuery(
    DateTime? FromUtc = null,
    DateTime? ToUtc = null,
    string? Party = null,
    string? Keyword = null,
    int Limit = 200,
    int Offset = 0,
    IReadOnlyList<string>? FolderPaths = null,
    EmailSortColumn SortColumn = EmailSortColumn.Date,
    EmailSortDirection SortDirection = EmailSortDirection.Descending);

public sealed record EmailPage(IReadOnlyList<EmailListItem> Items, long TotalCount);

public sealed record FolderCount(string FolderPath, long Count);

public sealed record EmailImportResult(int InputCount, int ImportedCount, string DatabasePath);

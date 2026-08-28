namespace EmailReviewViewer.App;

internal static class ReportViewerSession
{
    public static string? RequestedPath { get; set; }

    public static void Request(string databasePath) =>
        RequestedPath = Path.GetFullPath(databasePath);
}

using Microsoft.Data.Sqlite;

namespace EmailReviewViewer.App.Data;

public static class ReportDatabaseKindDetector
{
    public static async Task<ReportDatabaseKind> DetectAsync(
        string databasePath,
        CancellationToken cancellationToken = default)
    {
        var fullPath = Path.GetFullPath(databasePath);
        if (!File.Exists(fullPath))
            throw new InvalidDataException($"Database file not found:{Environment.NewLine}{fullPath}");

        var connectionString = new SqliteConnectionStringBuilder
        {
            DataSource = fullPath,
            Mode = SqliteOpenMode.ReadOnly
        }.ToString();
        await using var connection = new SqliteConnection(connectionString);
        await connection.OpenAsync(cancellationToken);
        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT name
            FROM sqlite_master
            WHERE name IN ('EmailMessages', 'TeamsMessages');
            """;
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
            names.Add(reader.GetString(0));

        if (names.Contains("TeamsMessages"))
            return ReportDatabaseKind.Teams;
        if (names.Contains("EmailMessages"))
            return ReportDatabaseKind.Email;
        throw new InvalidDataException(
            "The selected file is not a compatible Email Review Viewer database.");
    }
}

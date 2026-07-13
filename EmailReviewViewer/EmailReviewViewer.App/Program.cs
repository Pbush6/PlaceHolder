using System.Diagnostics;
using System.Text.Json;
using EmailReviewViewer.App.Data;
using Microsoft.Data.Sqlite;

namespace EmailReviewViewer.App;

internal static class Program
{
    [STAThread]
    private static async Task<int> Main(string[] args)
    {
        try
        {
            if (args.FirstOrDefault()?.Equals("--import", StringComparison.OrdinalIgnoreCase) == true)
                return await ImportAsync(args);
            if (args.FirstOrDefault()?.Equals("--inspect", StringComparison.OrdinalIgnoreCase) == true)
                return await InspectAsync(args);
            if (args.FirstOrDefault()?.Equals("--generate", StringComparison.OrdinalIgnoreCase) == true)
                return await GenerateAsync(args);
            if (args.FirstOrDefault()?.Equals("--benchmark", StringComparison.OrdinalIgnoreCase) == true)
                return await BenchmarkAsync(args);

            var smoke = args.FirstOrDefault()?.Equals("--smoke-ui", StringComparison.OrdinalIgnoreCase) == true;
            string? databasePath = smoke
                ? GetPath(args, 1)
                : args.Length > 0 ? Path.GetFullPath(args[0]) : null;

            ApplicationConfiguration.Initialize();
            using var form = new MainForm(databasePath);
            if (smoke)
            {
                var timer = new System.Windows.Forms.Timer { Interval = 2000 };
                timer.Tick += (_, _) =>
                {
                    timer.Stop();
                    form.Close();
                };
                form.Shown += (_, _) => timer.Start();
            }
            Application.Run(form);
            return 0;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    private static async Task<int> ImportAsync(string[] args)
    {
        var inputPath = GetPath(args, 1);
        var databasePath = GetOptionPath(args, "--database");
        var expectedText = GetOption(args, "--expected-count");
        int? expectedCount = expectedText is null
            ? null
            : int.TryParse(expectedText, out var parsed) && parsed >= 0
                ? parsed
                : throw new ArgumentException("--expected-count must be a non-negative integer.");
        var result = await EmailNdjsonImporter.ImportAsync(inputPath, databasePath, expectedCount);
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            mode = "import",
            inputPath,
            databasePath = result.DatabasePath,
            inputCount = result.InputCount,
            importedCount = result.ImportedCount
        }));
        return 0;
    }

    private static async Task<int> InspectAsync(string[] args)
    {
        var databasePath = GetPath(args, 1);
        var keyword = args.Length > 2 ? args[2] : "phoenix";
        var repository = new EmailRepository(databasePath);
        var all = await repository.SearchPageAsync(new EmailQuery(Limit: 1));
        var folders = await repository.GetFolderCountsAsync();
        var matches = await repository.SearchPageAsync(new EmailQuery(Keyword: keyword, Limit: 1));
        var detail = matches.Items.Count > 0
            ? await repository.GetByIdAsync(matches.Items[0].Id)
            : null;
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            mode = "inspect",
            databasePath,
            rowCount = all.TotalCount,
            folderCount = folders.Count,
            folders,
            keyword,
            keywordHits = matches.TotalCount,
            detailBody = detail?.BodyText
        }, JsonOptions));
        return 0;
    }

    private static async Task<int> GenerateAsync(string[] args)
    {
        var databasePath = GetPath(args, 1);
        var count = args.Length > 2 && int.TryParse(args[2], out var parsed) ? parsed : 25_000;
        var stopwatch = Stopwatch.StartNew();
        await SampleDatabaseGenerator.GenerateAsync(databasePath, count);
        SqliteConnection.ClearAllPools();
        stopwatch.Stop();
        var size = new FileInfo(databasePath).Length;
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            mode = "generate",
            databasePath,
            count,
            durationMs = stopwatch.ElapsedMilliseconds,
            sizeBytes = size,
            sizeMiB = Math.Round(size / 1024d / 1024d, 2)
        }, JsonOptions));
        return 0;
    }

    private static async Task<int> BenchmarkAsync(string[] args)
    {
        var databasePath = GetPath(args, 1);
        if (!File.Exists(databasePath))
            throw new FileNotFoundException("Database not found.", databasePath);

        var repository = new EmailRepository(databasePath);
        await repository.EnsureCreatedAsync();
        var benchmarks = new[]
        {
            ("keyword", new EmailQuery(Keyword: "phoenix", Limit: 50)),
            ("date-sort", new EmailQuery(
                FromUtc: new DateTime(2025, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                ToUtc: new DateTime(2025, 12, 31, 23, 59, 59, DateTimeKind.Utc),
                Limit: 50,
                SortColumn: EmailSortColumn.Date,
                SortDirection: EmailSortDirection.Descending)),
            ("subject-sort", new EmailQuery(
                Limit: 50,
                SortColumn: EmailSortColumn.Subject,
                SortDirection: EmailSortDirection.Ascending)),
            ("sender", new EmailQuery(Party: "alice.morgan@example.com", Limit: 50)),
            ("folder", new EmailQuery(FolderPaths: [@"Mailbox\Projects\Phoenix"], Limit: 50))
        };
        var results = new List<object>();
        foreach (var (name, query) in benchmarks)
        {
            var stopwatch = Stopwatch.StartNew();
            var page = await repository.SearchPageAsync(query);
            stopwatch.Stop();
            if (page.Items.Count > query.Limit)
                throw new InvalidOperationException("Paging limit was not honored.");
            results.Add(new
            {
                filter = name,
                durationMs = stopwatch.Elapsed.TotalMilliseconds,
                totalMatches = page.TotalCount,
                rowsLoaded = page.Items.Count,
                pageLimit = query.Limit
            });
        }

        Console.WriteLine(JsonSerializer.Serialize(new
        {
            mode = "benchmark",
            databasePath,
            bodyLoadedInListQuery = false,
            results
        }, JsonOptions));
        return 0;
    }

    private static string GetPath(string[] args, int index)
    {
        if (args.Length <= index || string.IsNullOrWhiteSpace(args[index]))
            throw new ArgumentException("A database path is required.");
        return Path.GetFullPath(args[index]);
    }

    private static string GetOptionPath(string[] args, string name) =>
        Path.GetFullPath(GetOption(args, name)
            ?? throw new ArgumentException($"{name} requires a path."));

    private static string? GetOption(string[] args, string name)
    {
        var index = Array.FindIndex(args,
            argument => argument.Equals(name, StringComparison.OrdinalIgnoreCase));
        return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
    }

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
}
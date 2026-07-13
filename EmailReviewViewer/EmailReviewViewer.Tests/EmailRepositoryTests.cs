using EmailReviewViewer.App.Data;
using Microsoft.Data.Sqlite;

namespace EmailReviewViewer.Tests;

public sealed class EmailRepositoryTests : IDisposable
{
    private readonly string _databasePath = Path.Combine(Path.GetTempPath(), $"email-viewer-{Guid.NewGuid():N}.db");

    [Fact]
    public void MainForm_constructs_at_its_initial_layout_size()
    {
        using var form = new EmailReviewViewer.App.MainForm(_databasePath);

        Assert.True(form.Width >= 1000);
        Assert.True(form.Height >= 650);
    }

    [Fact]
    public void MainForm_without_database_has_open_database_commands_and_welcome_state()
    {
        using var form = new EmailReviewViewer.App.MainForm(null);
        form.CreateControl();

        var openButton = Assert.IsType<System.Windows.Forms.Button>(
            form.Controls.Find("OpenDatabaseButton", true).Single());
        var openMenu = Assert.IsType<System.Windows.Forms.ToolStripMenuItem>(
            form.MainMenuStrip!.Items[0]).DropDownItems
            .OfType<System.Windows.Forms.ToolStripMenuItem>()
            .Single(item => item.Name == "OpenDatabaseMenuItem");
        var subject = Assert.IsType<System.Windows.Forms.Label>(
            form.Controls.Find("ReadingSubject", true).Single());

        Assert.True(openButton.Enabled);
        Assert.Equal("Open Database…", openMenu.Text);
        Assert.Contains("Open", subject.Text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MainForm_date_filters_reserve_localized_text_and_native_buttons()
    {
        using var form = new EmailReviewViewer.App.MainForm(_databasePath);
        form.CreateControl();
        form.Size = form.MinimumSize;
        form.PerformLayout();

        var from = Assert.IsType<System.Windows.Forms.DateTimePicker>(
            form.Controls.Find("FromDate", true).Single());
        var to = Assert.IsType<System.Windows.Forms.DateTimePicker>(
            form.Controls.Find("ToDate", true).Single());
        var search = Assert.IsType<System.Windows.Forms.Button>(
            form.Controls.Find("SearchButton", true).Single());
        var widestDate = Enumerable.Range(1, 12)
            .Select(month => new DateTime(DateTime.Today.Year, month, 28).ToString("d"))
            .Max(text => System.Windows.Forms.TextRenderer.MeasureText(text, from.Font).Width);
        var requiredWidth =
            widestDate + (System.Windows.Forms.SystemInformation.VerticalScrollBarWidth * 2);

        Assert.True(from.MinimumSize.Width >= requiredWidth);
        Assert.Equal(from.MinimumSize.Width, to.MinimumSize.Width);
        Assert.True(search.Right <= search.Parent!.ClientSize.Width);
    }

    [Fact]
    public void MainForm_has_resizable_three_area_layout_and_empty_reading_state()
    {
        using var form = new EmailReviewViewer.App.MainForm(_databasePath);
        form.CreateControl();
        form.Size = form.MinimumSize;
        form.PerformLayout();

        var folderResultsSplit = Assert.IsType<System.Windows.Forms.SplitContainer>(
            form.Controls.Find("FolderResultsSplit", true).Single());
        var messageReadingSplit = Assert.IsType<System.Windows.Forms.SplitContainer>(
            form.Controls.Find("MessageReadingSplit", true).Single());
        var subject = Assert.IsType<System.Windows.Forms.Label>(
            form.Controls.Find("ReadingSubject", true).Single());
        var body = Assert.IsType<System.Windows.Forms.RichTextBox>(
            form.Controls.Find("ReadingBody", true).Single());

        Assert.Equal(System.Windows.Forms.Orientation.Vertical, folderResultsSplit.Orientation);
        Assert.Equal(System.Windows.Forms.Orientation.Vertical, messageReadingSplit.Orientation);
        Assert.True(folderResultsSplit.Panel1MinSize > 0);
        Assert.True(messageReadingSplit.Panel1MinSize > 0);
        Assert.True(messageReadingSplit.Panel2MinSize > 0);
        Assert.True(
            subject.Text.Contains("Select an email", StringComparison.OrdinalIgnoreCase) ||
            subject.Text.Contains("Open a database", StringComparison.OrdinalIgnoreCase));
        Assert.True(body.ReadOnly);
        Assert.Equal(System.Windows.Forms.DockStyle.Fill, body.Dock);
        Assert.True(form.MinimumSize.Width >= 1200);
    }

    [Fact]
    public void MainForm_grid_headers_are_distinct_and_only_meaningful_columns_are_sortable()
    {
        using var form = new EmailReviewViewer.App.MainForm(_databasePath);
        form.CreateControl();
        var grid = Assert.IsType<System.Windows.Forms.DataGridView>(
            form.Controls.Find("EmailGrid", true).Single());

        Assert.False(grid.EnableHeadersVisualStyles);
        Assert.True(grid.ColumnHeadersDefaultCellStyle.Font!.Bold);
        Assert.NotEqual(grid.DefaultCellStyle.BackColor, grid.ColumnHeadersDefaultCellStyle.BackColor);
        Assert.All(
            grid.Columns.Cast<System.Windows.Forms.DataGridViewColumn>().Take(4),
            column => Assert.Equal(System.Windows.Forms.DataGridViewColumnSortMode.Programmatic, column.SortMode));
        Assert.Equal(
            System.Windows.Forms.DataGridViewColumnSortMode.NotSortable,
            grid.Columns[nameof(EmailListItem.Preview)].SortMode);
        Assert.Equal(
            System.Windows.Forms.SortOrder.Descending,
            grid.Columns[nameof(EmailListItem.DateUtc)].HeaderCell.SortGlyphDirection);
    }

    [Fact]
    public void Enterprise_theme_uses_restrained_palette_and_system_high_contrast_fallbacks()
    {
        var standard = EmailReviewViewer.App.AppTheme.Resolve(highContrast: false);
        var accessible = EmailReviewViewer.App.AppTheme.Resolve(highContrast: true);

        Assert.Equal(System.Drawing.Color.FromArgb(27, 48, 74), standard.Primary);
        Assert.Equal(System.Drawing.Color.FromArgb(45, 105, 166), standard.Accent);
        Assert.NotEqual(standard.Surface, standard.Canvas);
        Assert.Equal(System.Drawing.SystemColors.Window, accessible.Surface);
        Assert.Equal(System.Drawing.SystemColors.Highlight, accessible.Accent);
        Assert.Equal(System.Drawing.SystemColors.WindowText, accessible.Text);
    }

    [Fact]
    public void MainForm_exposes_professional_header_filter_reset_and_section_hierarchy()
    {
        using var form = new EmailReviewViewer.App.MainForm(null);
        form.CreateControl();
        form.Size = form.MinimumSize;
        form.PerformLayout();

        var title = Assert.IsType<System.Windows.Forms.Label>(
            form.Controls.Find("ApplicationTitle", true).Single());
        var database = Assert.IsType<System.Windows.Forms.Label>(
            form.Controls.Find("HeaderDatabasePath", true).Single());
        var reset = Assert.IsType<System.Windows.Forms.Button>(
            form.Controls.Find("ResetFiltersButton", true).Single());
        var folderHeading = Assert.IsType<System.Windows.Forms.Label>(
            form.Controls.Find("FolderHeading", true).Single());
        var resultsHeading = Assert.IsType<System.Windows.Forms.Label>(
            form.Controls.Find("ResultsHeading", true).Single());

        Assert.Equal("Email Review Viewer", title.Text);
        Assert.Contains("database", database.Text, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("Reset filters", reset.Text);
        Assert.Equal("Folders", folderHeading.Text);
        Assert.Equal("Messages", resultsHeading.Text);
        Assert.True(reset.Right <= reset.Parent!.ClientSize.Width);
    }

    [Theory]
    [InlineData(nameof(EmailListItem.DateUtc), EmailSortColumn.Date)]
    [InlineData(nameof(EmailListItem.SenderName), EmailSortColumn.From)]
    [InlineData(nameof(EmailListItem.ToRecipients), EmailSortColumn.To)]
    [InlineData(nameof(EmailListItem.Subject), EmailSortColumn.Subject)]
    public void Sort_mapping_allowlists_only_supported_grid_columns(string key, EmailSortColumn expected)
    {
        Assert.True(EmailSortSql.TryMapColumnKey(key, out var actual));
        Assert.Equal(expected, actual);
        Assert.False(EmailSortSql.TryMapColumnKey(nameof(EmailListItem.Preview), out _));
        Assert.False(EmailSortSql.TryMapColumnKey("m.Id DESC; DROP TABLE EmailMessages", out _));
    }

    [Theory]
    [InlineData(EmailSortColumn.Date, EmailSortDirection.Descending, "COALESCE(m.ReceivedUtc, m.SentUtc) DESC, m.Id ASC")]
    [InlineData(EmailSortColumn.From, EmailSortDirection.Ascending, "m.SenderName COLLATE NOCASE ASC, m.Id ASC")]
    [InlineData(EmailSortColumn.To, EmailSortDirection.Descending, "m.ToRecipients COLLATE NOCASE DESC, m.Id ASC")]
    [InlineData(EmailSortColumn.Subject, EmailSortDirection.Ascending, "m.Subject COLLATE NOCASE ASC, m.Id ASC")]
    public void Sort_sql_uses_allowlisted_expression_direction_and_stable_tie_breaker(
        EmailSortColumn column,
        EmailSortDirection direction,
        string expected)
    {
        Assert.Equal(expected, EmailSortSql.BuildOrderBy(column, direction));
    }

    [Fact]
    public async Task EnsureCreated_builds_message_and_fts_tables()
    {
        var repository = new EmailRepository(_databasePath);

        await repository.EnsureCreatedAsync();

        await using var connection = new SqliteConnection($"Data Source={_databasePath}");
        await connection.OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT name
            FROM sqlite_master
            WHERE type = 'table' AND name IN ('EmailMessages', 'EmailMessagesFts')
            ORDER BY name;
            """;
        await using var reader = await command.ExecuteReaderAsync();
        var names = new List<string>();
        while (await reader.ReadAsync())
            names.Add(reader.GetString(0));

        Assert.Equal(["EmailMessages", "EmailMessagesFts"], names);
    }

    [Fact]
    public async Task SearchPage_uses_fts_for_subject_body_and_participants()
    {
        var repository = await CreateSeededRepositoryAsync();

        var subject = await repository.SearchPageAsync(new EmailQuery(Keyword: "budget"));
        var body = await repository.SearchPageAsync(new EmailQuery(Keyword: "phoenix"));
        var participant = await repository.SearchPageAsync(new EmailQuery(Keyword: "alice"));

        Assert.Single(subject.Items);
        Assert.Equal("Quarterly budget", subject.Items[0].Subject);
        Assert.Single(body.Items);
        Assert.Equal("Project status", body.Items[0].Subject);
        Assert.Equal(2, participant.Items.Count);
        Assert.Contains(participant.Items, item => item.SenderAddress == "alice@example.com");
        Assert.Contains(participant.Items, item => item.ToRecipients == "alice@example.com");
    }

    [Fact]
    public async Task SearchPage_applies_date_and_party_filters()
    {
        var repository = await CreateSeededRepositoryAsync();

        var result = await repository.SearchPageAsync(new EmailQuery(
            FromUtc: new DateTime(2026, 1, 2, 0, 0, 0, DateTimeKind.Utc),
            ToUtc: new DateTime(2026, 1, 2, 23, 59, 59, DateTimeKind.Utc),
            Party: "bob@example.com"));

        Assert.Single(result.Items);
        Assert.Equal("Project status", result.Items[0].Subject);
    }

    [Fact]
    public async Task SearchPage_returns_only_requested_page_without_bodies()
    {
        var repository = await CreateSeededRepositoryAsync();

        var result = await repository.SearchPageAsync(new EmailQuery(Limit: 1, Offset: 1));

        Assert.Equal(3, result.TotalCount);
        Assert.Single(result.Items);
        Assert.Equal("Project status", result.Items[0].Subject);
        Assert.Null(typeof(EmailListItem).GetProperty("BodyText"));
    }

    [Fact]
    public async Task SearchPage_honors_ascending_and_descending_text_sort()
    {
        var repository = await CreateSeededRepositoryAsync();

        var ascending = await repository.SearchPageAsync(new EmailQuery(
            SortColumn: EmailSortColumn.Subject,
            SortDirection: EmailSortDirection.Ascending));
        var descending = await repository.SearchPageAsync(new EmailQuery(
            SortColumn: EmailSortColumn.Subject,
            SortDirection: EmailSortDirection.Descending));

        Assert.Equal(["Project status", "Quarterly budget", "Welcome"], ascending.Items.Select(item => item.Subject));
        Assert.Equal(["Welcome", "Quarterly budget", "Project status"], descending.Items.Select(item => item.Subject));
    }

    [Fact]
    public async Task SearchPage_uses_id_as_a_stable_sort_tie_breaker()
    {
        var repository = await CreateSeededRepositoryAsync();
        await repository.InsertBatchAsync(
        [
            Message(5, "Same subject", "Fifth.", "Zed", "zed@example.com",
                "reviewer@example.com", new DateTime(2026, 1, 4, 9, 0, 0, DateTimeKind.Utc)),
            Message(4, "Same subject", "Fourth.", "Zed", "zed@example.com",
                "reviewer@example.com", new DateTime(2026, 1, 5, 9, 0, 0, DateTimeKind.Utc))
        ]);

        var result = await repository.SearchPageAsync(new EmailQuery(
            Keyword: "\"Same subject\"",
            SortColumn: EmailSortColumn.Subject,
            SortDirection: EmailSortDirection.Ascending));

        Assert.Equal([4L, 5L], result.Items.Select(item => item.Id));
    }

    [Fact]
    public async Task SearchPage_preserves_filters_and_paging_when_sorted()
    {
        var repository = await CreateSeededRepositoryAsync();

        var result = await repository.SearchPageAsync(new EmailQuery(
            Party: "alice",
            Limit: 1,
            Offset: 1,
            FolderPaths: [@"Mailbox\Inbox"],
            SortColumn: EmailSortColumn.Subject,
            SortDirection: EmailSortDirection.Ascending));

        Assert.Equal(2, result.TotalCount);
        Assert.Single(result.Items);
        Assert.Equal("Quarterly budget", result.Items[0].Subject);
    }

    [Fact]
    public async Task SearchPage_filters_one_or_multiple_full_folder_paths()
    {
        var repository = await CreateSeededRepositoryAsync();

        var single = await repository.SearchPageAsync(new EmailQuery(
            FolderPaths: [@"Mailbox\Inbox"]));
        var multiple = await repository.SearchPageAsync(new EmailQuery(
            FolderPaths: [@"Mailbox\Inbox", @"Mailbox\Archive\Inbox"]));

        Assert.Equal(2, single.TotalCount);
        Assert.All(single.Items, item => Assert.NotEqual("Welcome", item.Subject));
        Assert.Equal(3, multiple.TotalCount);
        Assert.Contains(multiple.Items, item => item.Subject == "Welcome");
    }

    [Fact]
    public async Task SearchPage_treats_null_or_empty_folder_selection_as_all_folders()
    {
        var repository = await CreateSeededRepositoryAsync();

        var all = await repository.SearchPageAsync(new EmailQuery());
        var empty = await repository.SearchPageAsync(new EmailQuery(FolderPaths: []));

        Assert.Equal(3, all.TotalCount);
        Assert.Equal(all.TotalCount, empty.TotalCount);
        Assert.Equal(all.Items.Select(item => item.Id), empty.Items.Select(item => item.Id));
    }

    [Fact]
    public async Task GetFolderCounts_groups_by_full_path_and_counts_messages()
    {
        var repository = await CreateSeededRepositoryAsync();

        var counts = await repository.GetFolderCountsAsync();

        Assert.Equal(
        [
            new FolderCount(@"Mailbox\Archive\Inbox", 1),
            new FolderCount(@"Mailbox\Inbox", 2)
        ], counts);
    }

    [Fact]
    public async Task GetById_returns_full_message_body()
    {
        var repository = await CreateSeededRepositoryAsync();
        var page = await repository.SearchPageAsync(new EmailQuery(Keyword: "phoenix"));

        var detail = await repository.GetByIdAsync(page.Items[0].Id);

        Assert.NotNull(detail);
        Assert.Contains("phoenix", detail.BodyText, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("entry-2", detail.EntryId);
    }

    [Fact]
    public async Task GetById_returns_null_when_message_no_longer_exists()
    {
        var repository = await CreateSeededRepositoryAsync();

        var detail = await repository.GetByIdAsync(999_999);

        Assert.Null(detail);
    }

    [Fact]
    public async Task Database_switcher_opens_valid_schema_and_replaces_current_database()
    {
        var first = await CreateSeededRepositoryAsync();
        var secondPath = Path.Combine(Path.GetTempPath(), $"email-viewer-switch-{Guid.NewGuid():N}.db");
        try
        {
            var second = new EmailRepository(secondPath);
            await second.EnsureCreatedAsync();
            await second.InsertBatchAsync(
            [
                Message(10, "Second database", "Loaded after switching.", "Dana", "dana@example.com",
                    "reviewer@example.com", DateTime.UtcNow)
            ]);
            first.Dispose();
            second.Dispose();

            using var switcher = new EmailDatabaseStoreSwitcher();
            await switcher.SwitchAsync(_databasePath);
            await switcher.SwitchAsync(secondPath);

            Assert.Equal(Path.GetFullPath(secondPath), switcher.DatabasePath);
            var result = await switcher.Repository!.SearchPageAsync(new EmailQuery());
            Assert.Single(result.Items);
            Assert.Equal("Second database", result.Items[0].Subject);
        }
        finally
        {
            SqliteConnection.ClearAllPools();
            if (File.Exists(secondPath))
                File.Delete(secondPath);
        }
    }

    [Fact]
    public async Task Database_switcher_rejects_wrong_schema_without_replacing_current_database()
    {
        var current = await CreateSeededRepositoryAsync();
        var wrongPath = Path.Combine(Path.GetTempPath(), $"email-viewer-wrong-{Guid.NewGuid():N}.db");
        try
        {
            await using (var connection = new SqliteConnection($"Data Source={wrongPath}"))
            {
                await connection.OpenAsync();
                await using var command = connection.CreateCommand();
                command.CommandText = "CREATE TABLE Unrelated (Id INTEGER PRIMARY KEY);";
                await command.ExecuteNonQueryAsync();
            }
            current.Dispose();

            using var switcher = new EmailDatabaseStoreSwitcher();
            await switcher.SwitchAsync(_databasePath);
            var error = await Assert.ThrowsAsync<InvalidDataException>(
                () => switcher.SwitchAsync(wrongPath));

            Assert.Contains("compatible", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(Path.GetFullPath(_databasePath), switcher.DatabasePath);
            Assert.NotNull(switcher.Repository);
        }
        finally
        {
            SqliteConnection.ClearAllPools();
            if (File.Exists(wrongPath))
                File.Delete(wrongPath);
        }
    }

    [Fact]
    public async Task Database_switcher_rejects_corrupt_file()
    {
        var corruptPath = Path.Combine(Path.GetTempPath(), $"email-viewer-corrupt-{Guid.NewGuid():N}.db");
        try
        {
            await File.WriteAllTextAsync(corruptPath, "not a sqlite database");
            using var switcher = new EmailDatabaseStoreSwitcher();

            var error = await Assert.ThrowsAsync<InvalidDataException>(
                () => switcher.SwitchAsync(corruptPath));

            Assert.Contains("valid", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Null(switcher.Repository);
        }
        finally
        {
            SqliteConnection.ClearAllPools();
            if (File.Exists(corruptPath))
                File.Delete(corruptPath);
        }
    }

    private async Task<EmailRepository> CreateSeededRepositoryAsync()
    {
        var repository = new EmailRepository(_databasePath);
        await repository.EnsureCreatedAsync();
        await repository.InsertBatchAsync(
        [
            Message(1, "Quarterly budget", "Budget figures attached.", "Alice", "alice@example.com",
                "cfo@example.com", new DateTime(2026, 1, 3, 9, 0, 0, DateTimeKind.Utc)),
            Message(2, "Project status", "The phoenix rollout is on schedule.", "Bob", "bob@example.com",
                "alice@example.com", new DateTime(2026, 1, 2, 10, 0, 0, DateTimeKind.Utc)),
            Message(3, "Welcome", "Welcome to the team.", "Carol", "carol@example.com",
                "newhire@example.com", new DateTime(2026, 1, 1, 11, 0, 0, DateTimeKind.Utc),
                @"Mailbox\Archive\Inbox")
        ]);
        return repository;
    }

    private static EmailMessage Message(
        long id,
        string subject,
        string body,
        string senderName,
        string senderAddress,
        string to,
        DateTime receivedUtc,
        string folderPath = @"Mailbox\Inbox") => new()
        {
            Id = id,
            FolderPath = folderPath,
            SenderName = senderName,
            SenderAddress = senderAddress,
            ToRecipients = to,
            CcRecipients = "",
            Subject = subject,
            SentUtc = receivedUtc.AddMinutes(-1),
            ReceivedUtc = receivedUtc,
            Preview = body.Split('.')[0] + ".",
            BodyText = body,
            MessageClass = "IPM.Note",
            EntryId = $"entry-{id}"
        };

    public void Dispose()
    {
        SqliteConnection.ClearAllPools();
        if (File.Exists(_databasePath))
            File.Delete(_databasePath);
    }
}

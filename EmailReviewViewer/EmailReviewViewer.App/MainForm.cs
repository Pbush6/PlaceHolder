using EmailReviewViewer.App.Data;

namespace EmailReviewViewer.App;

public sealed class MainForm : Form
{
    private const int PageSize = 200;
    private readonly EmailDatabaseStoreSwitcher _store = new();
    private readonly string? _startupDatabasePath;
    private readonly DateTimePicker _fromDate = new RoundedDateTimePicker
    {
        Name = "FromDate",
        Format = DateTimePickerFormat.Short,
        ShowCheckBox = false
    };
    private readonly CheckBox _fromDateEnabled = new()
    {
        Name = "FromDateEnabled",
        AccessibleName = "Enable from date filter",
        AutoSize = true
    };
    private readonly DateTimePicker _toDate = new RoundedDateTimePicker
    {
        Name = "ToDate",
        Format = DateTimePickerFormat.Short,
        ShowCheckBox = false
    };
    private readonly CheckBox _toDateEnabled = new()
    {
        Name = "ToDateEnabled",
        AccessibleName = "Enable to date filter",
        AutoSize = true
    };
    private readonly TextBox _party = new() { PlaceholderText = "Sender or recipient", Width = 180 };
    private readonly TextBox _keyword = new() { PlaceholderText = "Subject, body, from, to", Width = 220 };
    private readonly TextBox _folderSearch = new() { PlaceholderText = "Search folders", Dock = DockStyle.Fill };
    private readonly CheckBox _allFolders = new() { AutoSize = true, Checked = true, Text = "All Folders" };
    private readonly CheckedListBox _folderList = new()
    {
        BorderStyle = BorderStyle.None,
        CheckOnClick = true,
        Dock = DockStyle.Fill,
        IntegralHeight = false,
        ItemHeight = 26
    };
    private readonly Label _folderStatus = new() { AutoSize = true, ForeColor = SystemColors.GrayText };
    private readonly ToolTip _folderToolTip = new();
    private readonly DataGridView _grid = new() { Name = "EmailGrid" };
    private readonly Label _resultsHeading = new() { Name = "ResultsHeading", AutoSize = true, Text = "Messages" };
    private readonly Label _resultsTotal = new() { Name = "ResultsTotal", AutoSize = true };
    private readonly Label _status = new() { AutoSize = true, TextAlign = ContentAlignment.MiddleLeft };
    private readonly ToolStripStatusLabel _databasePathStatus = new() { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
    private readonly Label _headerDatabasePath = new()
    {
        Name = "HeaderDatabasePath",
        AutoEllipsis = true,
        Text = "No database open"
    };
    private readonly Button _openDatabase = new RoundedButton { Name = "OpenDatabaseButton", Text = "Open Database…" };
    private readonly Button _search = new RoundedButton { Name = "SearchButton", Text = "Apply filters" };
    private readonly Button _resetFilters = new RoundedButton { Name = "ResetFiltersButton", Text = "Reset filters" };
    private readonly Button _previous = new RoundedButton { Text = "Previous" };
    private readonly Button _next = new RoundedButton { Text = "Next" };
    private readonly Label _subject = new()
    {
        Name = "ReadingSubject",
        AutoSize = true,
        Font = new Font(SystemFonts.MessageBoxFont ?? Control.DefaultFont, FontStyle.Bold),
        Text = "Select an email"
    };
    private readonly Label _headers = new()
    {
        Name = "ReadingHeaders",
        AutoSize = true,
        ForeColor = SystemColors.GrayText,
        Text = "Choose a row to view the complete message."
    };
    private readonly RichTextBox _body = new()
    {
        Name = "ReadingBody",
        ReadOnly = true,
        BackColor = SystemColors.Window,
        BorderStyle = BorderStyle.None,
        DetectUrls = false,
        Dock = DockStyle.Fill,
        ScrollBars = RichTextBoxScrollBars.Vertical
    };
    private readonly System.Windows.Forms.Timer _debounce = new() { Interval = 350 };
    private CancellationTokenSource? _pageLoadCancellation;
    private CancellationTokenSource? _detailLoadCancellation;
    private IReadOnlyList<FolderCount> _folders = [];
    private readonly HashSet<string> _selectedFolderPaths = new(StringComparer.Ordinal);
    private bool _updatingFolders;
    private bool _updatingGrid;
    private bool _resetting;
    private long? _selectedMessageId;
    private int _pageIndex;
    private long _totalCount;
    private EmailSortColumn _sortColumn = EmailSortColumn.Date;
    private EmailSortDirection _sortDirection = EmailSortDirection.Descending;

    public MainForm(string? databasePath)
    {
        var palette = AppTheme.Current;
        _startupDatabasePath = string.IsNullOrWhiteSpace(databasePath) ? null : Path.GetFullPath(databasePath);
        Text = "Email Review Viewer";
        Width = 1500;
        Height = 900;
        MinimumSize = new Size(1200, 700);
        StartPosition = FormStartPosition.CenterScreen;
        Font = AppTheme.Font();
        BackColor = palette.Canvas;
        ForeColor = palette.Text;

        var menu = BuildMenu();
        MainMenuStrip = menu;
        Controls.Add(BuildLayout());
        Controls.Add(new StatusStrip
        {
            BackColor = palette.SubtleSurface,
            ForeColor = palette.MutedText,
            SizingGrip = false,
            Items = { _databasePathStatus }
        });
        Controls.Add(menu);
        AppTheme.StyleButton(_openDatabase, ButtonKind.Primary);
        AppTheme.StyleButton(_search, ButtonKind.Primary);
        AppTheme.StyleButton(_resetFilters, ButtonKind.Quiet);
        AppTheme.StyleButton(_previous, ButtonKind.Secondary);
        AppTheme.StyleButton(_next, ButtonKind.Secondary);
        ConfigureDatePickerWidth(_fromDate);
        ConfigureDatePickerWidth(_toDate);
        ConfigureGrid();
        WireEvents();
        ResetReviewState();
        SetDatabaseLoadedState(false);
    }

    private MenuStrip BuildMenu()
    {
        var open = new ToolStripMenuItem("Open Database…") { Name = "OpenDatabaseMenuItem" };
        open.Click += async (_, _) => await ChooseDatabaseAsync();
        var file = new ToolStripMenuItem("&File");
        file.DropDownItems.Add(open);
        return new MenuStrip { Items = { file } };
    }

    private static void ConfigureDatePickerWidth(DateTimePicker picker)
    {
        var firstDay = new DateTime(DateTime.Today.Year, 1, 1);
        var daysInYear = DateTime.IsLeapYear(firstDay.Year) ? 366 : 365;
        var widestDate = Enumerable.Range(0, daysInYear)
            .Select(day => firstDay.AddDays(day).ToString("d"))
            .Max(text => TextRenderer.MeasureText(text, picker.Font).Width);
        var minimumWidth = widestDate + (SystemInformation.VerticalScrollBarWidth * 2);

        picker.MinimumSize = new Size(minimumWidth, 0);
        picker.Width = minimumWidth;
    }

    private Control BuildLayout()
    {
        var palette = AppTheme.Current;
        var pager = BuildPager();
        var readingPane = BuildReadingPane();
        _resultsHeading.ForeColor = palette.Primary;
        _resultsHeading.Font = AppTheme.DisplayFont(12F, FontStyle.Bold);
        _resultsHeading.Anchor = AnchorStyles.Left;
        _resultsHeading.Margin = new Padding(0);
        _resultsTotal.ForeColor = palette.MutedText;
        _resultsTotal.Font = AppTheme.Font(10F, FontStyle.Bold);
        _resultsTotal.Anchor = AnchorStyles.Left;
        _resultsTotal.Margin = new Padding(12, 0, 0, 0);
        var resultsHeader = new TableLayoutPanel
        {
            Name = "ResultsHeader",
            Dock = DockStyle.Top,
            Height = 48,
            ColumnCount = 3,
            Padding = new Padding(AppTheme.SectionPadding, 0, AppTheme.SectionPadding, 0),
            BackColor = palette.Surface
        };
        resultsHeader.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        resultsHeader.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        resultsHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        resultsHeader.Controls.Add(_resultsHeading, 0, 0);
        resultsHeader.Controls.Add(_resultsTotal, 1, 0);
        var resultsPanel = new Panel { Dock = DockStyle.Fill, BackColor = palette.Surface };
        resultsPanel.Controls.Add(_grid);
        resultsPanel.Controls.Add(pager);
        resultsPanel.Controls.Add(resultsHeader);

        var messageSplit = new SplitContainer
        {
            Name = "MessageReadingSplit",
            Dock = DockStyle.Fill,
            Orientation = Orientation.Vertical,
            Size = new Size(1400, 800),
            SplitterDistance = 780,
            Panel1MinSize = 450,
            Panel2MinSize = 380,
            SplitterWidth = AppTheme.SplitterWidth,
            BackColor = palette.Border
        };
        messageSplit.Panel1.Controls.Add(resultsPanel);
        messageSplit.Panel2.Controls.Add(readingPane);

        var contentSplit = new SplitContainer
        {
            Name = "FolderResultsSplit",
            Dock = DockStyle.Fill,
            FixedPanel = FixedPanel.Panel1,
            Size = new Size(1400, 800),
            SplitterDistance = 250,
            Panel1MinSize = 200,
            Panel2MinSize = 850,
            SplitterWidth = AppTheme.SplitterWidth,
            BackColor = palette.Border
        };
        contentSplit.Panel1.Controls.Add(BuildFolderPanel());
        contentSplit.Panel2.Controls.Add(messageSplit);

        var root = new Panel { Dock = DockStyle.Fill, BackColor = palette.Canvas };
        root.Controls.Add(contentSplit);
        root.Controls.Add(BuildFilterToolbar());
        root.Controls.Add(BuildApplicationHeader());
        return root;
    }

    private Control BuildApplicationHeader()
    {
        var palette = AppTheme.Current;
        var title = new Label
        {
            Name = "ApplicationTitle",
            AutoSize = true,
            ForeColor = SystemInformation.HighContrast ? palette.Text : Color.White,
            Font = AppTheme.DisplayFont(20F, FontStyle.Bold),
            Text = "Email Reviewer"
        };
        _headerDatabasePath.Dock = DockStyle.Fill;
        _headerDatabasePath.ForeColor = SystemInformation.HighContrast
            ? palette.MutedText
            : Color.FromArgb(204, 216, 229);
        _headerDatabasePath.Padding = new Padding(0, 3, 0, 0);

        var text = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Margin = new Padding(0),
            BackColor = palette.Primary
        };
        text.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
        text.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        text.Controls.Add(title, 0, 0);
        text.Controls.Add(_headerDatabasePath, 0, 1);

        var header = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 84,
            ColumnCount = 2,
            Padding = new Padding(20, 11, 20, 11),
            BackColor = palette.Primary
        };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        header.Controls.Add(text, 0, 0);
        header.Controls.Add(_openDatabase, 1, 0);
        _openDatabase.Anchor = AnchorStyles.Right;
        _openDatabase.Width = 156;
        return header;
    }

    private Control BuildFilterToolbar()
    {
        var palette = AppTheme.Current;
        foreach (var input in new Control[] { _fromDate, _toDate, _party, _keyword })
            AppTheme.StyleInput(input);
        ConfigureDatePickerWidth(_fromDate);
        ConfigureDatePickerWidth(_toDate);
        _party.Width = 200;
        _keyword.Width = 260;
        _search.Width = 118;
        _resetFilters.Width = 118;
        var fromDateInput = DateFilterInput(
            _fromDateEnabled,
            new RoundedInputHost(_fromDate) { Width = _fromDate.Width + 20 });
        var toDateInput = DateFilterInput(
            _toDateEnabled,
            new RoundedInputHost(_toDate) { Width = _toDate.Width + 20 });
        var partyInput = new RoundedInputHost(_party) { Width = _party.Width, Margin = Padding.Empty };
        var keywordInput = new RoundedInputHost(_keyword) { Width = _keyword.Width, Margin = Padding.Empty };

        var fields = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            WrapContents = false,
            AutoScroll = true,
            Margin = new Padding(0),
            BackColor = palette.Surface
        };
        fields.Controls.AddRange(
        [
            FilterField("From date", fromDateInput),
            FilterField("To date", toDateInput),
            FilterField("Participant", partyInput),
            FilterField("Keyword", keywordInput)
        ]);

        var actions = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            WrapContents = false,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(8, 20, 0, 0),
            Margin = new Padding(0),
            BackColor = palette.Surface
        };
        actions.Controls.AddRange([_search, _resetFilters]);

        var toolbar = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 96,
            ColumnCount = 2,
            Padding = new Padding(20, 10, 20, 10),
            BackColor = palette.Surface
        };
        toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 280));
        toolbar.Controls.Add(fields, 0, 0);
        toolbar.Controls.Add(actions, 1, 0);
        return toolbar;
    }

    private static Control DateFilterInput(CheckBox enabled, RoundedInputHost picker)
    {
        enabled.Margin = new Padding(0, 10, 6, 0);
        picker.Margin = Padding.Empty;
        var panel = new FlowLayoutPanel
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Margin = Padding.Empty,
            BackColor = AppTheme.Current.Surface
        };
        panel.Controls.Add(enabled);
        panel.Controls.Add(picker);
        return panel;
    }

    private static Control FilterField(string label, Control input)
    {
        var panel = new TableLayoutPanel
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            RowCount = 2,
            ColumnCount = 1,
            Margin = new Padding(0, 0, 14, 0)
        };
        panel.Controls.Add(new Label
        {
            AutoSize = true,
            ForeColor = AppTheme.Current.MutedText,
            Font = AppTheme.DisplayFont(9.5F, FontStyle.Bold),
            Text = label,
            Margin = new Padding(0, 0, 0, 3)
        }, 0, 0);
        panel.Controls.Add(input, 0, 1);
        return panel;
    }

    private Control BuildPager()
    {
        var palette = AppTheme.Current;
        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            WrapContents = false,
            Margin = new Padding(0)
        };
        _previous.Width = 92;
        _next.Width = 92;
        buttons.Controls.AddRange([_next, _previous]);

        var pager = new TableLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 52,
            ColumnCount = 2,
            Padding = new Padding(AppTheme.SectionPadding, 9, AppTheme.SectionPadding, 7),
            BackColor = palette.SubtleSurface
        };
        pager.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        pager.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 204));
        pager.Controls.Add(_status, 0, 0);
        pager.Controls.Add(buttons, 1, 0);
        _status.Anchor = AnchorStyles.Left;
        return pager;
    }

    private Control BuildReadingPane()
    {
        var palette = AppTheme.Current;
        _subject.Font = AppTheme.DisplayFont(17F, FontStyle.Bold);
        _subject.ForeColor = palette.Primary;
        _headers.Font = AppTheme.Font(9F);
        _headers.ForeColor = palette.MutedText;
        _body.Font = AppTheme.Font(10F);
        _body.BackColor = palette.Surface;
        _body.ForeColor = palette.Text;

        var eyebrow = new Label
        {
            AutoSize = true,
            ForeColor = palette.Accent,
            Font = AppTheme.DisplayFont(9F, FontStyle.Bold),
            Text = "MESSAGE",
            Margin = new Padding(0, 0, 0, 8)
        };
        var readingHeader = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 1,
            Padding = new Padding(AppTheme.SectionPadding, AppTheme.SectionPadding, AppTheme.SectionPadding, 12),
            BackColor = palette.Surface
        };
        readingHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        readingHeader.Controls.Add(eyebrow);
        readingHeader.Controls.Add(_subject);
        readingHeader.Controls.Add(_headers);
        readingHeader.SizeChanged += (_, _) =>
        {
            var width = Math.Max(1, readingHeader.ClientSize.Width - readingHeader.Padding.Horizontal);
            _subject.MaximumSize = new Size(width, 0);
            _headers.MaximumSize = new Size(width, 0);
        };

        var bodyPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(AppTheme.SectionPadding, 14, AppTheme.SectionPadding, AppTheme.SectionPadding),
            BackColor = palette.Surface
        };
        bodyPanel.Controls.Add(_body);
        var separator = new Panel { Dock = DockStyle.Top, Height = 1, BackColor = palette.Border };
        var pane = new Panel { Dock = DockStyle.Fill, BackColor = palette.Surface };
        pane.Controls.Add(bodyPanel);
        pane.Controls.Add(separator);
        pane.Controls.Add(readingHeader);
        return pane;
    }

    private Control BuildFolderPanel()
    {
        var palette = AppTheme.Current;
        AppTheme.StyleInput(_folderSearch);
        var folderSearchInput = new RoundedInputHost(_folderSearch)
        {
            Dock = DockStyle.Fill,
            Margin = new Padding(0, 2, 0, 10)
        };
        _allFolders.ForeColor = palette.Text;
        _allFolders.Margin = new Padding(2, 2, 0, 8);
        _folderList.BackColor = palette.Surface;
        _folderList.ForeColor = palette.Text;
        _folderList.Font = AppTheme.Font(9.5F);
        _folderStatus.ForeColor = palette.MutedText;
        _folderStatus.Margin = new Padding(2, 8, 0, 0);

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 5,
            Padding = new Padding(AppTheme.SectionPadding),
            BackColor = palette.SubtleSurface
        };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.Controls.Add(new Label
        {
            Name = "FolderHeading",
            AutoSize = true,
            Font = AppTheme.DisplayFont(12F, FontStyle.Bold),
            ForeColor = palette.Primary,
            Text = "Folders",
            Margin = new Padding(0, 0, 0, 10)
        }, 0, 0);
        layout.Controls.Add(folderSearchInput, 0, 1);
        layout.Controls.Add(_allFolders, 0, 2);
        layout.Controls.Add(_folderList, 0, 3);
        layout.Controls.Add(_folderStatus, 0, 4);
        return layout;
    }

    private void ConfigureGrid()
    {
        var palette = AppTheme.Current;
        _grid.Dock = DockStyle.Fill;
        _grid.ReadOnly = true;
        _grid.AllowUserToAddRows = false;
        _grid.AllowUserToDeleteRows = false;
        _grid.AllowUserToResizeRows = false;
        _grid.AutoGenerateColumns = false;
        _grid.MultiSelect = false;
        _grid.RowHeadersVisible = false;
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _grid.RowTemplate.Height = 34;
        _grid.EnableHeadersVisualStyles = false;
        _grid.BackgroundColor = palette.Surface;
        _grid.BorderStyle = BorderStyle.None;
        _grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
        _grid.GridColor = palette.Border;
        _grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
        _grid.ColumnHeadersHeight = 38;
        _grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        _grid.DefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = palette.Surface,
            ForeColor = palette.Text,
            SelectionBackColor = palette.Selection,
            SelectionForeColor = palette.SelectionText,
            Padding = new Padding(7, 2, 7, 2),
            Font = AppTheme.Font(9F)
        };
        _grid.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = palette.SubtleSurface,
            ForeColor = palette.Text,
            SelectionBackColor = palette.Selection,
            SelectionForeColor = palette.SelectionText,
            Padding = new Padding(7, 2, 7, 2)
        };
        _grid.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = SystemInformation.HighContrast ? SystemColors.Control : palette.Primary,
            ForeColor = SystemInformation.HighContrast ? SystemColors.ControlText : Color.White,
            Font = AppTheme.DisplayFont(9F, FontStyle.Bold),
            SelectionBackColor = SystemInformation.HighContrast ? SystemColors.Control : palette.Primary,
            SelectionForeColor = SystemInformation.HighContrast ? SystemColors.ControlText : Color.White,
            Padding = new Padding(7, 0, 7, 0)
        };
        _grid.Columns.AddRange(
        [
            Column(nameof(EmailListItem.DateUtc), "Date", 135, "g"),
            Column(nameof(EmailListItem.SenderName), "From", 135),
            Column(nameof(EmailListItem.ToRecipients), "To", 180),
            Column(nameof(EmailListItem.Subject), "Subject", 220),
            new DataGridViewTextBoxColumn
            {
                Name = nameof(EmailListItem.Preview),
                DataPropertyName = nameof(EmailListItem.Preview),
                HeaderText = "Preview",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 200,
                SortMode = DataGridViewColumnSortMode.NotSortable
            }
        ]);
        UpdateSortGlyph();
    }

    private static DataGridViewTextBoxColumn Column(
        string property,
        string title,
        int width,
        string? format = null) => new()
        {
            Name = property,
            DataPropertyName = property,
            HeaderText = title,
            Width = width,
            SortMode = DataGridViewColumnSortMode.Programmatic,
            DefaultCellStyle = format is null ? null : new DataGridViewCellStyle { Format = format }
        };

    private void WireEvents()
    {
        _fromDate.FontChanged += (_, _) => ConfigureDatePickerWidth(_fromDate);
        _toDate.FontChanged += (_, _) => ConfigureDatePickerWidth(_toDate);
        _fromDate.DpiChangedAfterParent += (_, _) => ConfigureDatePickerWidth(_fromDate);
        _toDate.DpiChangedAfterParent += (_, _) => ConfigureDatePickerWidth(_toDate);
        Shown += async (_, _) =>
        {
            if (_startupDatabasePath is not null)
                await OpenDatabaseAsync(_startupDatabasePath);
        };
        _debounce.Tick += async (_, _) =>
        {
            _debounce.Stop();
            _pageIndex = 0;
            await LoadPageAsync();
        };
        _party.TextChanged += (_, _) => RestartDebounce();
        _keyword.TextChanged += (_, _) => RestartDebounce();
        _fromDate.ValueChanged += (_, _) => RestartDebounce();
        _toDate.ValueChanged += (_, _) => RestartDebounce();
        _fromDateEnabled.CheckedChanged += (_, _) => RestartDebounce();
        _toDateEnabled.CheckedChanged += (_, _) => RestartDebounce();
        _folderSearch.TextChanged += (_, _) => RefreshFolderList();
        _openDatabase.Click += async (_, _) => await ChooseDatabaseAsync();
        _resetFilters.Click += async (_, _) => await ResetFiltersAsync();
        _folderList.ItemCheck += FolderItemCheck;
        _folderList.MouseMove += FolderListMouseMove;
        _allFolders.Click += async (_, _) =>
        {
            if (!_allFolders.Checked)
            {
                _allFolders.Checked = true;
                return;
            }

            _selectedFolderPaths.Clear();
            RefreshFolderList();
            _pageIndex = 0;
            await LoadPageAsync();
        };
        _search.Click += async (_, _) =>
        {
            _debounce.Stop();
            _pageIndex = 0;
            await LoadPageAsync();
        };
        _previous.Click += async (_, _) =>
        {
            if (_pageIndex == 0) return;
            _pageIndex--;
            await LoadPageAsync();
        };
        _next.Click += async (_, _) =>
        {
            if ((_pageIndex + 1L) * PageSize >= _totalCount) return;
            _pageIndex++;
            await LoadPageAsync();
        };
        _grid.ColumnHeaderMouseClick += async (_, e) => await SortByColumnAsync(e.ColumnIndex);
        _grid.CellMouseEnter += (sender, e) =>
        {
            if (e.RowIndex == -1 && e.ColumnIndex >= 0 &&
                EmailSortSql.TryMapColumnKey(_grid.Columns[e.ColumnIndex].DataPropertyName, out _))
                _grid.Cursor = Cursors.Hand;
        };
        _grid.CellMouseLeave += (_, _) => _grid.Cursor = Cursors.Default;
        _grid.SelectionChanged += async (_, _) => await LoadSelectedMessageAsync();
        FormClosed += (_, _) =>
        {
            _pageLoadCancellation?.Cancel();
            _detailLoadCancellation?.Cancel();
            _store.Dispose();
        };
    }

    private async Task SortByColumnAsync(int columnIndex)
    {
        if (_updatingGrid || columnIndex < 0 ||
            !EmailSortSql.TryMapColumnKey(_grid.Columns[columnIndex].DataPropertyName, out var column))
            return;

        if (_sortColumn == column)
        {
            _sortDirection = _sortDirection == EmailSortDirection.Ascending
                ? EmailSortDirection.Descending
                : EmailSortDirection.Ascending;
        }
        else
        {
            _sortColumn = column;
            _sortDirection = column == EmailSortColumn.Date
                ? EmailSortDirection.Descending
                : EmailSortDirection.Ascending;
        }

        UpdateSortGlyph();
        _pageIndex = 0;
        await LoadPageAsync();
    }

    private void UpdateSortGlyph()
    {
        foreach (DataGridViewColumn column in _grid.Columns)
            column.HeaderCell.SortGlyphDirection = SortOrder.None;

        var key = _sortColumn switch
        {
            EmailSortColumn.Date => nameof(EmailListItem.DateUtc),
            EmailSortColumn.From => nameof(EmailListItem.SenderName),
            EmailSortColumn.To => nameof(EmailListItem.ToRecipients),
            EmailSortColumn.Subject => nameof(EmailListItem.Subject),
            _ => throw new ArgumentOutOfRangeException()
        };
        _grid.Columns[key].HeaderCell.SortGlyphDirection =
            _sortDirection == EmailSortDirection.Ascending ? SortOrder.Ascending : SortOrder.Descending;
    }

    private async Task ChooseDatabaseAsync()
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Email databases (*.db)|*.db|All files (*.*)|*.*",
            Title = "Open Email Review Database",
            CheckFileExists = true
        };
        if (dialog.ShowDialog(this) == DialogResult.OK)
            await OpenDatabaseAsync(dialog.FileName);
    }

    private async Task OpenDatabaseAsync(string databasePath)
    {
        _debounce.Stop();
        _pageLoadCancellation?.Cancel();
        _detailLoadCancellation?.Cancel();
        _status.Text = "Validating database…";
        _openDatabase.Enabled = false;
        try
        {
            await _store.SwitchAsync(databasePath);
            ResetReviewState();
            SetDatabaseLoadedState(true);
            Text = $"Email Review Viewer — {Path.GetFileName(databasePath)}";
            _databasePathStatus.Text = Path.GetFullPath(databasePath);
            _headerDatabasePath.Text = Path.GetFullPath(databasePath);
            _headerDatabasePath.AccessibleDescription = $"Current database: {Path.GetFullPath(databasePath)}";
            await LoadFoldersAsync();
            await LoadPageAsync();
        }
        catch (Exception exception)
        {
            SetDatabaseLoadedState(_store.Repository is not null);
            _status.Text = _store.Repository is null ? "Open a database to begin" : "Database unchanged";
            MessageBox.Show(
                this,
                exception.Message,
                "Unable to open database",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        finally
        {
            _openDatabase.Enabled = true;
        }
    }

    private void SetDatabaseLoadedState(bool loaded)
    {
        foreach (var control in new Control[]
                 {
                     _fromDate, _toDate, _fromDateEnabled, _toDateEnabled,
                     _party, _keyword, _folderSearch, _allFolders,
                     _folderList, _search, _resetFilters, _previous, _next, _grid
                 })
        {
            control.Enabled = loaded;
        }
    }

    private void ResetReviewState()
    {
        _resetting = true;
        try
        {
            _pageIndex = 0;
            _totalCount = 0;
            UpdateResultsTotal();
            _folders = [];
            _selectedFolderPaths.Clear();
            _folderSearch.Clear();
            _party.Clear();
            _keyword.Clear();
            _fromDateEnabled.Checked = false;
            _toDateEnabled.Checked = false;
            _allFolders.Checked = true;
            _folderList.Items.Clear();
            _folderStatus.Text = "No database open";
            _grid.DataSource = null;
            ClearGridSelectionAndReadingPane();
            _subject.Text = _store.Repository is null ? "Open a database" : "Select an email";
            _headers.Text = _store.Repository is null
                ? "Use File > Open Database… or the Open Database button to begin."
                : "Choose a row to view the complete message.";
            _status.Text = _store.Repository is null ? "Open a database to begin" : "Loading…";
            if (_store.Repository is null)
            {
                _databasePathStatus.Text = "No database open";
                _headerDatabasePath.Text = "No database open";
            }
        }
        finally
        {
            _resetting = false;
        }
    }

    private async Task LoadFoldersAsync()
    {
        if (_store.Repository is not { } repository)
            return;
        _folders = await repository.GetFolderCountsAsync();
        _allFolders.Text = $"All Folders ({_folders.Sum(folder => folder.Count):N0})";
        RefreshFolderList();
    }

    private async Task ResetFiltersAsync()
    {
        if (_store.Repository is null)
            return;

        _debounce.Stop();
        _resetting = true;
        try
        {
            _fromDateEnabled.Checked = false;
            _toDateEnabled.Checked = false;
            _party.Clear();
            _keyword.Clear();
            _folderSearch.Clear();
            _selectedFolderPaths.Clear();
            _allFolders.Checked = true;
            _pageIndex = 0;
            _sortColumn = EmailSortColumn.Date;
            _sortDirection = EmailSortDirection.Descending;
            RefreshFolderList();
            UpdateSortGlyph();
        }
        finally
        {
            _resetting = false;
        }

        await LoadPageAsync();
    }

    private void RefreshFolderList()
    {
        var search = _folderSearch.Text.Trim();
        var visibleFolders = _folders.Where(folder =>
            search.Length == 0 ||
            folder.FolderPath.Contains(search, StringComparison.OrdinalIgnoreCase)).ToArray();

        _updatingFolders = true;
        _folderList.BeginUpdate();
        try
        {
            _folderList.Items.Clear();
            foreach (var folder in visibleFolders)
            {
                var item = new FolderListItem(folder);
                _folderList.Items.Add(item, _selectedFolderPaths.Contains(folder.FolderPath));
            }
        }
        finally
        {
            _folderList.EndUpdate();
            _updatingFolders = false;
        }

        _allFolders.Checked = _selectedFolderPaths.Count == 0;
        _folderStatus.Text = search.Length == 0
            ? $"{_folders.Count:N0} folders"
            : $"{visibleFolders.Length:N0} of {_folders.Count:N0} folders";
    }

    private void FolderItemCheck(object? sender, ItemCheckEventArgs e)
    {
        if (_updatingFolders || _folderList.Items[e.Index] is not FolderListItem item)
            return;

        if (e.NewValue == CheckState.Checked)
            _selectedFolderPaths.Add(item.Folder.FolderPath);
        else
            _selectedFolderPaths.Remove(item.Folder.FolderPath);

        _allFolders.Checked = _selectedFolderPaths.Count == 0;
        _pageIndex = 0;
        BeginInvoke(new Action(async () => await LoadPageAsync()));
    }

    private void FolderListMouseMove(object? sender, MouseEventArgs e)
    {
        var index = _folderList.IndexFromPoint(e.Location);
        var path = index >= 0 && _folderList.Items[index] is FolderListItem item
            ? item.Folder.FolderPath
            : "";
        if (_folderToolTip.GetToolTip(_folderList) != path)
            _folderToolTip.SetToolTip(_folderList, path);
    }

    private void RestartDebounce()
    {
        if (_resetting || _store.Repository is null)
            return;
        _debounce.Stop();
        _debounce.Start();
    }

    private async Task LoadPageAsync()
    {
        if (_store.Repository is not { } repository)
            return;
        _pageLoadCancellation?.Cancel();
        _pageLoadCancellation?.Dispose();
        _pageLoadCancellation = new CancellationTokenSource();
        var cancellationToken = _pageLoadCancellation.Token;
        ClearGridSelectionAndReadingPane();
        _status.Text = "Loading…";
        try
        {
            var result = await repository.SearchPageAsync(BuildQuery(), cancellationToken);
            if (cancellationToken.IsCancellationRequested) return;
            _totalCount = result.TotalCount;
            UpdateResultsTotal();
            _updatingGrid = true;
            try
            {
                _grid.DataSource = result.Items.ToList();
                _grid.ClearSelection();
                _grid.CurrentCell = null;
            }
            finally
            {
                _updatingGrid = false;
            }
            var first = _totalCount == 0 ? 0 : _pageIndex * PageSize + 1;
            var last = Math.Min((_pageIndex + 1L) * PageSize, _totalCount);
            _status.Text = $"{first:N0}–{last:N0} of {_totalCount:N0}";
            _previous.Enabled = _pageIndex > 0;
            _next.Enabled = last < _totalCount;
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception exception)
        {
            _status.Text = "Query failed";
            MessageBox.Show(this, exception.Message, "Email Review Viewer", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void UpdateResultsTotal() =>
        _resultsTotal.Text = $"Total: {_totalCount:N0}";

    private EmailQuery BuildQuery()
    {
        DateTime? from = _fromDateEnabled.Checked
            ? DateTime.SpecifyKind(_fromDate.Value.Date, DateTimeKind.Local).ToUniversalTime()
            : null;
        DateTime? to = _toDateEnabled.Checked
            ? DateTime.SpecifyKind(_toDate.Value.Date.AddDays(1).AddTicks(-1), DateTimeKind.Local).ToUniversalTime()
            : null;
        return new EmailQuery(
            from,
            to,
            _party.Text,
            _keyword.Text,
            PageSize,
            _pageIndex * PageSize,
            _selectedFolderPaths.ToArray(),
            _sortColumn,
            _sortDirection);
    }

    private async Task LoadSelectedMessageAsync()
    {
        if (_store.Repository is not { } repository)
            return;
        if (_updatingGrid)
            return;
        if (_grid.CurrentRow?.DataBoundItem is not EmailListItem item)
        {
            ClearGridSelectionAndReadingPane();
            return;
        }
        if (_selectedMessageId == item.Id)
            return;

        _detailLoadCancellation?.Cancel();
        _detailLoadCancellation?.Dispose();
        _detailLoadCancellation = new CancellationTokenSource();
        var cancellationToken = _detailLoadCancellation.Token;
        _selectedMessageId = item.Id;
        _subject.Text = string.IsNullOrWhiteSpace(item.Subject) ? "(No subject)" : item.Subject;
        _headers.ForeColor = AppTheme.Current.MutedText;
        _headers.Text = "Loading full message…";
        _body.Clear();

        try
        {
            var message = await repository.GetByIdAsync(item.Id, cancellationToken);
            if (cancellationToken.IsCancellationRequested || _selectedMessageId != item.Id)
                return;
            if (message is null)
            {
                _subject.Text = "Message unavailable";
                _headers.Text = "This email may have been deleted or moved.";
                _body.Text = "The selected record no longer exists in the review database.";
                return;
            }

            _subject.Text = string.IsNullOrWhiteSpace(message.Subject) ? "(No subject)" : message.Subject;
            _headers.ForeColor = AppTheme.Current.Text;
            _headers.Text =
                $"Date: {FormatMessageDate(message)}{Environment.NewLine}" +
                $"From: {FormatSender(message)}{Environment.NewLine}" +
                $"To: {message.ToRecipients}{Environment.NewLine}" +
                $"Cc: {message.CcRecipients}{Environment.NewLine}" +
                $"Folder: {message.FolderPath}";
            _body.Text = message.BodyText;
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception exception)
        {
            if (cancellationToken.IsCancellationRequested || _selectedMessageId != item.Id)
                return;
            _headers.ForeColor = AppTheme.Current.MutedText;
            _headers.Text = "Unable to load this email.";
            _body.Text = $"Unable to load message: {exception.Message}";
        }
    }

    private void ClearGridSelectionAndReadingPane()
    {
        _detailLoadCancellation?.Cancel();
        _selectedMessageId = null;
        _updatingGrid = true;
        try
        {
            _grid.ClearSelection();
            _grid.CurrentCell = null;
        }
        finally
        {
            _updatingGrid = false;
        }

        _subject.Text = "Select an email";
        _headers.ForeColor = AppTheme.Current.MutedText;
        _headers.Text = "Choose a row to view the complete message.";
        _body.Clear();
    }

    private static string FormatMessageDate(EmailMessage message) =>
        (message.ReceivedUtc ?? message.SentUtc)?.ToLocalTime().ToString("g") ?? "(Unknown)";

    private static string FormatSender(EmailMessage message)
    {
        if (string.IsNullOrWhiteSpace(message.SenderName))
            return message.SenderAddress;
        return string.IsNullOrWhiteSpace(message.SenderAddress)
            ? message.SenderName
            : $"{message.SenderName} <{message.SenderAddress}>";
    }

    private sealed record FolderListItem(FolderCount Folder)
    {
        public override string ToString()
        {
            var path = Folder.FolderPath.TrimEnd('\\', '/');
            var separator = path.LastIndexOfAny(['\\', '/']);
            var leaf = separator >= 0 ? path[(separator + 1)..] : path;
            return $"{(leaf.Length == 0 ? "(No folder)" : leaf)} ({Folder.Count:N0})";
        }
    }
}

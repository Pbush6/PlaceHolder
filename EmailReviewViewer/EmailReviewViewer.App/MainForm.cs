using EmailReviewViewer.App.Data;

namespace EmailReviewViewer.App;

public sealed class MainForm : Form
{
    private const int PageSize = 200;
    private readonly EmailRepository _repository;
    private readonly DateTimePicker _fromDate = new()
    {
        Name = "FromDate",
        Format = DateTimePickerFormat.Short,
        ShowCheckBox = true
    };
    private readonly DateTimePicker _toDate = new()
    {
        Name = "ToDate",
        Format = DateTimePickerFormat.Short,
        ShowCheckBox = true
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
        IntegralHeight = false
    };
    private readonly Label _folderStatus = new() { AutoSize = true, ForeColor = SystemColors.GrayText };
    private readonly ToolTip _folderToolTip = new();
    private readonly DataGridView _grid = new() { Name = "EmailGrid" };
    private readonly Label _status = new() { AutoSize = true, Padding = new Padding(8, 7, 0, 0) };
    private readonly Button _previous = new() { Text = "Previous", AutoSize = true };
    private readonly Button _next = new() { Text = "Next", AutoSize = true };
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
    private long? _selectedMessageId;
    private int _pageIndex;
    private long _totalCount;

    public MainForm(string databasePath)
    {
        _repository = new EmailRepository(databasePath);
        Text = $"Email Review Viewer — {Path.GetFileName(databasePath)}";
        Width = 1500;
        Height = 900;
        MinimumSize = new Size(1200, 700);
        StartPosition = FormStartPosition.CenterScreen;

        Controls.Add(BuildLayout());
        ConfigureDatePickerWidth(_fromDate);
        ConfigureDatePickerWidth(_toDate);
        ConfigureGrid();
        WireEvents();
    }

    private static void ConfigureDatePickerWidth(DateTimePicker picker)
    {
        var firstDay = new DateTime(DateTime.Today.Year, 1, 1);
        var daysInYear = DateTime.IsLeapYear(firstDay.Year) ? 366 : 365;
        var widestDate = Enumerable.Range(0, daysInYear)
            .Select(day => firstDay.AddDays(day).ToString("d"))
            .Max(text => TextRenderer.MeasureText(text, picker.Font).Width);
        using var graphics = picker.CreateGraphics();
        var checkboxWidth = CheckBoxRenderer.GetGlyphSize(
            graphics,
            System.Windows.Forms.VisualStyles.CheckBoxState.UncheckedNormal).Width;
        var safetyPadding = (int)Math.Ceiling(12 * picker.DeviceDpi / 96d);
        var minimumWidth =
            widestDate + checkboxWidth + SystemInformation.VerticalScrollBarWidth + safetyPadding;

        picker.MinimumSize = new Size(minimumWidth, 0);
        picker.Width = minimumWidth;
    }

    private Control BuildLayout()
    {
        var filters = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 42,
            Padding = new Padding(6),
            WrapContents = false
        };
        filters.Controls.AddRange(
        [
            new Label { Text = "From:", AutoSize = true, Padding = new Padding(0, 7, 0, 0) },
            _fromDate,
            new Label { Text = "To:", AutoSize = true, Padding = new Padding(8, 7, 0, 0) },
            _toDate,
            new Label { Text = "People:", AutoSize = true, Padding = new Padding(8, 7, 0, 0) },
            _party,
            new Label { Text = "Keyword:", AutoSize = true, Padding = new Padding(8, 7, 0, 0) },
            _keyword,
            new Button { Name = "SearchButton", Text = "Search", AutoSize = true },
            _status
        ]);

        var pager = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 38,
            Padding = new Padding(6),
            FlowDirection = FlowDirection.RightToLeft
        };
        pager.Controls.AddRange([_next, _previous]);

        var readingHeader = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 1,
            Padding = new Padding(10)
        };
        readingHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        readingHeader.Controls.Add(_subject);
        readingHeader.Controls.Add(_headers);
        readingHeader.SizeChanged += (_, _) =>
        {
            var maximumTextWidth = Math.Max(1, readingHeader.ClientSize.Width - readingHeader.Padding.Horizontal);
            _subject.MaximumSize = new Size(maximumTextWidth, 0);
            _headers.MaximumSize = new Size(maximumTextWidth, 0);
        };

        var readingPane = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
        readingPane.Controls.Add(_body);
        readingPane.Controls.Add(readingHeader);

        var messageSplit = new SplitContainer
        {
            Name = "MessageReadingSplit",
            Dock = DockStyle.Fill,
            Orientation = Orientation.Vertical,
            Size = new Size(1400, 800),
            SplitterDistance = 780,
            Panel1MinSize = 450,
            Panel2MinSize = 380
        };
        messageSplit.Panel1.Controls.Add(_grid);
        messageSplit.Panel1.Controls.Add(pager);
        messageSplit.Panel2.Controls.Add(readingPane);

        var contentSplit = new SplitContainer
        {
            Name = "FolderResultsSplit",
            Dock = DockStyle.Fill,
            FixedPanel = FixedPanel.Panel1,
            Size = new Size(1400, 800),
            SplitterDistance = 250,
            Panel1MinSize = 200,
            Panel2MinSize = 850
        };
        contentSplit.Panel1.Controls.Add(BuildFolderPanel());
        contentSplit.Panel2.Controls.Add(messageSplit);

        var root = new Panel { Dock = DockStyle.Fill };
        root.Controls.Add(contentSplit);
        root.Controls.Add(filters);
        return root;
    }

    private Control BuildFolderPanel()
    {
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 5,
            Padding = new Padding(8)
        };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.Controls.Add(new Label
        {
            AutoSize = true,
            Font = new Font(SystemFonts.MessageBoxFont ?? Control.DefaultFont, FontStyle.Bold),
            Text = "Folders",
            Margin = new Padding(3, 0, 3, 6)
        }, 0, 0);
        layout.Controls.Add(_folderSearch, 0, 1);
        layout.Controls.Add(_allFolders, 0, 2);
        layout.Controls.Add(_folderList, 0, 3);
        layout.Controls.Add(_folderStatus, 0, 4);
        return layout;
    }

    private void ConfigureGrid()
    {
        _grid.Dock = DockStyle.Fill;
        _grid.ReadOnly = true;
        _grid.AllowUserToAddRows = false;
        _grid.AllowUserToDeleteRows = false;
        _grid.AllowUserToResizeRows = false;
        _grid.AutoGenerateColumns = false;
        _grid.MultiSelect = false;
        _grid.RowHeadersVisible = false;
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _grid.RowTemplate.Height = 25;
        _grid.Columns.AddRange(
        [
            Column(nameof(EmailListItem.DateUtc), "Date", 135, "g"),
            Column(nameof(EmailListItem.SenderName), "From", 135),
            Column(nameof(EmailListItem.ToRecipients), "To", 180),
            Column(nameof(EmailListItem.Subject), "Subject", 220),
            new DataGridViewTextBoxColumn
            {
                DataPropertyName = nameof(EmailListItem.Preview),
                HeaderText = "Preview",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 200
            }
        ]);
    }

    private static DataGridViewTextBoxColumn Column(
        string property,
        string title,
        int width,
        string? format = null) => new()
        {
            DataPropertyName = property,
            HeaderText = title,
            Width = width,
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
            await _repository.EnsureCreatedAsync();
            await LoadFoldersAsync();
            await LoadPageAsync();
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
        _folderSearch.TextChanged += (_, _) => RefreshFolderList();
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
        Controls.Find("SearchButton", true).Single().Click += async (_, _) =>
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
        _grid.SelectionChanged += async (_, _) => await LoadSelectedMessageAsync();
        FormClosed += (_, _) =>
        {
            _pageLoadCancellation?.Cancel();
            _detailLoadCancellation?.Cancel();
        };
    }

    private async Task LoadFoldersAsync()
    {
        _folders = await _repository.GetFolderCountsAsync();
        _allFolders.Text = $"All Folders ({_folders.Sum(folder => folder.Count):N0})";
        RefreshFolderList();
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
        _debounce.Stop();
        _debounce.Start();
    }

    private async Task LoadPageAsync()
    {
        _pageLoadCancellation?.Cancel();
        _pageLoadCancellation?.Dispose();
        _pageLoadCancellation = new CancellationTokenSource();
        var cancellationToken = _pageLoadCancellation.Token;
        ClearGridSelectionAndReadingPane();
        _status.Text = "Loading…";
        try
        {
            var result = await _repository.SearchPageAsync(BuildQuery(), cancellationToken);
            if (cancellationToken.IsCancellationRequested) return;
            _totalCount = result.TotalCount;
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

    private EmailQuery BuildQuery()
    {
        DateTime? from = _fromDate.Checked
            ? DateTime.SpecifyKind(_fromDate.Value.Date, DateTimeKind.Local).ToUniversalTime()
            : null;
        DateTime? to = _toDate.Checked
            ? DateTime.SpecifyKind(_toDate.Value.Date.AddDays(1).AddTicks(-1), DateTimeKind.Local).ToUniversalTime()
            : null;
        return new EmailQuery(
            from,
            to,
            _party.Text,
            _keyword.Text,
            PageSize,
            _pageIndex * PageSize,
            _selectedFolderPaths.ToArray());
    }

    private async Task LoadSelectedMessageAsync()
    {
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
        _headers.ForeColor = SystemColors.GrayText;
        _headers.Text = "Loading full message…";
        _body.Clear();

        try
        {
            var message = await _repository.GetByIdAsync(item.Id, cancellationToken);
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
            _headers.ForeColor = SystemColors.ControlText;
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
            _headers.ForeColor = SystemColors.GrayText;
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
        _headers.ForeColor = SystemColors.GrayText;
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

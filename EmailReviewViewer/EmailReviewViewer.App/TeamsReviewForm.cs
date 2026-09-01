using System.Globalization;
using System.Text.Json;
using EmailReviewViewer.App.Data;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace EmailReviewViewer.App;

public sealed class TeamsReviewForm : Form
{
    private const int PageSize = 25;
    private const string VirtualHost = "teams-report.local";
    private readonly string _htmlFolder = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "EmailReviewViewer",
        "TeamsReportHtml");
    private readonly TeamsDatabaseStoreSwitcher _store = new();
    private readonly string? _startupDatabasePath;
    private readonly WebView2 _webView = new() { Name = "TeamsReportView", Dock = DockStyle.Fill };
    private CancellationTokenSource? _pageLoadCancellation;
    private IReadOnlyList<string> _people = [];
    private IReadOnlyList<string> _otherNames = [];
    private IReadOnlyList<FolderCount> _folders = [];
    private TeamsReportSummary? _summary;
    private string _personSearch = "";
    private string _keyword = "";
    private TeamsParticipantMatchMode _matchMode;
    private string _fromDate = "";
    private string _toDate = "";
    private bool _newestFirst = true;
    private string[] _selected = [];
    private int _pageIndex;
    private long _conversationCount;
    private bool _webViewReady;

    public TeamsReviewForm(string? databasePath)
    {
        _startupDatabasePath = string.IsNullOrWhiteSpace(databasePath) ? null : Path.GetFullPath(databasePath);
        Text = "Teams Review Viewer";
        Width = 1500;
        Height = 900;
        MinimumSize = new Size(1200, 700);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 9F);
        BackColor = Color.FromArgb(244, 246, 251);

        var file = new ToolStripMenuItem("&File");
        var open = new ToolStripMenuItem("&Open Database…") { Name = "OpenDatabaseMenuItem" };
        open.Click += async (_, _) => await ChooseDatabaseAsync();
        file.DropDownItems.Add(open);
        var menu = new MenuStrip { Items = { file } };
        MainMenuStrip = menu;
        Controls.Add(_webView);
        Controls.Add(menu);
        Shown += async (_, _) => await InitializeAsync();
        FormClosed += (_, _) =>
        {
            _pageLoadCancellation?.Cancel();
            _store.Dispose();
            _webView.Dispose();
        };
    }

    private async Task InitializeAsync()
    {
        try
        {
            var userData = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "EmailReviewViewer",
                "WebView2-Teams");
            Directory.CreateDirectory(userData);
            var environment = await CoreWebView2Environment.CreateAsync(null, userData);
            await _webView.EnsureCoreWebView2Async(environment);
            Directory.CreateDirectory(_htmlFolder);
            _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                VirtualHost,
                _htmlFolder,
                CoreWebView2HostResourceAccessKind.Allow);
            _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
            _webView.CoreWebView2.Settings.AreDevToolsEnabled = false;
            _webView.CoreWebView2.WebMessageReceived += (_, args) =>
                _ = HandleWebMessageAsync(args.TryGetWebMessageAsString());
            _webViewReady = true;
        }
        catch (Exception exception)
        {
            MessageBox.Show(
                this,
                "The Teams report needs the Microsoft Edge WebView2 Runtime.\n\n" + exception.Message,
                "Teams Review Viewer",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return;
        }

        ShowHtml(EmptyModel());
        if (_startupDatabasePath is not null)
            await OpenDatabaseAsync(_startupDatabasePath);
    }

    private async Task HandleWebMessageAsync(string? json)
    {
        if (string.IsNullOrWhiteSpace(json))
            return;
        TeamsWebMessage? message;
        try
        {
            message = JsonSerializer.Deserialize<TeamsWebMessage>(json, JsonOptions);
        }
        catch (JsonException)
        {
            return;
        }

        if (message?.Action == "openDatabase")
        {
            await ChooseDatabaseAsync();
            return;
        }

        if (message?.Action == "page")
        {
            var next = _pageIndex + message.Delta;
            if (next < 0)
                return;
            if ((next * (long)PageSize) >= _conversationCount && _conversationCount > 0)
                return;
            _pageIndex = next;
            await LoadPageAsync();
            return;
        }

        if (message?.Action != "query")
            return;

        _personSearch = message.PersonSearch ?? "";
        _keyword = message.Keyword ?? "";
        _matchMode = ParseMatchMode(message.MatchMode);
        _fromDate = message.StartDate ?? "";
        _toDate = message.EndDate ?? "";
        _newestFirst = message.NewestFirst;
        _selected = message.Selected ?? [];
        _pageIndex = 0;
        await LoadPageAsync();
    }

    private async Task ChooseDatabaseAsync()
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Review databases (*.db)|*.db|All files (*.*)|*.*",
            Title = "Open Teams Review Database",
            CheckFileExists = true
        };
        if (dialog.ShowDialog(this) == DialogResult.OK)
            await OpenDatabaseAsync(dialog.FileName);
    }

    private async Task OpenDatabaseAsync(string databasePath)
    {
        _pageLoadCancellation?.Cancel();
        try
        {
            var kind = await ReportDatabaseKindDetector.DetectAsync(databasePath);
            if (kind == ReportDatabaseKind.Email)
            {
                ReportViewerSession.Request(databasePath);
                Close();
                return;
            }

            await _store.SwitchAsync(databasePath);
            ResetFilters();
            Text = $"Teams Review Viewer — {Path.GetFileName(databasePath)}";
            if (_store.Repository is { } repository)
            {
                _summary = await Task.Run(() => repository.GetReportSummaryAsync());
                _folders = await Task.Run(() => repository.GetFolderCountsAsync());
                _people = _summary.People;
                _otherNames = _summary.OtherNames;
            }
            await LoadPageAsync();
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Unable to open database", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ResetFilters()
    {
        _pageIndex = 0;
        _conversationCount = 0;
        _personSearch = "";
        _keyword = "";
        _matchMode = TeamsParticipantMatchMode.MessagesFromSelected;
        _fromDate = "";
        _toDate = "";
        _newestFirst = true;
        _selected = [];
        _people = [];
        _otherNames = [];
        _folders = [];
        _summary = null;
    }

    private async Task LoadPageAsync()
    {
        if (_store.Repository is not { } repository || !_webViewReady)
            return;
        _pageLoadCancellation?.Cancel();
        _pageLoadCancellation?.Dispose();
        _pageLoadCancellation = new CancellationTokenSource();
        var cancellationToken = _pageLoadCancellation.Token;
        try
        {
            var query = BuildQuery();
            var html = await Task.Run(async () =>
            {
                var result = await repository.SearchConversationsAsync(query, cancellationToken);
                var conversations = new List<TeamsConversationHtml>(result.Items.Count);
                foreach (var conversation in result.Items)
                {
                    var messages = await repository.GetConversationMessagesAsync(conversation.ConversationKey, cancellationToken);
                    conversations.Add(new TeamsConversationHtml(conversation, FilterMessages(messages, query)));
                }
                return (result, conversations);
            }, cancellationToken);
            if (cancellationToken.IsCancellationRequested)
                return;
            _conversationCount = html.result.TotalCount;
            ShowHtml(BuildModel(html.result, html.conversations));
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, exception.Message, "Teams Review Viewer", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private TeamsQuery BuildQuery()
    {
        DateTime? from = ParseDay(_fromDate, endOfDay: false);
        DateTime? to = ParseDay(_toDate, endOfDay: true);
        return new TeamsQuery(
            from,
            to,
            Keyword: _keyword,
            Limit: PageSize,
            Offset: _pageIndex * PageSize,
            SelectedParticipants: _selected,
            MatchMode: _matchMode,
            NewestFirst: _newestFirst);
    }

    private TeamsReportHtmlModel BuildModel(
        TeamsConversationPage page,
        IReadOnlyList<TeamsConversationHtml> conversations)
    {
        var selectedSet = _selected.ToHashSet(StringComparer.Ordinal);
        return new TeamsReportHtmlModel(
            PstName: TeamsReportHtml.ResolvePstName(_folders, _store.Repository?.DatabasePath),
            Generated: _store.Repository is null
                ? "—"
                : TeamsReportHtml.FormatGenerated(File.GetLastWriteTime(_store.Repository.DatabasePath)),
            MessageCount: _summary?.MessageCount ?? 0,
            PeopleCount: _summary?.PeopleCount ?? 0,
            FolderCount: _summary?.FolderCount ?? 0,
            ReadWarnings: "Items: 0; Attachments: 0",
            People: _people,
            OtherNames: _otherNames,
            SelectedPeople: _people.Where(selectedSet.Contains).ToArray(),
            SelectedOtherNames: _otherNames.Where(selectedSet.Contains).ToArray(),
            Folders: _folders,
            Conversations: conversations,
            PersonSearch: _personSearch,
            Keyword: _keyword,
            MatchMode: _matchMode,
            FromDate: _fromDate,
            ToDate: _toDate,
            NewestFirst: _newestFirst,
            ConversationCount: page.TotalCount,
            VisibleMessageCount: page.VisibleMessageCount,
            PageIndex: _pageIndex,
            PageSize: PageSize,
            DatabasePath: _store.Repository?.DatabasePath);
    }

    private TeamsReportHtmlModel EmptyModel() => new(
        PstName: "—",
        Generated: "—",
        MessageCount: 0,
        PeopleCount: 0,
        FolderCount: 0,
        ReadWarnings: "Items: 0; Attachments: 0",
        People: [],
        OtherNames: [],
        SelectedPeople: [],
        SelectedOtherNames: [],
        Folders: [],
        Conversations: [],
        PersonSearch: "",
        Keyword: "",
        MatchMode: TeamsParticipantMatchMode.MessagesFromSelected,
        FromDate: "",
        ToDate: "",
        NewestFirst: true,
        ConversationCount: 0,
        VisibleMessageCount: 0,
        PageIndex: 0,
        PageSize: PageSize,
        DatabasePath: null);

    private void ShowHtml(TeamsReportHtmlModel model)
    {
        if (!_webViewReady || _webView.CoreWebView2 is null)
            return;
        TeamsReportHtml.WriteReport(_htmlFolder, model);
        _webView.CoreWebView2.Navigate($"https://{VirtualHost}/report.html?t={DateTime.UtcNow.Ticks}");
    }

    private static IReadOnlyList<TeamsMessage> FilterMessages(IReadOnlyList<TeamsMessage> messages, TeamsQuery query)
    {
        var selected = query.SelectedParticipants ?? [];
        return messages.Where(message =>
        {
            var when = message.ReceivedUtc ?? message.SentUtc;
            if (query.FromUtc.HasValue && (when is null || when < query.FromUtc))
                return false;
            if (query.ToUtc.HasValue && (when is null || when > query.ToUtc))
                return false;
            return TeamsParticipantMatching.Matches(
                query.MatchMode,
                selected,
                TeamsParticipantMatching.Split(message.Participants),
                message.SenderDisplay);
        }).ToList();
    }

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private static TeamsParticipantMatchMode ParseMatchMode(string? value) => value switch
    {
        "involving" => TeamsParticipantMatchMode.Involving,
        "exactAll" => TeamsParticipantMatchMode.ExactAll,
        "exactPeopleOnly" => TeamsParticipantMatchMode.ExactPeopleOnly,
        _ => TeamsParticipantMatchMode.MessagesFromSelected
    };

    private static DateTime? ParseDay(string? value, bool endOfDay)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;
        if (!DateTime.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var day))
            return null;
        var local = endOfDay ? day.Date.AddDays(1).AddTicks(-1) : day.Date;
        return DateTime.SpecifyKind(local, DateTimeKind.Local).ToUniversalTime();
    }

    private sealed class TeamsWebMessage
    {
        public string? Action { get; set; }
        public string? PersonSearch { get; set; }
        public string? Keyword { get; set; }
        public string? MatchMode { get; set; }
        public string? StartDate { get; set; }
        public string? EndDate { get; set; }
        public bool NewestFirst { get; set; } = true;
        public string[]? Selected { get; set; }
        public int Delta { get; set; }
    }
}

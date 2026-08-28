using System.Globalization;
using System.Net;
using System.Text;
using EmailReviewViewer.App.Data;

namespace EmailReviewViewer.App;

public sealed record TeamsConversationHtml(
    TeamsConversationListItem Header,
    IReadOnlyList<TeamsMessage> Messages);

public sealed record TeamsReportHtmlModel(
    string PstName,
    string Generated,
    long MessageCount,
    int PeopleCount,
    int FolderCount,
    string ReadWarnings,
    IReadOnlyList<string> People,
    IReadOnlyList<string> OtherNames,
    IReadOnlyList<string> SelectedPeople,
    IReadOnlyList<string> SelectedOtherNames,
    IReadOnlyList<FolderCount> Folders,
    IReadOnlyList<TeamsConversationHtml> Conversations,
    string PersonSearch,
    string Keyword,
    TeamsParticipantMatchMode MatchMode,
    string FromDate,
    string ToDate,
    bool NewestFirst,
    long ConversationCount,
    long VisibleMessageCount,
    int PageIndex,
    int PageSize,
    string? DatabasePath);

public static class TeamsReportHtml
{
    private static readonly string[] SenderPalette =
        ["blue", "green", "purple", "orange", "red", "teal", "pink", "brown", "slate", "indigo"];

    private static readonly Lazy<string> CssValue = new(LoadCss);

    public static string Css => CssValue.Value;

    public static string Build(TeamsReportHtmlModel model)
    {
        var selected = new HashSet<string>(
            (model.SelectedPeople ?? []).Concat(model.SelectedOtherNames ?? []),
            StringComparer.Ordinal);
        var senderClasses = new Dictionary<string, string>(StringComparer.Ordinal);
        var nextSender = 0;
        var html = new StringBuilder(32_768);
        html.Append("""
            <!doctype html>
            <html lang="en">
            <head>
            <meta charset="utf-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <title>Purview Teams PST Conversation Report - 
            """);
        html.Append(Enc(model.PstName));
        html.Append("</title><style>\n");
        html.Append(Css);
        html.Append("</style></head><body><div class='page'>");
        html.Append("""
            <header class='hero'>
              <h1>Microsoft Purview eDiscovery Teams Conversation Report</h1>
              <p>Use the simple filter below to show conversations between only the people you choose.</p>
              <button type='button' class='hero-open' id='openDatabase'>Open Database…</button>
              <div class='hero-credit'>By Patrick Bush</div>
            </header>
            <section class='summary-grid' aria-label='Report summary'>
            """);
        html.Append(SummaryCard("PST", model.PstName));
        html.Append(SummaryCard("Generated", model.Generated));
        html.Append(SummaryCard("Messages exported", model.MessageCount.ToString("N0")));
        html.Append(SummaryCard("People detected", model.PeopleCount.ToString("N0")));
        html.Append(SummaryCard("Folders scanned", model.FolderCount.ToString("N0")));
        html.Append(SummaryCard("Read warnings", model.ReadWarnings));
        html.Append("</section><div class='review-layout'><aside class='filter-panel' aria-label='Conversation filters'>");
        html.Append("<div class='filter-title'><h2>Choose whose conversations to view</h2>");
        html.Append("<span id='resultCount' class='result-count'>");
        html.Append(Enc($"{model.ConversationCount:N0} conversations / {model.VisibleMessageCount:N0} messages shown"));
        html.Append("</span></div>");
        html.Append("""
            <p class='filter-help'>Check one or more names, then choose how strictly to match them. Use "Exact selected people, ignoring Other/IDs" when Purview adds bots, meeting artifacts, or system IDs that should not count as real conversation participants.</p>
            <div class='controls'>
            """);
        html.Append($"<input id='personSearch' type='text' placeholder='Find a person in the list...' value='{Enc(model.PersonSearch)}'/>");
        html.Append($"<input id='messageSearch' type='text' placeholder='Search within shown conversations...' value='{Enc(model.Keyword)}'/>");
        html.Append("<label for='participantMatchMode'>Participant match</label>");
        html.Append("<select id='participantMatchMode'>");
        html.Append(Option("messagesFromSelected", "Messages from selected people", model.MatchMode == TeamsParticipantMatchMode.MessagesFromSelected));
        html.Append(Option("involving", "Conversations involving selected people", model.MatchMode == TeamsParticipantMatchMode.Involving));
        html.Append(Option("exactAll", "Exact selected people only", model.MatchMode == TeamsParticipantMatchMode.ExactAll));
        html.Append(Option("exactPeopleOnly", "Exact selected people, ignoring Other/IDs", model.MatchMode == TeamsParticipantMatchMode.ExactPeopleOnly));
        html.Append("</select></div>");
        html.Append("<div class='people-heading'><h3>People</h3><button type='button' class='secondary' id='clearAll'>Clear filters</button></div>");
        html.Append("<div id='peopleBox' class='people-box'>");
        html.Append("<label class='person-option'><input type='checkbox' id='selectAllPeople'/> <span>Select All</span></label>");
        foreach (var person in model.People)
            html.Append(PersonOption(person, selected.Contains(person)));
        html.Append("</div>");
        html.Append($"<details class='other-names'><summary>Other detected names / IDs ({model.OtherNames.Count})</summary>");
        html.Append("<div id='otherPeopleBox' class='people-box'>");
        html.Append("<label class='person-option'><input type='checkbox' id='selectAllOther'/> <span>Select All</span></label>");
        foreach (var name in model.OtherNames)
            html.Append(PersonOption(name, selected.Contains(name)));
        html.Append("</div></details>");
        html.Append("<div class='notice'>Message bodies are HTML-encoded as text for safe review. Attachment files are not extracted; attachment metadata is listed where Outlook exposes it.</div>");
        html.Append("<details class='folder-summary'><summary>Folder summary</summary><table><thead><tr><th>Folder</th><th>Items</th></tr></thead><tbody>");
        foreach (var folder in model.Folders)
            html.Append($"<tr><td>{Enc(folder.FolderPath)}</td><td>{folder.Count:N0}</td></tr>");
        html.Append("</tbody></table></details>");
        html.Append("<div class='footer'>");
        html.Append(string.IsNullOrWhiteSpace(model.DatabasePath)
            ? "Open a database from File &gt; Open Database."
            : "Database: " + Enc(model.DatabasePath));
        html.Append("<br/>Created by Convert-PurviewTeamsPstToHtml.ps1</div>");
        html.Append("</aside>");
        html.Append("<div id='resizeHandle' class='resize-handle' role='separator' aria-orientation='vertical' aria-label='Resize filter column' tabindex='0' title='Drag to resize the filter column'></div>");
        html.Append("<main id='conversationList' class='conversation-pane'>");
        html.Append("""
            <section class='conversation-toolbar' aria-label='Date filter and conversation sorting'>
              <div class='toolbar-field'>
                <label for='startDateFilter'>Start date</label>
            """);
        html.Append($"<input id='startDateFilter' type='date' value='{Enc(model.FromDate)}'/>");
        html.Append("</div><div class='toolbar-field'><label for='endDateFilter'>End date</label>");
        html.Append($"<input id='endDateFilter' type='date' value='{Enc(model.ToDate)}'/>");
        html.Append("</div><div class='toolbar-field'><label for='sortOrder'>Sort conversations</label><select id='sortOrder'>");
        html.Append(Option("newestFirst", "Newest first", model.NewestFirst));
        html.Append(Option("oldestFirst", "Oldest first", !model.NewestFirst));
        html.Append("</select></div></section>");
        html.Append(Pager(model));
        foreach (var conversation in model.Conversations)
            AppendConversation(html, conversation, senderClasses, ref nextSender);
        html.Append("</main></div></div><script>");
        html.Append(Script);
        html.Append("</script></body></html>");
        return html.ToString();
    }

    // ponytail: WebView2 NavigateToString is capped at 2MB and throws E_INVALIDARG
    public static string WriteReport(string directory, TeamsReportHtmlModel model)
    {
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "report.html");
        File.WriteAllText(path, Build(model), new UTF8Encoding(false));
        return path;
    }

    private static string Pager(TeamsReportHtmlModel model)
    {
        var first = model.ConversationCount == 0 ? 0 : model.PageIndex * model.PageSize + 1;
        var last = Math.Min((model.PageIndex + 1L) * model.PageSize, model.ConversationCount);
        var prevDisabled = model.PageIndex <= 0 ? " disabled" : "";
        var nextDisabled = last >= model.ConversationCount ? " disabled" : "";
        return $"""
            <div class='conversation-pager'>
              <button type='button' class='secondary' id='previousPage'{prevDisabled}>Previous</button>
              <span class='result-count'>{first:N0}–{last:N0} of {model.ConversationCount:N0}</span>
              <button type='button' class='secondary' id='nextPage'{nextDisabled}>Next</button>
            </div>
            """;
    }

    private static void AppendConversation(
        StringBuilder html,
        TeamsConversationHtml conversation,
        Dictionary<string, string> senderClasses,
        ref int nextSender)
    {
        var header = conversation.Header;
        var people = FormatPeople(header.Participants);
        var title = string.IsNullOrWhiteSpace(header.ConversationTitle) ? people : header.ConversationTitle;
        var participants = TeamsParticipantMatching.Split(header.Participants);
        var personParticipants = participants.Where(TeamsParticipantMatching.IsLikelyPersonName).ToArray();
        var otherParticipants = participants.Where(name => !TeamsParticipantMatching.IsLikelyPersonName(name)).ToArray();
        var countText = header.MessageCount == header.TotalMessageCount
            ? $"{header.TotalMessageCount:N0} messages"
            : $"{header.MessageCount:N0} of {header.TotalMessageCount:N0} messages";
        var sortTime = header.LastUtc?.ToUniversalTime().ToString("o") ?? "";
        html.Append($"<section class='conversation' data-participants='{Enc(string.Join("||", participants))}' data-person-participants='{Enc(string.Join("||", personParticipants))}' data-other-participants='{Enc(string.Join("||", otherParticipants))}' data-sort-time='{Enc(sortTime)}'>");
        html.Append("<div class='conversation-header'><div>");
        html.Append($"<h2>{Enc(title)}</h2>");
        html.Append($"<div class='conversation-people'>{Enc(people)}</div></div>");
        html.Append($"<div class='conversation-count'>{Enc(countText)}</div></div>");
        html.Append("<div class='conversation-messages'>");
        foreach (var message in conversation.Messages)
            AppendMessage(html, message, title, senderClasses, ref nextSender);
        html.Append("</div></section>");
    }

    private static void AppendMessage(
        StringBuilder html,
        TeamsMessage message,
        string conversationTitle,
        Dictionary<string, string> senderClasses,
        ref int nextSender)
    {
        if (!senderClasses.TryGetValue(message.SenderDisplay, out var senderClass))
        {
            senderClass = SenderPalette[nextSender % SenderPalette.Length];
            senderClasses[message.SenderDisplay] = senderClass;
            nextSender++;
        }

        var when = message.ReceivedUtc ?? message.SentUtc;
        var timeText = when?.ToLocalTime().ToString("MMM d, yyyy h:mm tt", CultureInfo.CurrentCulture) ?? "";
        var messageDate = when?.ToLocalTime().ToString("yyyy-MM-dd") ?? "";
        var messageSort = when?.ToUniversalTime().ToString("o") ?? "";
        var others = FormatPeople(string.Join("||",
            TeamsParticipantMatching.Split(message.Participants).Where(name =>
                !name.Equals(message.SenderDisplay, StringComparison.Ordinal))));
        var context = string.IsNullOrWhiteSpace(others)
            ? $"Chat: {conversationTitle} • no other named participants shown"
            : $"Chat: {conversationTitle} • with {others}";
        html.Append($"<article class='message-card sender-{senderClass}' data-sender='{Enc(message.SenderDisplay)}' data-date='{Enc(messageDate)}' data-time='{Enc(messageSort)}'>");
        html.Append("<div class='speaker-row'><span class='speaker-block'>");
        html.Append($"<span class='speaker-name'>{Enc(message.SenderDisplay)}</span>");
        html.Append($"<span class='speaker-context'>{Enc(context)}</span></span>");
        html.Append($"<span class='message-time'>{Enc(timeText)}</span></div>");
        html.Append($"<div class='message-body'>{HtmlBody(message.BodyText)}</div>");
        if (!string.IsNullOrWhiteSpace(message.AttachmentsText))
            html.Append($"<div class='attachments'><strong>Attachments:</strong> {Enc(message.AttachmentsText)}</div>");
        html.Append("<details class='message-details'><summary>Details</summary>");
        html.Append($"<div><strong>Folder:</strong> {Enc(message.FolderPath)}</div>");
        html.Append($"<div><strong>Subject:</strong> {Enc(message.Subject)}</div>");
        html.Append($"<div><strong>Message class:</strong> {Enc(message.MessageClass)}</div>");
        html.Append($"<div><strong>From:</strong> {Enc(message.SenderName)} &lt;{Enc(message.SenderAddress)}&gt;</div>");
        html.Append($"<div><strong>To:</strong> {Enc(message.ToRecipients)}</div>");
        html.Append($"<div><strong>Cc:</strong> {Enc(message.CcRecipients)}</div>");
        html.Append($"<div><strong>Sent:</strong> {Enc(FormatDetailDate(message.SentUtc))}</div>");
        html.Append($"<div><strong>Received:</strong> {Enc(FormatDetailDate(message.ReceivedUtc))}</div>");
        html.Append($"<div><strong>Entry ID:</strong> <span class='wrap'>{Enc(message.EntryId)}</span></div>");
        html.Append("</details></article>");
    }

    private static string HtmlBody(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return "<em>(empty)</em>";
        return Enc(value.Trim()).Replace("\r\n", "<br/>").Replace("\n", "<br/>");
    }

    private static string FormatPeople(string participants)
    {
        var names = TeamsParticipantMatching.Split(participants);
        if (names.Count == 0)
            return "(unknown participants)";
        if (names.Count > 4)
            return string.Join(", ", names.Take(4)) + ", ...";
        return string.Join(", ", names);
    }

    private static string FormatDetailDate(DateTime? value) =>
        value?.ToLocalTime().ToString("g") ?? "";

    private static string SummaryCard(string label, string value) =>
        $"<div class='summary-card'><div class='label'>{Enc(label)}</div><div class='value'>{Enc(value)}</div></div>";

    private static string PersonOption(string name, bool isChecked)
    {
        var check = isChecked ? " checked='checked'" : "";
        return $"<label class='person-option'><input type='checkbox' class='person-check' value='{Enc(name)}'{check}/> <span>{Enc(name)}</span></label>";
    }

    private static string Option(string value, string label, bool selected) =>
        selected
            ? $"<option value='{Enc(value)}' selected='selected'>{Enc(label)}</option>"
            : $"<option value='{Enc(value)}'>{Enc(label)}</option>";

    private static string Enc(string? value) => WebUtility.HtmlEncode(value ?? "");

    private static string LoadCss()
    {
        var assembly = typeof(TeamsReportHtml).Assembly;
        using var stream = assembly.GetManifestResourceStream("EmailReviewViewer.App.Assets.TeamsReport.css")
            ?? throw new InvalidOperationException("Teams report CSS resource is missing.");
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    // ponytail: viewer-only paging + postMessage; visual CSS/markup stay with the HTML report
    private const string Script = """
        (function () {
          const checks = Array.from(document.querySelectorAll('.person-check'));
          const personSearch = document.getElementById('personSearch');
          const messageSearch = document.getElementById('messageSearch');
          const participantMatchMode = document.getElementById('participantMatchMode');
          const startDateFilter = document.getElementById('startDateFilter');
          const endDateFilter = document.getElementById('endDateFilter');
          const sortOrder = document.getElementById('sortOrder');
          const selectAllPeople = document.getElementById('selectAllPeople');
          const selectAllOther = document.getElementById('selectAllOther');
          const clearAll = document.getElementById('clearAll');
          const layout = document.querySelector('.review-layout');
          const resizeHandle = document.getElementById('resizeHandle');
          let debounce = null;
          let posting = false;

          function selectedPeople() { return checks.filter(c => c.checked).map(c => c.value); }
          function post(message) {
            if (!window.chrome || !window.chrome.webview) return;
            posting = true;
            window.chrome.webview.postMessage(JSON.stringify(message));
            posting = false;
          }
          function postQuery() {
            post({
              action: 'query',
              personSearch: personSearch ? personSearch.value : '',
              keyword: messageSearch ? messageSearch.value : '',
              matchMode: participantMatchMode ? participantMatchMode.value : 'messagesFromSelected',
              startDate: startDateFilter ? startDateFilter.value : '',
              endDate: endDateFilter ? endDateFilter.value : '',
              newestFirst: !sortOrder || sortOrder.value !== 'oldestFirst',
              selected: selectedPeople()
            });
          }
          function scheduleQuery() {
            if (posting) return;
            clearTimeout(debounce);
            debounce = setTimeout(postQuery, 350);
          }
          function filterPersonList() {
            const text = (personSearch.value || '').trim().toLowerCase();
            checks.forEach(check => {
              const label = check.closest('.person-option');
              label.style.display = !text || check.value.toLowerCase().includes(text) ? '' : 'none';
            });
          }
          function safeGetLocalStorage(key) {
            try { return localStorage.getItem(key); } catch (_) { return null; }
          }
          function safeSetLocalStorage(key, value) {
            try { localStorage.setItem(key, value); } catch (_) { }
          }
          function setupColumnResize() {
            if (!layout || !resizeHandle) return;
            const savedWidth = safeGetLocalStorage('purviewTeamsReport.leftColumnWidth');
            if (savedWidth) layout.style.setProperty('--left-column-width', savedWidth);
            let dragging = false;
            function setWidthFromPointer(clientX) {
              const rect = layout.getBoundingClientRect();
              const min = 180;
              const max = Math.max(min, Math.min(720, rect.width - 320));
              const width = Math.max(min, Math.min(max, clientX - rect.left));
              const value = Math.round(width) + 'px';
              layout.style.setProperty('--left-column-width', value);
              safeSetLocalStorage('purviewTeamsReport.leftColumnWidth', value);
            }
            resizeHandle.addEventListener('pointerdown', event => { dragging = true; resizeHandle.classList.add('dragging'); resizeHandle.setPointerCapture(event.pointerId); event.preventDefault(); });
            resizeHandle.addEventListener('pointermove', event => { if (dragging) setWidthFromPointer(event.clientX); });
            function stopDragging(event) { if (!dragging) return; dragging = false; resizeHandle.classList.remove('dragging'); try { resizeHandle.releasePointerCapture(event.pointerId); } catch (_) { } }
            resizeHandle.addEventListener('pointerup', stopDragging);
            resizeHandle.addEventListener('pointercancel', stopDragging);
            resizeHandle.addEventListener('keydown', event => {
              if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
              const current = parseInt(getComputedStyle(layout).getPropertyValue('--left-column-width'), 10) || 340;
              const delta = event.key === 'ArrowRight' ? 30 : -30;
              const value = Math.max(180, Math.min(720, current + delta)) + 'px';
              layout.style.setProperty('--left-column-width', value);
              safeSetLocalStorage('purviewTeamsReport.leftColumnWidth', value);
              event.preventDefault();
            });
          }
          function setGroupChecked(selectAllCheckbox, groupId) {
            if (!selectAllCheckbox) return;
            const group = document.getElementById(groupId);
            if (!group) return;
            group.querySelectorAll('.person-check').forEach(c => { c.checked = selectAllCheckbox.checked; });
            postQuery();
          }

          checks.forEach(c => c.addEventListener('change', scheduleQuery));
          if (personSearch) personSearch.addEventListener('input', filterPersonList);
          if (messageSearch) messageSearch.addEventListener('input', scheduleQuery);
          if (participantMatchMode) participantMatchMode.addEventListener('change', postQuery);
          if (startDateFilter) startDateFilter.addEventListener('change', postQuery);
          if (endDateFilter) endDateFilter.addEventListener('change', postQuery);
          if (sortOrder) sortOrder.addEventListener('change', postQuery);
          if (selectAllPeople) selectAllPeople.addEventListener('change', () => setGroupChecked(selectAllPeople, 'peopleBox'));
          if (selectAllOther) selectAllOther.addEventListener('change', () => setGroupChecked(selectAllOther, 'otherPeopleBox'));
          if (clearAll) clearAll.addEventListener('click', () => {
            checks.forEach(c => c.checked = false);
            if (selectAllPeople) selectAllPeople.checked = false;
            if (selectAllOther) selectAllOther.checked = false;
            if (participantMatchMode) participantMatchMode.value = 'messagesFromSelected';
            if (startDateFilter) startDateFilter.value = '';
            if (endDateFilter) endDateFilter.value = '';
            if (sortOrder) sortOrder.value = 'newestFirst';
            if (personSearch) personSearch.value = '';
            if (messageSearch) messageSearch.value = '';
            filterPersonList();
            postQuery();
          });
          const previousPage = document.getElementById('previousPage');
          const nextPage = document.getElementById('nextPage');
          if (previousPage) previousPage.addEventListener('click', () => post({ action: 'page', delta: -1 }));
          if (nextPage) nextPage.addEventListener('click', () => post({ action: 'page', delta: 1 }));
          const openDatabase = document.getElementById('openDatabase');
          if (openDatabase) openDatabase.addEventListener('click', () => post({ action: 'openDatabase' }));
          setupColumnResize();
          filterPersonList();
        })();
        """;
}

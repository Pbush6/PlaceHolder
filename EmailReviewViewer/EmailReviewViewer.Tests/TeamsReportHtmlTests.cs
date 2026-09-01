using System.Globalization;
using EmailReviewViewer.App;
using EmailReviewViewer.App.Data;

namespace EmailReviewViewer.Tests;

public sealed class TeamsReportHtmlTests
{
    [Fact]
    public void Css_uses_the_Teams_HTML_report_colors_fonts_and_shapes()
    {
        var css = TeamsReportHtml.Css;

        Assert.Contains("--bg: #f4f6fb", css);
        Assert.Contains("--accent: #315fbd", css);
        Assert.Contains("font-family: \"Segoe UI\", Arial, sans-serif", css);
        Assert.Contains("linear-gradient(135deg, #213b78, #5c7dde)", css);
        Assert.Contains("border-radius: 18px", css);
        Assert.Contains("border-radius: 16px", css);
        Assert.Contains("border-radius: 14px", css);
        Assert.Contains(".sender-blue { border-left-color: #2f6fec; }", css);
        Assert.Contains(".hero-credit", css);
        Assert.Contains(".hero-open { position: absolute; right: 18px; top: 14px;", css);
        Assert.Contains(".message-card", css);
        Assert.Contains(".conversation-header { display: flex;", css);
    }

    [Fact]
    public void Document_emits_the_Teams_HTML_report_chrome_and_message_cards()
    {
        var html = TeamsReportHtml.Build(SampleModel());

        Assert.Contains("<header class='hero'>", html);
        Assert.Contains("Microsoft Purview eDiscovery Teams Conversation Report", html);
        Assert.Contains("By Patrick Bush", html);
        Assert.Contains("class='summary-grid'", html);
        Assert.Contains("class='filter-panel'", html);
        Assert.Contains("Choose whose conversations to view", html);
        Assert.Contains("id='participantMatchMode'", html);
        Assert.Contains("Exact selected people, ignoring Other/IDs", html);
        Assert.Contains("class='person-check'", html);
        Assert.Contains("Torey Page", html);
        Assert.Contains("class='conversation'", html);
        Assert.Contains("class='message-card sender-blue'", html);
        Assert.Contains("class='speaker-name'", html);
        Assert.Contains("class='message-body'", html);
        Assert.Contains("class='message-details'", html);
        Assert.Contains("Hello Linda.", html);
        Assert.Contains("id='conversationList'", html);
        Assert.Contains("id='openDatabase'", html);
        Assert.Contains("Open Database", html);
        Assert.Contains("class='hero-open'", html);
        Assert.Contains("action: 'openDatabase'", html);
        Assert.DoesNotContain("Teams Review Viewer", html);
    }

    [Fact]
    public void Script_filters_loaded_conversations_in_the_browser_before_asking_the_host()
    {
        var html = TeamsReportHtml.Build(SampleModel());

        Assert.Contains("function applyFilters()", html);
        Assert.Contains("message.hidden = !showMessage", html);
        Assert.Contains("conversation.hidden = !showConversation", html);
        Assert.Contains("applyFilters(); scheduleQuery();", html);
        Assert.Contains("c.addEventListener('change'", html);
    }

    [Theory]
    [InlineData("2026-01-15T12:30:45", "2026-01-15 12:30 CST")]
    [InlineData("2026-07-15T12:30:45", "2026-07-15 12:30 CDT")]
    public void FormatGenerated_omits_seconds_and_uses_a_timezone_abbreviation(string localTime, string expected)
    {
        var when = DateTime.Parse(localTime, CultureInfo.InvariantCulture);
        var central = TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time");

        Assert.Equal(expected, TeamsReportHtml.FormatGenerated(when, central));
    }

    [Theory]
    [InlineData(@"LArtley@perfectionlearning.com.001\TeamsMessagesData", "LArtley@perfectionlearning.com.001.pst")]
    [InlineData(@"SamplePst\TeamsMessagesData", "SamplePst.pst")]
    [InlineData(@"export.pst\TeamsMessagesData", "export.pst")]
    public void ResolvePstName_uses_the_Outlook_store_name_from_folder_paths(string folderPath, string expected)
    {
        var folders = new[] { new FolderCount(folderPath, 4) };

        Assert.Equal(expected, TeamsReportHtml.ResolvePstName(folders, @"C:\temp\LArtley Messages_Teams.db"));
    }

    [Fact]
    public void ResolvePstName_falls_back_to_the_database_file_when_folders_are_missing()
    {
        Assert.Equal(
            "LArtley Messages_Teams.db",
            TeamsReportHtml.ResolvePstName([], @"C:\temp\LArtley Messages_Teams.db"));
    }

    [Fact]
    public void Filter_heading_shows_filtered_totals_for_the_whole_result_not_the_current_page()
    {
        var model = SampleModel() with
        {
            ConversationCount = 872,
            VisibleMessageCount = 13653,
            PageIndex = 0,
            PageSize = 25
        };

        var html = TeamsReportHtml.Build(model);

        Assert.Contains("id='resultCount'", html);
        Assert.Contains("872 conversations / 13,653 messages shown", html);
        Assert.Contains("1–25 of 872", html);
        Assert.DoesNotContain("resultCount.textContent = visibleConversations", html);
    }

    [Fact]
    public void WriteReport_persists_html_larger_than_the_WebView2_NavigateToString_limit()
    {
        var hugeBody = new string('x', 2 * 1024 * 1024);
        var model = SampleModel() with
        {
            Conversations =
            [
                new TeamsConversationHtml(
                    new TeamsConversationListItem("chat-huge", DateTime.UtcNow, "Huge chat", "Torey Page", 1, 1),
                    [
                        new TeamsMessage
                        {
                            SenderDisplay = "Torey Page",
                            Participants = "Torey Page",
                            ConversationTitle = "Huge chat",
                            BodyText = hugeBody
                        }
                    ])
            ]
        };
        var directory = Path.Combine(Path.GetTempPath(), $"teams-html-{Guid.NewGuid():N}");
        try
        {
            var path = TeamsReportHtml.WriteReport(directory, model);
            var info = new FileInfo(path);

            Assert.True(info.Exists);
            Assert.True(info.Length > 2 * 1024 * 1024);
            Assert.Contains("Huge chat", File.ReadAllText(path), StringComparison.Ordinal);
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);
        }
    }

    [Fact]
    public void Details_To_and_Cc_list_ids_and_emails_after_people()
    {
        var model = SampleModel() with
        {
            Conversations =
            [
                new TeamsConversationHtml(
                    new TeamsConversationListItem("chat-1", DateTime.UtcNow, "Bot chat", "Linda Artley||Torey Page", 1, 1),
                    [
                        new TeamsMessage
                        {
                            SenderDisplay = "Torey Page",
                            Participants = "Linda Artley||Torey Page",
                            ConversationTitle = "Bot chat",
                            BodyText = "Hello.",
                            ToRecipients = "28:fd931076-bbfb-4a38-a85c-1f0fb5b61bee; Linda Artley; Torey Page",
                            CcRecipients = "19:meeting-thread; jane@example.com; Linda Artley"
                        }
                    ])
            ]
        };

        var html = TeamsReportHtml.Build(model);

        Assert.Contains(
            "<strong>To:</strong> Linda Artley; Torey Page; 28:fd931076-bbfb-4a38-a85c-1f0fb5b61bee",
            html);
        Assert.Contains(
            "<strong>Cc:</strong> Linda Artley; 19:meeting-thread; jane@example.com",
            html);
        Assert.DoesNotContain(
            "<strong>To:</strong> 28:fd931076-bbfb-4a38-a85c-1f0fb5b61bee; Linda Artley; Torey Page",
            html);
        Assert.DoesNotContain(
            "<strong>Cc:</strong> 19:meeting-thread; jane@example.com; Linda Artley",
            html);
    }

    [Fact]
    public void Empty_body_matches_the_HTML_report_placeholder()
    {
        var model = SampleModel() with
        {
            Conversations =
            [
                new TeamsConversationHtml(
                    new TeamsConversationListItem("chat-1", DateTime.UtcNow, "Empty chat", "Torey Page", 1, 1),
                    [
                        new TeamsMessage
                        {
                            SenderDisplay = "Torey Page",
                            Participants = "Torey Page",
                            ConversationTitle = "Empty chat",
                            BodyText = "   "
                        }
                    ])
            ]
        };

        var html = TeamsReportHtml.Build(model);

        Assert.Contains("<em>(empty)</em>", html);
    }

    private static TeamsReportHtmlModel SampleModel() => new(
        PstName: "sample.pst",
        Generated: "2026-08-28 12:00:00 -05:00",
        MessageCount: 4,
        PeopleCount: 2,
        FolderCount: 1,
        ReadWarnings: "Items: 0; Attachments: 0",
        People: ["Torey Page"],
        OtherNames: ["Meeting Bot"],
        SelectedPeople: [],
        SelectedOtherNames: [],
        Folders: [new FolderCount(@"SamplePst\TeamsMessagesData", 4)],
        Conversations:
        [
            new TeamsConversationHtml(
                new TeamsConversationListItem(
                    "chat-1",
                    new DateTime(2024, 1, 1, 9, 0, 5, DateTimeKind.Utc),
                    "Sample Teams chat",
                    "Linda Artley||Torey Page",
                    1,
                    2),
                [
                    new TeamsMessage
                    {
                        FolderPath = @"SamplePst\TeamsMessagesData",
                        SenderName = "Torey Page",
                        SenderAddress = "torey@example.com",
                        SenderDisplay = "Torey Page",
                        Participants = "Linda Artley||Torey Page",
                        ConversationKey = "chat-1",
                        ConversationTitle = "Sample Teams chat",
                        Subject = "Sample Teams chat",
                        SentUtc = new DateTime(2024, 1, 1, 9, 0, 0, DateTimeKind.Utc),
                        ReceivedUtc = new DateTime(2024, 1, 1, 9, 0, 5, DateTimeKind.Utc),
                        BodyText = "Hello Linda.",
                        MessageClass = "IPM.Note",
                        EntryId = "entry-1",
                        AttachmentsText = "sample.pdf"
                    }
                ])
        ],
        PersonSearch: "",
        Keyword: "",
        MatchMode: TeamsParticipantMatchMode.MessagesFromSelected,
        FromDate: "",
        ToDate: "",
        NewestFirst: true,
        ConversationCount: 1,
        VisibleMessageCount: 1,
        PageIndex: 0,
        PageSize: 25,
        DatabasePath: @"C:\temp\sample_Teams.db");
}

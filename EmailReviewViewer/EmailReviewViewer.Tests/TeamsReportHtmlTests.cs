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

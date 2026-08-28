using EmailReviewViewer.App.Data;
using Microsoft.Data.Sqlite;

namespace EmailReviewViewer.Tests;

public sealed class TeamsReviewFormTests : IDisposable
{
    private readonly string _databasePath = Path.Combine(Path.GetTempPath(), $"teams-ui-{Guid.NewGuid():N}.db");

    [Fact]
    public void TeamsReviewForm_hosts_the_HTML_report_in_a_webview_not_WinForms_cards()
    {
        using var form = new EmailReviewViewer.App.TeamsReviewForm(null);
        form.CreateControl();

        Assert.True(form.Width >= 1000);
        Assert.True(form.Height >= 650);
        Assert.Equal("Teams Review Viewer", form.Text);
        var view = Assert.Single(form.Controls.Find("TeamsReportView", true));
        Assert.Equal(System.Windows.Forms.DockStyle.Fill, view.Dock);
        Assert.Empty(form.Controls.Find("TeamsConversationPane", true));
        Assert.Empty(form.Controls.Find("ApplicationTitle", true));
    }

    [Fact]
    public async Task Teams_switcher_rejects_email_database_without_replacing_current()
    {
        using var teams = await CreateTeamsDatabaseAsync();
        var emailPath = Path.Combine(Path.GetTempPath(), $"email-for-teams-{Guid.NewGuid():N}.db");
        try
        {
            var email = new EmailRepository(emailPath);
            await email.EnsureCreatedAsync();
            email.Dispose();
            teams.Dispose();

            using var switcher = new TeamsDatabaseStoreSwitcher();
            await switcher.SwitchAsync(_databasePath);
            var error = await Assert.ThrowsAsync<InvalidDataException>(() => switcher.SwitchAsync(emailPath));

            Assert.Contains("compatible", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(Path.GetFullPath(_databasePath), switcher.DatabasePath);
        }
        finally
        {
            SqliteConnection.ClearAllPools();
            if (File.Exists(emailPath))
                File.Delete(emailPath);
        }
    }

    private async Task<TeamsRepository> CreateTeamsDatabaseAsync()
    {
        var repository = new TeamsRepository(_databasePath);
        await repository.EnsureCreatedAsync();
        await repository.InsertBatchAsync(
        [
            new TeamsMessage
            {
                Id = 1,
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
                Preview = "Hello",
                BodyText = "Hello",
                MessageClass = "IPM.Note",
                EntryId = "ui-entry-1"
            }
        ]);
        return repository;
    }

    public void Dispose()
    {
        SqliteConnection.ClearAllPools();
        foreach (var candidate in new[] { _databasePath, _databasePath + "-wal", _databasePath + "-shm" })
            if (File.Exists(candidate))
                File.Delete(candidate);
    }
}

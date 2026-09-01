using EmailReviewViewer.App.Data;

namespace EmailReviewViewer.Tests;

public sealed class TeamsParticipantMatchingTests
{
    [Fact]
    public void FormatRecipientsPeopleFirst_lists_28_bots_after_people()
    {
        var formatted = TeamsParticipantMatching.FormatRecipientsPeopleFirst(
            "28:fd931076-bbfb-4a38-a85c-1f0fb5b61bee; Linda Artley; Torey Page");

        Assert.Equal(
            "Linda Artley; Torey Page; 28:fd931076-bbfb-4a38-a85c-1f0fb5b61bee",
            formatted);
    }

    [Fact]
    public void FormatRecipientsPeopleFirst_keeps_multiple_bots_last_in_their_original_order()
    {
        var formatted = TeamsParticipantMatching.FormatRecipientsPeopleFirst(
            "28:aaaa; Linda Artley; 28:bbbb; Torey Page");

        Assert.Equal("Linda Artley; Torey Page; 28:aaaa; 28:bbbb", formatted);
    }

    [Fact]
    public void FormatRecipientsPeopleFirst_lists_other_ids_and_emails_after_people()
    {
        var formatted = TeamsParticipantMatching.FormatRecipientsPeopleFirst(
            "19:meeting-thread; jane@example.com; Linda Artley; 28:aaaa; Torey Page");

        Assert.Equal(
            "Linda Artley; Torey Page; 19:meeting-thread; jane@example.com; 28:aaaa",
            formatted);
    }

    [Theory]
    [InlineData(null, "")]
    [InlineData("", "")]
    [InlineData("   ", "")]
    [InlineData("Linda Artley", "Linda Artley")]
    [InlineData("28:fd931076-bbfb-4a38-a85c-1f0fb5b61bee", "28:fd931076-bbfb-4a38-a85c-1f0fb5b61bee")]
    [InlineData("Linda Artley; Torey Page", "Linda Artley; Torey Page")]
    public void FormatRecipientsPeopleFirst_leaves_simple_lists_intact(string? value, string expected)
    {
        Assert.Equal(expected, TeamsParticipantMatching.FormatRecipientsPeopleFirst(value));
    }
}

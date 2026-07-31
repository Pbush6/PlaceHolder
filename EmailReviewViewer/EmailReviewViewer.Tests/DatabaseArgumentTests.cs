using EmailReviewViewer.App;

namespace EmailReviewViewer.Tests;

public class DatabaseArgumentTests
{
    [Fact]
    public void Resolve_ReturnsFullPath_ForPlainPath()
    {
        var path = Path.Combine(Path.GetTempPath(), "case_Email.db");

        Assert.Equal(Path.GetFullPath(path), DatabaseArgument.Resolve(path));
    }

    [Fact]
    public void Resolve_DecodesProtocolUrl()
    {
        var path = Path.Combine(Path.GetTempPath(), "Review files", "case (final)_Email.db");
        var url = DatabaseArgument.ProtocolScheme + Uri.EscapeDataString(path);

        Assert.Equal(Path.GetFullPath(path), DatabaseArgument.Resolve(url));
    }

    [Fact]
    public void Resolve_AcceptsUppercaseSchemeAndTrailingSlash()
    {
        var path = Path.Combine(Path.GetTempPath(), "case_Email.db");
        var url = "PURVIEW-EMAIL:" + Uri.EscapeDataString(path) + "/";

        Assert.Equal(Path.GetFullPath(path), DatabaseArgument.Resolve(url));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("purview-email:")]
    public void Resolve_RejectsEmptyArguments(string argument)
    {
        Assert.Throws<ArgumentException>(() => DatabaseArgument.Resolve(argument));
    }
}

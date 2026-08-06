namespace EmailReviewViewer.App;

/// <summary>
/// Resolves the database argument the viewer was started with. The conversion dashboard links to a
/// <c>purview-email:</c> URL because a browser cannot hand a plain file path to a desktop program.
/// </summary>
public static class DatabaseArgument
{
    public const string ProtocolScheme = "purview-email:";

    public static string Resolve(string argument)
    {
        if (string.IsNullOrWhiteSpace(argument))
            throw new ArgumentException("A database path is required.", nameof(argument));

        var value = argument.Trim();
        if (value.StartsWith(ProtocolScheme, StringComparison.OrdinalIgnoreCase))
        {
            value = value[ProtocolScheme.Length..].TrimEnd('/');
            value = Uri.UnescapeDataString(value);
        }

        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException("A database path is required.", nameof(argument));

        return Path.GetFullPath(value);
    }
}

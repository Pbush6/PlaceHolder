using System.Text.RegularExpressions;

namespace EmailReviewViewer.App.Data;

public static class TeamsParticipantMatching
{
    private static readonly Regex UnlikelyToken = new(
        @"@|https?://|www\.|/|\\|\d|[{}\[\]_]|\b(thread|communication|meeting|call|teamsvisitor|visitor|unknown|recipient|exchange|admin|onmicrosoft|perfectionlearning\.com)\b",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex NamePart = new(
        @"^[\p{L}][\p{L}'.\-]*$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    public static IReadOnlyList<string> Split(string? value) =>
        string.IsNullOrWhiteSpace(value)
            ? []
            : value.Split("||", StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

    public static bool IsLikelyPersonName(string? value)
    {
        var name = value?.Trim() ?? "";
        if (name.Length == 0 || name.Length > 60)
            return false;
        if (name.StartsWith("Fireflies.ai", StringComparison.OrdinalIgnoreCase))
            return false;
        if (UnlikelyToken.IsMatch(name))
            return false;
        var parts = name.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length is < 2 or > 5)
            return false;
        return parts.All(part => NamePart.IsMatch(part));
    }

    public static bool Matches(
        TeamsParticipantMatchMode mode,
        IReadOnlyList<string> selected,
        IReadOnlyList<string> participants,
        string senderDisplay)
    {
        if (selected.Count == 0)
            return true;

        if (mode == TeamsParticipantMatchMode.MessagesFromSelected)
            return selected.Contains(senderDisplay, StringComparer.Ordinal);

        var people = participants.Where(IsLikelyPersonName).ToArray();
        if (mode == TeamsParticipantMatchMode.Involving)
            return selected.All(participants.Contains);
        if (mode == TeamsParticipantMatchMode.ExactAll)
            return SameSet(participants, selected);
        if (mode == TeamsParticipantMatchMode.ExactPeopleOnly)
        {
            var selectedPeople = selected.Where(people.Contains).ToArray();
            return selected.All(participants.Contains) && SameSet(people, selectedPeople);
        }

        return true;
    }

    public static bool ConversationMatches(
        TeamsParticipantMatchMode mode,
        IReadOnlyList<string> selected,
        IReadOnlyList<string> participants)
    {
        if (selected.Count == 0 || mode == TeamsParticipantMatchMode.MessagesFromSelected)
            return true;
        return Matches(mode, selected, participants, senderDisplay: "");
    }

    private static bool SameSet(IReadOnlyList<string> actual, IReadOnlyList<string> selected) =>
        selected.All(actual.Contains) && actual.Count == selected.Count;
}

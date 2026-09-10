using System.Text.RegularExpressions;

namespace AgentTeamMateBot.Services;

public static class MeetingJoinParser
{
    private static readonly HashSet<string> IgnoredTokens = new(StringComparer.OrdinalIgnoreCase)
    {
        "join", "meeting", "id", "meetingid", "passcode", "password", "or", "and", "the"
    };

    public static bool TryParse(string? text, out string meetingId, out string? passcode)
    {
        meetingId = string.Empty;
        passcode = null;

        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var cleaned = Normalize(text);

        TryFromTeamsMeetUrl(cleaned, out var urlId, out _);
        TryFromLabels(cleaned, out var labeledId, out var labeledPasscode);
        TryFromTokens(cleaned, out var tokenId, out var tokenPasscode);

        meetingId = FirstId(urlId, labeledId, tokenId);
        passcode = FirstPasscode(labeledPasscode, tokenPasscode);

        return meetingId.Length >= 10;
    }

    private static string Normalize(string text)
    {
        var cleaned = text.Replace('\u00a0', ' ');
        cleaned = Regex.Replace(
            cleaned,
            @"\[([^\]]*)\]\((https?://[^)]+)\)",
            "$1 $2");
        cleaned = Regex.Replace(cleaned, "<[^>]+>", " ");
        return cleaned.Trim();
    }

    private static void TryFromTeamsMeetUrl(
        string text,
        out string meetingId,
        out string? passcode)
    {
        meetingId = string.Empty;
        passcode = null;

        var meet = Regex.Match(
            text,
            @"https?://(?:www\.)?teams\.microsoft\.com/meet/(\d{10,20})\?p=[^\s<>""']+",
            RegexOptions.IgnoreCase);
        if (!meet.Success)
        {
            return;
        }

        meetingId = meet.Groups[1].Value;
    }

    private static void TryFromLabels(
        string text,
        out string meetingId,
        out string? passcode)
    {
        meetingId = string.Empty;
        passcode = null;

        var idMatch = Regex.Match(
            text,
            @"Meeting\s*ID\s*:\s*([\d\s]{10,40})",
            RegexOptions.IgnoreCase);
        if (idMatch.Success)
        {
            meetingId = new string(idMatch.Groups[1].Value.Where(char.IsDigit).ToArray());
        }

        var passMatch = Regex.Match(
            text,
            @"Passcode\s*:\s*([A-Za-z0-9]{8})\b",
            RegexOptions.IgnoreCase);
        if (passMatch.Success)
        {
            passcode = passMatch.Groups[1].Value;
        }
    }

    private static void TryFromTokens(
        string text,
        out string meetingId,
        out string? passcode)
    {
        meetingId = string.Empty;
        passcode = null;

        var digitGroups = new List<string>();
        string? trailingPasscode = null;

        foreach (var raw in Regex.Split(text, @"\s+"))
        {
            var token = raw.Trim().Trim(':', ',', ';', '.', '"', '\'');
            if (token.Length == 0 ||
                IgnoredTokens.Contains(token) ||
                token.Contains("://", StringComparison.Ordinal))
            {
                continue;
            }

            if (Regex.IsMatch(token, @"^\d{2,20}$"))
            {
                digitGroups.Add(token);
                continue;
            }

            if (token.Any(char.IsLetter) &&
                Regex.IsMatch(token, @"^[A-Za-z0-9]{8}$"))
            {
                trailingPasscode = token;
            }
        }

        meetingId = string.Concat(digitGroups);
        passcode = trailingPasscode;
    }

    private static string FirstId(params string[] values)
    {
        return values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v) && v.Length >= 10)
            ?? string.Empty;
    }

    private static string? FirstPasscode(params string?[] values)
    {
        return values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));
    }
}

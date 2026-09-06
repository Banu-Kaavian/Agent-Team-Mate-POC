using System.Text.RegularExpressions;

namespace AgentTeamMateBot.Services;

public static class MeetingJoinParser
{
    public static bool TryParse(string? text, out string meetingId, out string? passcode)
    {
        meetingId = string.Empty;
        passcode = null;

        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var cleaned = text.Replace('\u00a0', ' ').Trim();

        var passcodeMatch = Regex.Match(
            cleaned,
            @"(?:passcode|password|pin)\s*[:=]?\s*([A-Za-z0-9]{4,20})",
            RegexOptions.IgnoreCase);
        if (passcodeMatch.Success)
        {
            passcode = passcodeMatch.Groups[1].Value;
        }

        var spacedId = Regex.Match(cleaned, @"\b(\d{3}\s+\d{3}\s+\d{3}\s+\d{3,5})\b");
        var compactId = Regex.Match(cleaned, @"\b(\d{12,17})\b");

        if (spacedId.Success)
        {
            meetingId = new string(spacedId.Value.Where(char.IsDigit).ToArray());
        }
        else if (compactId.Success)
        {
            meetingId = compactId.Value;
        }
        else
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(passcode))
        {
            var tokens = Regex.Split(cleaned, @"\s+");
            var last = tokens.LastOrDefault();
            if (!string.IsNullOrWhiteSpace(last) &&
                last.Any(char.IsLetter) &&
                Regex.IsMatch(last, @"^[A-Za-z0-9]{6,12}$"))
            {
                passcode = last;
            }
        }

        return meetingId.Length >= 10;
    }
}

using System.Text.RegularExpressions;

namespace AgentTeamMateBot.Services;

public static class WakeWordDetector
{
    private static readonly Regex AgentInvocationPattern =
        new(
            @"\b(?:hey\s+)?(?:agent\s+)?nova\b|" +
            @"\bajanova\b|" +
            @"\bage(?:nt)?\s+(?:nova|nover|noble|nova's)\b",
            RegexOptions.IgnoreCase |
            RegexOptions.CultureInvariant |
            RegexOptions.Compiled);

    private static readonly Regex LeftoverGreetingPattern =
        new(
            @"^(?:hey|hi|hello)(?:\s+there)?\s*[,:]?\s*",
            RegexOptions.IgnoreCase |
            RegexOptions.CultureInvariant |
            RegexOptions.Compiled);

    public static bool IsAgentInvocation(
        string? recognizedText)
    {
        if (string.IsNullOrWhiteSpace(recognizedText))
        {
            return false;
        }

        return AgentInvocationPattern.IsMatch(recognizedText);
    }

    public static string RemoveActivationPhrase(
        string recognizedText)
    {
        if (string.IsNullOrWhiteSpace(recognizedText))
        {
            return recognizedText;
        }

        var stripped =
            AgentInvocationPattern.Replace(
                recognizedText,
                " ",
                1);

        stripped =
            LeftoverGreetingPattern.Replace(
                stripped,
                string.Empty,
                1);

        stripped =
            Regex.Replace(
                stripped,
                @"\s+",
                " ");

        stripped =
            stripped.Trim(
                ' ', ',', '.', '!', '?',
                ':', ';', '-', '"', '\'');

        if (string.IsNullOrWhiteSpace(stripped))
        {
            return recognizedText;
        }

        if (char.IsLetter(stripped[0]) &&
            char.IsLower(stripped[0]))
        {
            stripped =
                char.ToUpperInvariant(stripped[0]) +
                stripped[1..];
        }

        return stripped;
    }
}

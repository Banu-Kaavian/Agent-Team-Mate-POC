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

    private static readonly Regex LeaveMeetingPattern =
        new(
            @"\b(?:please\s+)?(?:" +
            @"quit|exit|leave|disconnect|hang\s*up|log\s*off|logoff|logout|log\s*out|" +
            @"sign\s*off|sign\s*out|go\s+away|you\s+can\s+(?:go|leave|quit)|" +
            @"get\s+out|end\s+(?:the\s+)?(?:call|meeting)|bye(?:\s+bye)?|goodbye" +
            @")\b",
            RegexOptions.IgnoreCase |
            RegexOptions.CultureInvariant |
            RegexOptions.Compiled);

    // Phrases that remain after stripping "Agent Nova" but are not a real ask.
    private static readonly Regex FillerOnlyPattern =
        new(
            @"^(?:" +
            @"(?:yeah|yes|yep|yup|ok|okay|alright|all\s+right|sure|thanks|thank\s+you|" +
            @"thank\s+you\s+so\s+much|hi|hello|hey|there|please|um|uh|hmm|right|cool|" +
            @"great|perfect|got\s+it|understood|i\s+see|nice|good)+" +
            @"(?:\s+|,|\.|!|\?|;|:|-)*" +
            @")+$",
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

    public static bool IsLeaveMeetingRequest(
        string? recognizedText)
    {
        if (string.IsNullOrWhiteSpace(recognizedText))
        {
            return false;
        }

        if (!IsAgentInvocation(recognizedText))
        {
            return false;
        }

        return LeaveMeetingPattern.IsMatch(recognizedText);
    }

    /// <summary>
    /// True when the utterance names Agent Nova and includes a real request,
    /// not just a mention or filler like "Yeah, Agent Nova."
    /// </summary>
    public static bool IsActionableRequest(
        string? recognizedText)
    {
        if (!IsAgentInvocation(recognizedText))
        {
            return false;
        }

        if (IsLeaveMeetingRequest(recognizedText))
        {
            return true;
        }

        var question = RemoveActivationPhrase(recognizedText!);
        if (string.IsNullOrWhiteSpace(question))
        {
            return false;
        }

        // RemoveActivationPhrase falls back to the full utterance when empty;
        // treat that as non-actionable unless leave was already matched.
        if (string.Equals(
                question.Trim(),
                recognizedText!.Trim(),
                StringComparison.OrdinalIgnoreCase) &&
            AgentInvocationPattern.IsMatch(question) &&
            !HasExtraContentBeyondWakeWord(question))
        {
            return false;
        }

        var normalized = Regex.Replace(question, @"\s+", " ").Trim();
        if (normalized.Length < 4)
        {
            return false;
        }

        if (FillerOnlyPattern.IsMatch(normalized))
        {
            return false;
        }

        // Strip filler prefixes like "Yeah, OK, thank you, ..." then re-check.
        var withoutLeadingFiller = Regex.Replace(
            normalized,
            @"^(?:yeah|yes|yep|yup|ok|okay|alright|sure|thanks|thank you|hi|hello|hey|um|uh)(?:\s+|,|\.|!|\?)*",
            "",
            RegexOptions.IgnoreCase).Trim();

        withoutLeadingFiller = withoutLeadingFiller.Trim(' ', ',', '.', '!', '?', ':', ';');

        if (string.IsNullOrWhiteSpace(withoutLeadingFiller) ||
            withoutLeadingFiller.Length < 4 ||
            FillerOnlyPattern.IsMatch(withoutLeadingFiller))
        {
            return false;
        }

        return true;
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
            return string.Empty;
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

    private static bool HasExtraContentBeyondWakeWord(string text)
    {
        var withoutWake = AgentInvocationPattern.Replace(text, " ");
        withoutWake = Regex.Replace(withoutWake, @"[\s,\.!?;:\-]+", " ").Trim();
        return withoutWake.Length >= 4 && !FillerOnlyPattern.IsMatch(withoutWake);
    }
}

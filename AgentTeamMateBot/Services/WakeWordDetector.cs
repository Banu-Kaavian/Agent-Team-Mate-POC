using System.Text.RegularExpressions;

namespace AgentTeamMateBot.Services;

public static class WakeWordDetector
{
    private static readonly Regex AgentInvocationPattern =
        new(
            @"\b(?:hey|hi|hello)\s+(?:there\s+)?(?:agent\s+)?(?:nova|nover|noble|noha|nola|nava|noah|nower|nova's)\b|" +
            @"\b(?:agent\s+)?(?:nova|nover|noble|noha|nola|nava|noah|nower|nova's)\b|" +
            @"\bajanova\b|" +
            @"\bage(?:nt)?\s+(?:nova|nover|noble|noha|nola|nava|noah|nower|nova's)\b",
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
            @"quit|exit|leave|left|leaving|leaved|disconnect|hang\s*up|" +
            @"log\s*off|logoff|logout|log\s*out|" +
            @"sign\s*off|sign\s*out|go\s+away|" +
            @"you\s+(?:can|may|should)\s+(?:go|leave|quit|exit)|" +
            @"get\s+out|move\s+out|drop\s+(?:off|out)|" +
            @"remove\s+(?:yourself|the\s+bot)|" +
            @"end\s+(?:the\s+)?(?:call|meeting)|bye(?:\s+bye)?|goodbye|" +
            @"(?:left|let|leave)\s+(?:the\s+)?(?:meeting|call)|" +
            @"off\s+(?:the\s+)?(?:call|meeting)" +
            @")\b",
            RegexOptions.IgnoreCase |
            RegexOptions.CultureInvariant |
            RegexOptions.Compiled);

    private static readonly Regex StopSpeakingPattern =
        new(
            @"\b(?:stop(?:\s+talking)?|quiet|be\s+quiet|that's\s+enough|thats\s+enough|hold\s+on)\b",
            RegexOptions.IgnoreCase |
            RegexOptions.CultureInvariant |
            RegexOptions.Compiled);

    private static readonly Regex SkipWorkflowPattern =
        new(
            @"\b(?:do\s+not|don't|dont|never|skip|without|not\s+to|no\s+need\s+to)\s+" +
            @"(?:to\s+)?" +
            @"(?:call(?:ing)?|trigger(?:ing)?|invoke(?:ing)?|run(?:ning)?|start(?:ing)?|send(?:ing)?|export(?:ing)?|post(?:ing)?|fire)\s+" +
            @"(?:to\s+)?(?:the\s+)?(?:workflow|work\s*flow|logic\s*app|export)\b|" +
            @"\b(?:do\s+not|don't|dont|never)\s+trigger(?:ing)?\b|" +
            @"\bnot\s+to\s+trigger(?:ing)?\b|" +
            @"\bno\s+(?:need\s+to\s+)?trigger(?:ing)?\b|" +
            @"\b(?:don't|do\s+not|dont)\s+(?:send|export|post)\b|" +
            @"\bno\s+workflow\b|" +
            @"\bskip\s+(?:the\s+)?(?:workflow|work\s*flow|export)\b|" +
            @"\bleave\s+without\s+(?:the\s+)?(?:workflow|trigger(?:ing)?)\b",
            RegexOptions.IgnoreCase |
            RegexOptions.CultureInvariant |
            RegexOptions.Compiled);

    private static readonly Regex SpokenRecapPattern =
        new(
            @"\b(?:full\s+)?(?:summary|summarize|summarise|summarization|recap)\b|" +
            @"\bprd\b|" +
            @"\bproduct\s+requirements?\b|" +
            @"\bmeeting\s+(?:notes|document|summary|recap)\b",
            RegexOptions.IgnoreCase |
            RegexOptions.CultureInvariant |
            RegexOptions.Compiled);

    private static readonly Regex WorkflowSendPattern =
        new(
            @"\bsend\s+(?:it\s+|the\s+)?(?:to\s+)?(?:the\s+)?workflow\b|" +
            @"\bcall(?:ing)?\s+(?:the\s+)?workflow\b|" +
            @"\bexport\s+(?:to\s+)?(?:the\s+)?workflow\b|" +
            @"\bpost\s+(?:to\s+)?(?:the\s+)?workflow\b|" +
            @"\bsend\s+(?:the\s+)?(?:meeting\s+)?(?:transcript|notes|summary)\s+to\b|" +
            @"\bsend\s+(?:the\s+)?transcript\b",
            RegexOptions.IgnoreCase |
            RegexOptions.CultureInvariant |
            RegexOptions.Compiled);

    private static readonly Regex AffirmativePattern =
        new(
            @"^\s*(?:yes|yeah|yep|yup|ok|okay|sure|please|go\s+ahead|send\s+it|do\s+it|confirm)" +
            @"(?:\s+please)?(?:\s|,|\.|!|\?)*\s*$",
            RegexOptions.IgnoreCase |
            RegexOptions.CultureInvariant |
            RegexOptions.Compiled);

    private static readonly Regex NegativePattern =
        new(
            @"^\s*(?:no|nope|don't|dont|do\s+not|not\s+now|skip(?:\s+it)?|cancel)" +
            @"(?:\s+please)?(?:\s|,|\.|!|\?)*\s*$",
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

    public static bool IsStopSpeakingRequest(string? recognizedText)
    {
        if (string.IsNullOrWhiteSpace(recognizedText) ||
            !IsAgentInvocation(recognizedText))
        {
            return false;
        }

        return StopSpeakingPattern.IsMatch(recognizedText);
    }

    public static bool IsSpokenRecapRequest(string? recognizedText)
    {
        if (string.IsNullOrWhiteSpace(recognizedText))
        {
            return false;
        }

        return SpokenRecapPattern.IsMatch(recognizedText);
    }

    public static bool IsWorkflowSendRequest(string? recognizedText)
    {
        if (string.IsNullOrWhiteSpace(recognizedText))
        {
            return false;
        }

        if (IsSkipWorkflowRequest(recognizedText))
        {
            return false;
        }

        return WorkflowSendPattern.IsMatch(recognizedText);
    }

    public static bool IsShortAffirmative(string? recognizedText)
    {
        return IsShortReply(recognizedText, AffirmativePattern);
    }

    public static bool IsShortNegative(string? recognizedText)
    {
        return IsSkipWorkflowRequest(recognizedText) ||
               IsShortReply(recognizedText, NegativePattern);
    }

    private static bool IsShortReply(string? recognizedText, Regex pattern)
    {
        if (string.IsNullOrWhiteSpace(recognizedText) ||
            recognizedText.Trim().Length > 80)
        {
            return false;
        }

        var stripped = RemoveActivationPhrase(recognizedText);
        if (string.IsNullOrWhiteSpace(stripped))
        {
            stripped = recognizedText;
        }

        return pattern.IsMatch(stripped.Trim());
    }

    public static bool IsSkipWorkflowRequest(string? recognizedText)
    {
        if (string.IsNullOrWhiteSpace(recognizedText))
        {
            return false;
        }

        return SkipWorkflowPattern.IsMatch(recognizedText);
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

        if (IsLeaveMeetingRequest(recognizedText) ||
            IsSpokenRecapRequest(recognizedText) ||
            IsStopSpeakingRequest(recognizedText) ||
            IsSkipWorkflowRequest(recognizedText) ||
            IsWorkflowSendRequest(recognizedText))
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
                " ");

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

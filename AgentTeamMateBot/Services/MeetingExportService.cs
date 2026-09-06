using System.Net.Http.Json;
using System.Text;
using System.Text.RegularExpressions;

namespace AgentTeamMateBot.Services;

public class MeetingExportService
{
    private static readonly Regex TitleLinePattern = new(
        @"^\s*TITLE\s*:\s*(.+)\s*$",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Multiline);

    private readonly IConfiguration _configuration;
    private readonly MeetingContextService _meetingContextService;
    private readonly AiResponseService _aiResponseService;
    private readonly HttpClient _httpClient;

    public MeetingExportService(
        IConfiguration configuration,
        MeetingContextService meetingContextService,
        AiResponseService aiResponseService,
        IHttpClientFactory httpClientFactory)
    {
        _configuration = configuration;
        _meetingContextService = meetingContextService;
        _aiResponseService = aiResponseService;
        _httpClient = httpClientFactory.CreateClient(nameof(MeetingExportService));
        _httpClient.Timeout = TimeSpan.FromSeconds(90);
    }

    public async Task<string> ExportMeetingSummaryAsync(string callId)
    {
        var liveTranscript = _meetingContextService.GetLiveTranscript(callId);
        if (string.IsNullOrWhiteSpace(liveTranscript))
        {
            BotLog.Info("Meeting export skipped: no live transcript yet.");
            return "I do not have enough meeting notes yet to send a summary.";
        }

        var logicAppUrl = _configuration["MeetingExport:LogicAppUrl"];
        if (string.IsNullOrWhiteSpace(logicAppUrl))
        {
            BotLog.Info("Meeting export failed: MeetingExport:LogicAppUrl is missing.");
            return "The meeting export URL is not configured.";
        }

        var document = await _aiResponseService.GenerateMeetingDocumentAsync(
            callId,
            liveTranscript);

        var rawDocument = string.IsNullOrWhiteSpace(document)
            ? liveTranscript
            : document;

        var (title, body) = SplitTitleAndBody(rawDocument, liveTranscript);
        var fileName = BuildFileName(title);
        var transcriptPayload = $"{title}\n\n{body}";

        // content/fileName/title are for SharePoint Create file mapping.
        // transcript stays for the original trigger contract.
        var payload = new Dictionary<string, string>
        {
            ["title"] = title,
            ["fileName"] = fileName,
            ["transcript"] = transcriptPayload,
            ["content"] = transcriptPayload
        };

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" MEETING EXPORT");
        Console.WriteLine("================================================");
        Console.WriteLine($"Call ID   : {callId}");
        Console.WriteLine($"Title     : {title}");
        Console.WriteLine($"File name : {fileName}");
        Console.WriteLine($"Chars     : {transcriptPayload.Length}");

        try
        {
            using var response = await _httpClient.PostAsJsonAsync(
                logicAppUrl,
                payload);

            var responseBody = await response.Content.ReadAsStringAsync();

            Console.WriteLine($"Status    : {(int)response.StatusCode} {response.StatusCode}");
            if (!string.IsNullOrWhiteSpace(responseBody))
            {
                Console.WriteLine(
                    responseBody.Length > 500 ? responseBody[..500] : responseBody);
            }

            Console.WriteLine("================================================");

            if (!response.IsSuccessStatusCode)
            {
                BotLog.Info($"Error: Meeting export HTTP {(int)response.StatusCode}.");
                return "I could not send the meeting summary. Please try again.";
            }

            BotLog.Info($"Meeting summary posted as {fileName}.");
            return $"I sent the {title} summary to your workflow.";
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[MEETING EXPORT] {ex.Message}");
            Console.WriteLine("================================================");
            BotLog.Info($"Error: Meeting export failed. {ex.Message}");
            return "I could not send the meeting summary. Please try again.";
        }
    }

    private static (string Title, string Body) SplitTitleAndBody(
        string document,
        string liveTranscript)
    {
        var match = TitleLinePattern.Match(document);
        if (match.Success)
        {
            var title = CleanTitle(match.Groups[1].Value);
            var body = TitleLinePattern.Replace(document, string.Empty, 1).Trim();
            if (string.IsNullOrWhiteSpace(body))
            {
                body = liveTranscript.Trim();
            }

            return (title, body);
        }

        return (CleanTitle(GuessTitle(liveTranscript)), document.Trim());
    }

    private static string GuessTitle(string liveTranscript)
    {
        var firstLine = liveTranscript
            .Replace('\r', '\n')
            .Split('\n', StringSplitOptions.RemoveEmptyEntries)
            .Select(line => line.Trim())
            .FirstOrDefault(line => line.Length >= 8);

        if (string.IsNullOrWhiteSpace(firstLine))
        {
            return "Meeting Summary";
        }

        var words = Regex.Split(firstLine, @"\s+")
            .Where(word => word.Length > 0)
            .Take(8);
        return string.Join(' ', words);
    }

    private static string CleanTitle(string title)
    {
        var cleaned = Regex.Replace(title ?? string.Empty, @"\s+", " ").Trim();
        cleaned = cleaned.Trim(' ', '.', ',', ':', ';', '-', '"', '\'');
        if (string.IsNullOrWhiteSpace(cleaned))
        {
            return "Meeting Summary";
        }

        if (cleaned.Length > 80)
        {
            cleaned = cleaned[..80].Trim();
        }

        return cleaned;
    }

    private static string BuildFileName(string title)
    {
        var builder = new StringBuilder();
        foreach (var ch in title)
        {
            if (char.IsLetterOrDigit(ch))
            {
                builder.Append(ch);
                continue;
            }

            if (ch is ' ' or '-' or '_')
            {
                if (builder.Length > 0 && builder[^1] != '_')
                {
                    builder.Append('_');
                }
            }
        }

        var slug = builder.ToString().Trim('_');
        if (string.IsNullOrWhiteSpace(slug))
        {
            slug = "Meeting_Summary";
        }

        if (slug.Length > 60)
        {
            slug = slug[..60].Trim('_');
        }

        return $"{slug}_{DateTime.Now:yyyy-MM-dd}.txt";
    }
}

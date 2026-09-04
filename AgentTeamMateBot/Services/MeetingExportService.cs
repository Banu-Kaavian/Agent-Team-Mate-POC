using System.Net.Http.Json;

namespace AgentTeamMateBot.Services;

public class MeetingExportService
{
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

        var transcriptPayload = string.IsNullOrWhiteSpace(document)
            ? liveTranscript
            : document;

        var payload = new Dictionary<string, string>
        {
            ["transcript"] = transcriptPayload
        };

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" MEETING EXPORT");
        Console.WriteLine("================================================");
        Console.WriteLine($"Call ID : {callId}");
        Console.WriteLine($"Chars   : {transcriptPayload.Length}");

        try
        {
            using var response = await _httpClient.PostAsJsonAsync(
                logicAppUrl,
                payload);

            var body = await response.Content.ReadAsStringAsync();

            Console.WriteLine($"Status  : {(int)response.StatusCode} {response.StatusCode}");
            if (!string.IsNullOrWhiteSpace(body))
            {
                Console.WriteLine(body.Length > 500 ? body[..500] : body);
            }

            Console.WriteLine("================================================");

            if (!response.IsSuccessStatusCode)
            {
                BotLog.Info($"Error: Meeting export HTTP {(int)response.StatusCode}.");
                return "I could not send the meeting summary. Please try again.";
            }

            BotLog.Info("Meeting summary posted to Logic App.");
            return "I sent the full meeting summary to your workflow.";
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[MEETING EXPORT] {ex.Message}");
            Console.WriteLine("================================================");
            BotLog.Info($"Error: Meeting export failed. {ex.Message}");
            return "I could not send the meeting summary. Please try again.";
        }
    }
}

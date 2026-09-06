using System.Net.Http.Json;
using System.Text.Json;

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

        var result = await ExportFromNotesAsync(callId, liveTranscript);
        return result.Message;
    }

    public async Task<MeetingExportResult> ExportFromNotesAsync(
        string callId,
        string notes)
    {
        var logicAppUrl = _configuration["MeetingExport:LogicAppUrl"];
        if (string.IsNullOrWhiteSpace(logicAppUrl))
        {
            BotLog.Info("Meeting export failed: MeetingExport:LogicAppUrl is missing.");
            return new MeetingExportResult(
                false,
                0,
                null,
                "The meeting export URL is not configured.");
        }

        var document = await _aiResponseService.GenerateMeetingDocumentAsync(
            callId,
            notes);

        var transcript = string.IsNullOrWhiteSpace(document)
            ? notes
            : document;

        var payload = new Dictionary<string, string>
        {
            ["transcript"] = transcript
        };

        var requestJson = JsonSerializer.Serialize(payload);

        Console.WriteLine("MEETING EXPORT REQUEST BODY:");
        Console.WriteLine(requestJson);

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
                return new MeetingExportResult(
                    false,
                    (int)response.StatusCode,
                    requestJson,
                    "I could not send the meeting summary. Please try again.");
            }

            BotLog.Info("Meeting summary posted to Logic App.");
            return new MeetingExportResult(
                true,
                (int)response.StatusCode,
                requestJson,
                "I sent the full meeting summary to your workflow.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[MEETING EXPORT] {ex.Message}");
            Console.WriteLine("================================================");
            BotLog.Info($"Error: Meeting export failed. {ex.Message}");
            return new MeetingExportResult(
                false,
                0,
                requestJson,
                "I could not send the meeting summary. Please try again.");
        }
    }
}

public record MeetingExportResult(
    bool Succeeded,
    int StatusCode,
    string? RequestJson,
    string Message);

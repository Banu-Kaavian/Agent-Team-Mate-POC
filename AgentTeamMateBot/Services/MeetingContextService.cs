using System.Collections.Concurrent;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Graph.Models;

namespace AgentTeamMateBot.Services;

public class MeetingContextService
{
    private const int DefaultCacheSeconds = 45;
    private const int DefaultMaxCharacters = 8000;
    private const int LiveTranscriptMaxCharacters = 20000;

    private readonly IConfiguration _configuration;
    private readonly GraphAuthService _graphAuthService;
    private readonly HttpClient _httpClient = new();
    private readonly ConcurrentDictionary<string, CallMeetingState> _calls = new();
    private readonly ConcurrentDictionary<string, LiveTranscriptState> _liveTranscripts = new();

    public MeetingContextService(
        IConfiguration configuration,
        GraphAuthService graphAuthService)
    {
        _configuration = configuration;
        _graphAuthService = graphAuthService;
    }

    public void RegisterCall(
        string callId,
        string tenantId,
        string? joinMeetingId,
        string? organizerUserId = null,
        string? joinWebUrl = null)
    {
        if (string.IsNullOrWhiteSpace(callId))
        {
            return;
        }

        var state = _calls.GetOrAdd(callId, _ => new CallMeetingState(callId));

        lock (state.Gate)
        {
            state.TenantId = NullIfEmpty(tenantId) ?? state.TenantId;
            state.JoinMeetingId = NullIfEmpty(joinMeetingId) ?? state.JoinMeetingId;
            state.OrganizerUserId = NullIfEmpty(organizerUserId) ?? state.OrganizerUserId;
            state.JoinWebUrl = NullIfEmpty(joinWebUrl) ?? state.JoinWebUrl;
        }

        Console.WriteLine();
        Console.WriteLine("[MEETING CONTEXT] Call mapped");
        Console.WriteLine($"Call ID         : {callId}");
        Console.WriteLine($"Join meeting ID : {state.JoinMeetingId ?? "(unknown)"}");
        Console.WriteLine($"Organizer ID    : {state.OrganizerUserId ?? "(not yet known)"}");
    }

    public void EnrichFromCallResource(
        string callId,
        Call? call)
    {
        if (string.IsNullOrWhiteSpace(callId) || call == null)
        {
            return;
        }

        var state = _calls.GetOrAdd(callId, _ => new CallMeetingState(callId));

        lock (state.Gate)
        {
            state.TenantId = NullIfEmpty(call.TenantId) ?? state.TenantId;
            state.ThreadId = NullIfEmpty(call.ChatInfo?.ThreadId) ?? state.ThreadId;

            if (call.MeetingInfo is JoinMeetingIdMeetingInfo joinInfo)
            {
                state.JoinMeetingId =
                    NullIfEmpty(joinInfo.JoinMeetingId) ?? state.JoinMeetingId;
            }

            if (call.MeetingInfo is OrganizerMeetingInfo organizerInfo)
            {
                var organizerId =
                    NullIfEmpty(organizerInfo.Organizer?.User?.Id);

                if (!string.IsNullOrWhiteSpace(organizerId) &&
                    !string.Equals(
                        state.OrganizerUserId,
                        organizerId,
                        StringComparison.OrdinalIgnoreCase))
                {
                    state.OrganizerUserId = organizerId;
                    state.CachedAt = null;
                    state.CachedText = null;
                }
            }
        }
    }

    public async Task<string?> GetMeetingContextAsync(string callId)
    {
        if (string.IsNullOrWhiteSpace(callId))
        {
            return null;
        }

        var state = _calls.GetOrAdd(callId, _ => new CallMeetingState(callId));
        var cacheSeconds = _configuration.GetValue(
            "MeetingContext:CacheSeconds",
            DefaultCacheSeconds);

        await state.FetchLock.WaitAsync();

        try
        {
            if (state.CachedAt.HasValue &&
                DateTimeOffset.UtcNow - state.CachedAt.Value <
                    TimeSpan.FromSeconds(cacheSeconds))
            {
                if (!string.IsNullOrWhiteSpace(state.CachedText))
                {
                    LogContext(callId, true, state.CachedText.Length);
                    return state.CachedText;
                }

                LogContext(callId, false, 0);
                Console.WriteLine(
                    "[MEETING CONTEXT] Transcript not available. Continuing without transcript.");
                return null;
            }

            var transcript = await TryRetrieveTranscriptAsync(state);

            state.CachedAt = DateTimeOffset.UtcNow;
            state.CachedText = transcript;

            if (string.IsNullOrWhiteSpace(transcript))
            {
                LogContext(callId, false, 0);
                Console.WriteLine(
                    "[MEETING CONTEXT] Transcript not available. Continuing without transcript.");
                return null;
            }

            LogContext(callId, true, transcript.Length);
            return transcript;
        }
        catch (Exception ex)
        {
            Console.WriteLine(
                $"[MEETING CONTEXT] Retrieval failed: {ex.Message}");
            Console.WriteLine(
                "[MEETING CONTEXT] Transcript not available. Continuing without transcript.");
            LogContext(callId, false, 0);
            return null;
        }
        finally
        {
            state.FetchLock.Release();
        }
    }

    public string AppendLiveTranscript(string callId, string recognizedText)
    {
        if (string.IsNullOrWhiteSpace(callId) ||
            string.IsNullOrWhiteSpace(recognizedText))
        {
            return string.Empty;
        }

        var state = _liveTranscripts.GetOrAdd(callId, _ => new LiveTranscriptState());

        lock (state.Gate)
        {
            if (state.Text.Length > 0)
            {
                state.Text.AppendLine();
            }

            state.Text.Append(recognizedText.Trim());

            if (state.Text.Length > LiveTranscriptMaxCharacters)
            {
                var truncated = state.Text.ToString();
                truncated = truncated[(truncated.Length - LiveTranscriptMaxCharacters)..];
                state.Text.Clear();
                state.Text.Append(truncated);
            }

            return state.Text.ToString();
        }
    }

    public string? GetLiveTranscript(string callId)
    {
        if (string.IsNullOrWhiteSpace(callId) ||
            !_liveTranscripts.TryGetValue(callId, out var state))
        {
            return null;
        }

        lock (state.Gate)
        {
            var text = state.Text.ToString();
            return string.IsNullOrWhiteSpace(text) ? null : text;
        }
    }

    public void Clear(string callId)
    {
        if (string.IsNullOrWhiteSpace(callId))
        {
            return;
        }

        if (_calls.TryRemove(callId, out var state))
        {
            state.FetchLock.Dispose();
            Console.WriteLine(
                $"[MEETING CONTEXT] Cleared for call {callId}");
        }

        if (_liveTranscripts.TryRemove(callId, out _))
        {
            Console.WriteLine(
                $"[MEETING TRANSCRIPT] Cleared for call {callId}");
        }
    }

    private async Task<string?> TryRetrieveTranscriptAsync(CallMeetingState state)
    {
        if (string.IsNullOrWhiteSpace(state.OrganizerUserId))
        {
            Console.WriteLine(
                "[MEETING CONTEXT] Meeting cannot be resolved. Organizer user ID is not available.");
            return null;
        }

        if (string.IsNullOrWhiteSpace(state.OnlineMeetingId))
        {
            state.OnlineMeetingId = await ResolveOnlineMeetingIdAsync(state);
        }

        JsonElement? transcripts = null;

        if (!string.IsNullOrWhiteSpace(state.OnlineMeetingId))
        {
            transcripts = await ListTranscriptsAsync(
                state.OrganizerUserId,
                state.OnlineMeetingId);
        }

        if (transcripts == null)
        {
            transcripts = await ListTranscriptsByOrganizerAsync(state);
        }

        if (transcripts == null)
        {
            return null;
        }

        var transcriptId = SelectTranscriptId(transcripts.Value, state.CallId);
        if (string.IsNullOrWhiteSpace(transcriptId))
        {
            return null;
        }

        if (string.IsNullOrWhiteSpace(state.OnlineMeetingId))
        {
            state.OnlineMeetingId = FindMeetingIdOnTranscript(
                transcripts.Value,
                transcriptId);
        }

        if (string.IsNullOrWhiteSpace(state.OnlineMeetingId))
        {
            Console.WriteLine(
                "[MEETING CONTEXT] Meeting cannot be resolved.");
            return null;
        }

        var content = await DownloadTranscriptContentAsync(
            state.OrganizerUserId,
            state.OnlineMeetingId,
            transcriptId);

        if (string.IsNullOrWhiteSpace(content))
        {
            return null;
        }

        var plainText = ConvertTranscriptToPlainText(content);
        return TruncateToRecent(plainText, GetMaxCharacters());
    }

    private async Task<string?> ResolveOnlineMeetingIdAsync(CallMeetingState state)
    {
        if (string.IsNullOrWhiteSpace(state.OrganizerUserId))
        {
            return null;
        }

        if (!string.IsNullOrWhiteSpace(state.JoinMeetingId))
        {
            var byJoinId = await QueryOnlineMeetingsAsync(
                state.OrganizerUserId,
                $"joinMeetingIdSettings/joinMeetingId eq '{state.JoinMeetingId}'");

            if (!string.IsNullOrWhiteSpace(byJoinId))
            {
                Console.WriteLine(
                    "[MEETING CONTEXT] onlineMeeting resolved by joinMeetingId.");
                return byJoinId;
            }
        }

        if (!string.IsNullOrWhiteSpace(state.JoinWebUrl))
        {
            var encodedUrl = state.JoinWebUrl.Replace("'", "''");
            var byUrl = await QueryOnlineMeetingsAsync(
                state.OrganizerUserId,
                $"JoinWebUrl eq '{encodedUrl}'");

            if (!string.IsNullOrWhiteSpace(byUrl))
            {
                Console.WriteLine(
                    "[MEETING CONTEXT] onlineMeeting resolved by joinWebUrl.");
                return byUrl;
            }
        }

        Console.WriteLine(
            "[MEETING CONTEXT] Meeting cannot be resolved from join metadata.");
        return null;
    }

    private async Task<string?> QueryOnlineMeetingsAsync(
        string organizerUserId,
        string filter)
    {
        var url =
            "https://graph.microsoft.com/v1.0/users/" +
            Uri.EscapeDataString(organizerUserId) +
            "/onlineMeetings?$filter=" +
            Uri.EscapeDataString(filter);

        var (status, body) = await SendGraphGetAsync(url);
        if (status == HttpStatusCode.Forbidden)
        {
            LogGraphAccessDenied("onlineMeetings");
            return null;
        }

        if (!IsSuccess(status) || string.IsNullOrWhiteSpace(body))
        {
            LogGraphFailure("onlineMeetings", status);
            return null;
        }

        using var document = JsonDocument.Parse(body);
        if (!document.RootElement.TryGetProperty("value", out var values) ||
            values.ValueKind != JsonValueKind.Array)
        {
            return null;
        }

        foreach (var meeting in values.EnumerateArray())
        {
            var id = GetString(meeting, "id");
            if (!string.IsNullOrWhiteSpace(id))
            {
                return id;
            }
        }

        return null;
    }

    private async Task<JsonElement?> ListTranscriptsAsync(
        string organizerUserId,
        string onlineMeetingId)
    {
        var url =
            "https://graph.microsoft.com/v1.0/users/" +
            Uri.EscapeDataString(organizerUserId) +
            "/onlineMeetings/" +
            Uri.EscapeDataString(onlineMeetingId) +
            "/transcripts";

        var (status, body) = await SendGraphGetAsync(url);
        if (status == HttpStatusCode.Forbidden)
        {
            LogGraphAccessDenied("transcripts");
            return null;
        }

        if (status == HttpStatusCode.NotFound)
        {
            return null;
        }

        if (!IsSuccess(status) || string.IsNullOrWhiteSpace(body))
        {
            LogGraphFailure("transcripts", status);
            return null;
        }

        using var document = JsonDocument.Parse(body);
        if (!document.RootElement.TryGetProperty("value", out var values) ||
            values.ValueKind != JsonValueKind.Array)
        {
            return null;
        }

        if (values.GetArrayLength() == 0)
        {
            return null;
        }

        return values.Clone();
    }

    private async Task<JsonElement?> ListTranscriptsByOrganizerAsync(
        CallMeetingState state)
    {
        if (string.IsNullOrWhiteSpace(state.OrganizerUserId))
        {
            return null;
        }

        var start = DateTime.UtcNow.AddHours(-24).ToString("o");
        var url =
            "https://graph.microsoft.com/v1.0/users/" +
            Uri.EscapeDataString(state.OrganizerUserId) +
            "/onlineMeetings/getAllTranscripts(meetingOrganizerUserId='" +
            Uri.EscapeDataString(state.OrganizerUserId) +
            "',startDateTime=" +
            Uri.EscapeDataString(start) +
            ")";

        var (status, body) = await SendGraphGetAsync(url);
        if (status == HttpStatusCode.Forbidden)
        {
            LogGraphAccessDenied("getAllTranscripts");
            return null;
        }

        if (!IsSuccess(status) || string.IsNullOrWhiteSpace(body))
        {
            LogGraphFailure("getAllTranscripts", status);
            return null;
        }

        using var document = JsonDocument.Parse(body);
        if (!document.RootElement.TryGetProperty("value", out var values) ||
            values.ValueKind != JsonValueKind.Array ||
            values.GetArrayLength() == 0)
        {
            return null;
        }

        return values.Clone();
    }

    private async Task<string?> DownloadTranscriptContentAsync(
        string organizerUserId,
        string onlineMeetingId,
        string transcriptId)
    {
        var url =
            "https://graph.microsoft.com/v1.0/users/" +
            Uri.EscapeDataString(organizerUserId) +
            "/onlineMeetings/" +
            Uri.EscapeDataString(onlineMeetingId) +
            "/transcripts/" +
            Uri.EscapeDataString(transcriptId) +
            "/content";

        var content = await DownloadContentAsync(url, "text/vtt");
        if (!string.IsNullOrWhiteSpace(content))
        {
            return content;
        }

        return await DownloadContentAsync(
            url,
            "application/vnd.microsoft.graph.transcript+text");
    }

    private async Task<string?> DownloadContentAsync(
        string url,
        string accept)
    {
        var token = await _graphAuthService.GetAccessTokenAsync();

        using var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        request.Headers.Accept.Clear();
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue(accept));

        using var response = await _httpClient.SendAsync(request);
        if (response.StatusCode == HttpStatusCode.Forbidden)
        {
            var errorBody = await response.Content.ReadAsStringAsync();
            var inner = ReadGraphInnerError(errorBody);
            if (string.Equals(
                    inner,
                    "SpeakerAttributionNotAllowed",
                    StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine(
                    "[MEETING CONTEXT] Speaker attribution is disabled. Retrying unattributed format.");
                return null;
            }

            LogGraphAccessDenied("transcript content");
            return null;
        }

        if (response.StatusCode == HttpStatusCode.Accepted ||
            response.StatusCode == HttpStatusCode.NotFound)
        {
            return null;
        }

        if (!response.IsSuccessStatusCode)
        {
            LogGraphFailure("transcript content", response.StatusCode);
            return null;
        }

        return await response.Content.ReadAsStringAsync();
    }

    private async Task<(HttpStatusCode Status, string Body)> SendGraphGetAsync(
        string url)
    {
        var token = await _graphAuthService.GetAccessTokenAsync();

        using var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

        using var response = await _httpClient.SendAsync(request);
        var body = await response.Content.ReadAsStringAsync();
        return (response.StatusCode, body);
    }

    private static string? SelectTranscriptId(
        JsonElement transcripts,
        string callId)
    {
        JsonElement? matching = null;
        JsonElement? latest = null;
        DateTimeOffset latestTime = DateTimeOffset.MinValue;

        foreach (var item in transcripts.EnumerateArray())
        {
            var itemCallId = GetString(item, "callId");
            if (!string.IsNullOrWhiteSpace(itemCallId) &&
                string.Equals(itemCallId, callId, StringComparison.OrdinalIgnoreCase))
            {
                matching = item;
            }

            var created = GetDate(item, "createdDateTime") ??
                          GetDate(item, "endDateTime") ??
                          DateTimeOffset.MinValue;

            if (created >= latestTime)
            {
                latestTime = created;
                latest = item;
            }
        }

        var selected = matching ?? latest;
        return selected.HasValue ? GetString(selected.Value, "id") : null;
    }

    private static string? FindMeetingIdOnTranscript(
        JsonElement transcripts,
        string transcriptId)
    {
        foreach (var item in transcripts.EnumerateArray())
        {
            if (string.Equals(
                    GetString(item, "id"),
                    transcriptId,
                    StringComparison.OrdinalIgnoreCase))
            {
                return GetString(item, "meetingId");
            }
        }

        return null;
    }

    private int GetMaxCharacters()
    {
        return _configuration.GetValue(
            "MeetingContext:MaxCharacters",
            DefaultMaxCharacters);
    }

    private static string ConvertTranscriptToPlainText(string content)
    {
        if (string.IsNullOrWhiteSpace(content))
        {
            return string.Empty;
        }

        var lines = content.Replace("\r\n", "\n").Split('\n');
        var builder = new StringBuilder();

        foreach (var rawLine in lines)
        {
            var line = rawLine.Trim();
            if (string.IsNullOrWhiteSpace(line) ||
                line.Equals("WEBVTT", StringComparison.OrdinalIgnoreCase) ||
                line.StartsWith("NOTE", StringComparison.OrdinalIgnoreCase) ||
                line.Contains("-->") ||
                int.TryParse(line, out _))
            {
                continue;
            }

            if (line.StartsWith("<v ", StringComparison.OrdinalIgnoreCase))
            {
                var close = line.IndexOf('>');
                if (close > 3 && close < line.Length - 1)
                {
                    var speaker = line.Substring(3, close - 3).Trim();
                    var spoken = line[(close + 1)..].Trim();
                    line = string.IsNullOrWhiteSpace(speaker)
                        ? spoken
                        : $"{speaker}: {spoken}";
                }
            }

            if (builder.Length > 0)
            {
                builder.AppendLine();
            }

            builder.Append(line);
        }

        var text = builder.ToString().Trim();
        return string.IsNullOrWhiteSpace(text) ? content.Trim() : text;
    }

    private static string TruncateToRecent(string text, int maxCharacters)
    {
        if (string.IsNullOrWhiteSpace(text) || text.Length <= maxCharacters)
        {
            return text;
        }

        return text[(text.Length - maxCharacters)..];
    }

    private static void LogContext(
        string callId,
        bool available,
        int characters)
    {
        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" MEETING CONTEXT");
        Console.WriteLine("================================================");
        Console.WriteLine($"Call ID              : {callId}");
        Console.WriteLine($"Transcript available : {(available ? "Yes" : "No")}");
        Console.WriteLine($"Context characters   : {characters}");
        Console.WriteLine("================================================");
    }

    private static void LogGraphAccessDenied(string resource)
    {
        Console.WriteLine(
            $"[MEETING CONTEXT] Graph access not enabled for {resource}.");
        Console.WriteLine(
            "[MEETING CONTEXT] Transcript not available. Continuing without transcript.");
    }

    private static void LogGraphFailure(
        string resource,
        HttpStatusCode status)
    {
        Console.WriteLine(
            $"[MEETING CONTEXT] Graph {resource} returned {(int)status}.");
    }

    private static string? ReadGraphInnerError(string body)
    {
        try
        {
            using var document = JsonDocument.Parse(body);
            if (document.RootElement.TryGetProperty("error", out var error) &&
                error.TryGetProperty("innerError", out var inner) &&
                inner.TryGetProperty("code", out var code))
            {
                return code.GetString();
            }
        }
        catch
        {
            // Ignore malformed Graph error payloads.
        }

        return null;
    }

    private static bool IsSuccess(HttpStatusCode status)
    {
        var code = (int)status;
        return code >= 200 && code < 300;
    }

    private static string? GetString(JsonElement element, string name)
    {
        return element.TryGetProperty(name, out var value) &&
               value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;
    }

    private static DateTimeOffset? GetDate(JsonElement element, string name)
    {
        if (!element.TryGetProperty(name, out var value) ||
            value.ValueKind != JsonValueKind.String)
        {
            return null;
        }

        return DateTimeOffset.TryParse(value.GetString(), out var date)
            ? date
            : null;
    }

    private static string? NullIfEmpty(string? value)
    {
        return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
    }

    private sealed class CallMeetingState
    {
        public CallMeetingState(string callId)
        {
            CallId = callId;
        }

        public string CallId { get; }

        public string? TenantId { get; set; }

        public string? JoinMeetingId { get; set; }

        public string? OrganizerUserId { get; set; }

        public string? JoinWebUrl { get; set; }

        public string? ThreadId { get; set; }

        public string? OnlineMeetingId { get; set; }

        public string? CachedText { get; set; }

        public DateTimeOffset? CachedAt { get; set; }

        public object Gate { get; } = new();

        public SemaphoreSlim FetchLock { get; } = new(1, 1);
    }

    private sealed class LiveTranscriptState
    {
        public StringBuilder Text { get; } = new();

        public object Gate { get; } = new();
    }
}

using System.Net.Http.Headers;
using System.Text.Json;
using AgentTeamMateBot.Services;

namespace AgentTeamMateBot.Media;

public class MeetingMediaHandler
{
    private readonly MediaSessionService _mediaSessionService;

    public MeetingMediaHandler(MediaSessionService mediaSessionService)
    {
        _mediaSessionService = mediaSessionService;
    }

    public async Task<HttpResponseMessage> ProcessNotificationAsync(HttpRequest request)
    {
        Console.WriteLine();
        Console.WriteLine("======================================");
        Console.WriteLine(" PROCESSING TEAMS NOTIFICATION");
        Console.WriteLine("======================================");

        request.EnableBuffering();

        string body;
        using (var reader = new StreamReader(request.Body, leaveOpen: true))
        {
            body = await reader.ReadToEndAsync();
        }

        request.Body.Position = 0;

        Console.WriteLine(body);
        Console.WriteLine("================================================");

        HttpResponseMessage sdkResponse;
        try
        {
            using var requestMessage = ToHttpRequestMessage(request, body);
            sdkResponse = await _mediaSessionService.ProcessNotificationAsync(requestMessage);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Notification processing error : {ex.Message}");
            sdkResponse = new HttpResponseMessage(System.Net.HttpStatusCode.OK);
        }

        try
        {
            await StartMediaIfCallEstablished(body);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Notification processing error : {ex.Message}");
        }

        return sdkResponse;
    }

    private async Task StartMediaIfCallEstablished(string notificationBody)
    {
        if (string.IsNullOrWhiteSpace(notificationBody))
        {
            return;
        }

        using var json = JsonDocument.Parse(notificationBody);
        var root = json.RootElement;

        if (!root.TryGetProperty("value", out var values) || values.ValueKind != JsonValueKind.Array)
        {
            return;
        }

        foreach (var item in values.EnumerateArray())
        {
            var callId = ExtractCallId(GetString(item, "resource") ?? GetString(item, "resourceUrl"));
            var resourceData = item.TryGetProperty("resourceData", out var rd) ? rd : default;

            if (resourceData.ValueKind == JsonValueKind.Object &&
                resourceData.TryGetProperty("state", out var state) &&
                string.Equals(state.GetString(), "established", StringComparison.OrdinalIgnoreCase) &&
                callId != null)
            {
                Console.WriteLine($"Call established : {callId}");
                await _mediaSessionService.StartMediaSessionAsync(callId);
            }

            if (resourceData.ValueKind == JsonValueKind.Array)
            {
                foreach (var participant in resourceData.EnumerateArray())
                {
                    if (!participant.TryGetProperty("mediaStreams", out var streams))
                    {
                        continue;
                    }

                    foreach (var stream in streams.EnumerateArray())
                    {
                        if (stream.TryGetProperty("mediaType", out var mediaType) &&
                            mediaType.GetString() == "audio" &&
                            callId != null)
                        {
                            Console.WriteLine("Audio stream detected");
                            Console.WriteLine($"Call ID : {callId}");
                            await _mediaSessionService.StartMediaSessionAsync(callId);
                        }
                    }
                }
            }
        }
    }

    private static HttpRequestMessage ToHttpRequestMessage(HttpRequest request, string body)
    {
        var requestMessage = new HttpRequestMessage
        {
            Method = new HttpMethod(request.Method),
            RequestUri = new Uri(
                $"{request.Scheme}://{request.Host}{request.PathBase}{request.Path}{request.QueryString}"),
            Content = new StringContent(body)
        };

        requestMessage.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");

        foreach (var header in request.Headers)
        {
            if (!requestMessage.Headers.TryAddWithoutValidation(header.Key, header.Value.ToArray()))
            {
                requestMessage.Content.Headers.TryAddWithoutValidation(header.Key, header.Value.ToArray());
            }
        }

        return requestMessage;
    }

    private static string? GetString(JsonElement item, string name)
    {
        return item.TryGetProperty(name, out var value) ? value.GetString() : null;
    }

    private static string? ExtractCallId(string? resource)
    {
        if (string.IsNullOrEmpty(resource))
        {
            return null;
        }

        var parts = resource.Split('/');
        var index = Array.IndexOf(parts, "calls");

        if (index >= 0 && parts.Length > index + 1)
        {
            return parts[index + 1];
        }

        return null;
    }
}

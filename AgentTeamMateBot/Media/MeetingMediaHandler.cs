using System.Net.Http.Headers;
using System.Text.Json;
using AgentTeamMateBot.Services;

namespace AgentTeamMateBot.Media;

public class MeetingMediaHandler
{
    private readonly MediaSessionService _mediaSessionService;
    private readonly AppHostedMediaService _appHostedMediaService;

    public MeetingMediaHandler(
        MediaSessionService mediaSessionService,
        AppHostedMediaService appHostedMediaService)
    {
        _mediaSessionService =
            mediaSessionService;

        _appHostedMediaService =
            appHostedMediaService;
    }

    public async Task<HttpResponseMessage>
        ProcessNotificationAsync(
            HttpRequest request)
    {
        Console.WriteLine();
        Console.WriteLine("======================================");
        Console.WriteLine(" PROCESSING TEAMS NOTIFICATION");
        Console.WriteLine("======================================");

        request.EnableBuffering();

        string body;

        using (var reader =
               new StreamReader(
                   request.Body,
                   leaveOpen: true))
        {
            body =
                await reader
                    .ReadToEndAsync();
        }

        request.Body.Position = 0;

        Console.WriteLine(body);

        Console.WriteLine(
            "================================================");

        HttpResponseMessage sdkResponse;

        try
        {
            using var requestMessage =
                ToHttpRequestMessage(
                    request,
                    body);

            sdkResponse =
                await _mediaSessionService
                    .ProcessNotificationAsync(
                        requestMessage);

            if (_appHostedMediaService.IsInitialized)
            {
                using var appHostedRequest =
                    ToHttpRequestMessage(
                        request,
                        body);

                var appHostedResponse =
                    await _appHostedMediaService
                        .ProcessNotificationAsync(
                            appHostedRequest);

                if ((int)appHostedResponse.StatusCode >= 200 &&
                    (int)appHostedResponse.StatusCode < 300)
                {
                    sdkResponse = appHostedResponse;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(
                $"Notification SDK processing error : {ex.Message}");

            sdkResponse =
                new HttpResponseMessage(
                    System.Net.HttpStatusCode.OK);
        }

        try
        {
            await ProcessCustomNotificationAsync(
                body);
        }
        catch (Exception ex)
        {
            Console.WriteLine(
                $"Custom notification error : {ex.Message}");

            Console.WriteLine(
                ex);
        }

        return sdkResponse;
    }

    // ============================================================
    // CUSTOM GRAPH NOTIFICATION PROCESSING
    // ============================================================

    private async Task ProcessCustomNotificationAsync(
        string notificationBody)
    {
        if (string.IsNullOrWhiteSpace(
                notificationBody))
        {
            return;
        }

        using var json =
            JsonDocument.Parse(
                notificationBody);

        var root =
            json.RootElement;

        if (!root.TryGetProperty(
                "value",
                out var values) ||
            values.ValueKind !=
                JsonValueKind.Array)
        {
            return;
        }

        foreach (var item
                 in values.EnumerateArray())
        {
            var resource =
                GetString(
                    item,
                    "resource")
                ??
                GetString(
                    item,
                    "resourceUrl");

            var callId =
                ExtractCallId(
                    resource);

            if (!item.TryGetProperty(
                    "resourceData",
                    out var resourceData))
            {
                continue;
            }

            // ============================================================
            // RECORD OPERATION
            // ============================================================

            if (resourceData.ValueKind ==
                JsonValueKind.Object)
            {
                var odataType =
                    GetString(
                        resourceData,
                        "@odata.type");

                if (string.Equals(
                        odataType,
                        "#microsoft.graph.recordOperation",
                        StringComparison.OrdinalIgnoreCase))
                {
                    await HandleRecordOperationAsync(
                        callId,
                        resourceData);

                    continue;
                }

                if (string.Equals(
                        odataType,
                        "#microsoft.graph.playPromptOperation",
                        StringComparison.OrdinalIgnoreCase))
                {
                    await HandlePlayPromptOperationAsync(
                        callId,
                        resourceData);

                    continue;
                }
            }

            // ============================================================
            // CALL ESTABLISHED
            // ============================================================

            if (resourceData.ValueKind ==
                    JsonValueKind.Object &&
                resourceData.TryGetProperty(
                    "state",
                    out var state))
            {
                var callState =
                    state.GetString();

                if (string.Equals(
                        callState,
                        "establishing",
                        StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(
                        callState,
                        "established",
                        StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(
                        callState,
                        "terminated",
                        StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine();
                    Console.WriteLine(
                        "================================================");
                    Console.WriteLine(
                        " CALL STATE");
                    Console.WriteLine(
                        "================================================");
                    Console.WriteLine(
                        $"Call ID : {callId ?? "unknown"}");
                    Console.WriteLine(
                        $"State   : {callState}");

                    if (resourceData.TryGetProperty(
                            "resultInfo",
                            out var callResultInfo) &&
                        callResultInfo.ValueKind ==
                            JsonValueKind.Object)
                    {
                        var code =
                            GetInt(
                                callResultInfo,
                                "code");

                        var subcode =
                            GetInt(
                                callResultInfo,
                                "subcode");

                        var message =
                            GetString(
                                callResultInfo,
                                "message");

                        Console.WriteLine(
                            $"resultInfo.code    : {code}");
                        Console.WriteLine(
                            $"resultInfo.subcode : {subcode}");
                        Console.WriteLine(
                            $"resultInfo.message : {message}");

                        if (subcode == 1203002)
                        {
                            Console.WriteLine();
                            Console.WriteLine(
                                "BLOCKED: Graph media negotiation failed (1203002).");
                            Console.WriteLine(
                                "Continuous audio is NOT proven.");
                        }
                    }

                    Console.WriteLine(
                        "================================================");
                }

                if (string.Equals(
                        callState,
                        "established",
                        StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(
                        callId))
                {
                    Console.WriteLine();
                    Console.WriteLine(
                        $"Call established : {callId}");

                    if (_appHostedMediaService.IsAppHostedCall(
                            callId))
                    {
                        Console.WriteLine(
                            "[APP-HOSTED] Skipping recordResponse. Live AudioSocket listening is active.");

                        continue;
                    }

                    await _mediaSessionService
                        .StartMediaSessionAsync(
                            callId);
                }
            }
        }
    }

    // ============================================================
    // HANDLE RECORD OPERATION CALLBACK
    // ============================================================

    private async Task HandleRecordOperationAsync(
        string? callId,
        JsonElement resourceData)
    {
        var status =
            GetString(
                resourceData,
                "status");

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" RECORD OPERATION EVENT");
        Console.WriteLine("================================================");

        Console.WriteLine(
            $"Call ID : {callId ?? "unknown"}");

        Console.WriteLine(
            $"Status  : {status ?? "unknown"}");

        if (resourceData.TryGetProperty(
                "resultInfo",
                out var resultInfo) &&
            resultInfo.ValueKind ==
                JsonValueKind.Object)
        {
            var code =
                GetInt(
                    resultInfo,
                    "code");

            var subcode =
                GetInt(
                    resultInfo,
                    "subcode");

            var message =
                GetString(
                    resultInfo,
                    "message");

            Console.WriteLine(
                $"Result code    : {code}");

            Console.WriteLine(
                $"Result subcode : {subcode}");

            Console.WriteLine(
                $"Result message : {message}");
        }

        if (!string.Equals(
                status,
                "completed",
                StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine(
                "Record operation not completed yet.");

            Console.WriteLine(
                "================================================");

            return;
        }

        var recordingLocation =
            GetString(
                resourceData,
                "recordingLocation");

        var recordingAccessToken =
            GetString(
                resourceData,
                "recordingAccessToken");

        if (string.IsNullOrWhiteSpace(
                callId))
        {
            Console.WriteLine(
                "Call ID missing.");

            return;
        }

        if (_appHostedMediaService.IsAppHostedCall(
                callId))
        {
            Console.WriteLine(
                "[APP-HOSTED] Ignoring recordOperation. Continuous AudioSocket is the Phase 2 audio path.");

            return;
        }

        if (string.IsNullOrWhiteSpace(
                recordingLocation))
        {
            Console.WriteLine(
                "recordingLocation missing.");

            return;
        }

        if (string.IsNullOrWhiteSpace(
                recordingAccessToken))
        {
            Console.WriteLine(
                "recordingAccessToken missing.");

            return;
        }

        Console.WriteLine(
            "Recording is ready.");

        Console.WriteLine(
            "================================================");

        // Never print recordingAccessToken.

        await _mediaSessionService
            .ProcessCompletedRecordingAsync(
                callId,
                recordingLocation,
                recordingAccessToken);
    }

    // ============================================================
    // HANDLE PLAY PROMPT OPERATION CALLBACK
    // Restart listening only after playback has fully finished.
    // ============================================================

    private async Task HandlePlayPromptOperationAsync(
        string? callId,
        JsonElement resourceData)
    {
        var status =
            GetString(
                resourceData,
                "status");

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" PLAY PROMPT OPERATION EVENT");
        Console.WriteLine("================================================");

        Console.WriteLine(
            $"Call ID : {callId ?? "unknown"}");

        Console.WriteLine(
            $"Status  : {status ?? "unknown"}");

        int? code = null;
        int? subcode = null;
        string? message = null;

        if (resourceData.TryGetProperty(
                "resultInfo",
                out var resultInfo) &&
            resultInfo.ValueKind ==
                JsonValueKind.Object)
        {
            code =
                GetInt(
                    resultInfo,
                    "code");

            subcode =
                GetInt(
                    resultInfo,
                    "subcode");

            message =
                GetString(
                    resultInfo,
                    "message");

            Console.WriteLine(
                $"Result code    : {code}");

            Console.WriteLine(
                $"Result subcode : {subcode}");

            Console.WriteLine(
                $"Result message : {message}");
        }

        if (!string.Equals(
                status,
                "completed",
                StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine(
                "Play prompt operation not completed yet.");

            Console.WriteLine(
                "================================================");

            return;
        }

        if (code.HasValue &&
            code.Value >= 300)
        {
            Console.WriteLine(
                "Play prompt completed with an error. Not starting the next recording.");

            Console.WriteLine(
                "================================================");

            return;
        }

        if (string.IsNullOrWhiteSpace(
                callId))
        {
            Console.WriteLine(
                "Call ID missing.");

            Console.WriteLine(
                "================================================");

            return;
        }

        if (_appHostedMediaService.IsAppHostedCall(
                callId))
        {
            Console.WriteLine(
                "[APP-HOSTED] Ignoring playPromptOperation. Phase 2 Step 1 does not play TTS.");

            Console.WriteLine(
                "================================================");

            return;
        }

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" PLAY PROMPT COMPLETED");
        Console.WriteLine("================================================");

        Console.WriteLine(
            $"Call ID : {callId}");

        Console.WriteLine(
            "Starting next conversation turn...");

        Console.WriteLine(
            "================================================");

        await _mediaSessionService
            .StartNextConversationTurnAsync(
                callId);
    }

    // ============================================================
    // CONVERT ASP.NET REQUEST TO GRAPH SDK REQUEST
    // ============================================================

    private static HttpRequestMessage ToHttpRequestMessage(
        HttpRequest request,
        string body)
    {
        var requestMessage =
            new HttpRequestMessage
            {
                Method =
                    new HttpMethod(
                        request.Method),

                RequestUri =
                    new Uri(
                        $"{request.Scheme}://{request.Host}{request.PathBase}{request.Path}{request.QueryString}"),

                Content =
                    new StringContent(
                        body)
            };

        requestMessage
            .Content
            .Headers
            .ContentType =
            new MediaTypeHeaderValue(
                "application/json");

        foreach (var header
                 in request.Headers)
        {
            if (!requestMessage
                    .Headers
                    .TryAddWithoutValidation(
                        header.Key,
                        header.Value.ToArray()))
            {
                requestMessage
                    .Content
                    .Headers
                    .TryAddWithoutValidation(
                        header.Key,
                        header.Value.ToArray());
            }
        }

        return requestMessage;
    }

    // ============================================================
    // JSON STRING HELPER
    // ============================================================

    private static string? GetString(
        JsonElement element,
        string name)
    {
        if (!element.TryGetProperty(
                name,
                out var value))
        {
            return null;
        }

        if (value.ValueKind ==
            JsonValueKind.String)
        {
            return value.GetString();
        }

        return value.ToString();
    }

    // ============================================================
    // JSON INT HELPER
    // ============================================================

    private static int? GetInt(
        JsonElement element,
        string name)
    {
        if (!element.TryGetProperty(
                name,
                out var value))
        {
            return null;
        }

        if (value.ValueKind ==
                JsonValueKind.Number &&
            value.TryGetInt32(
                out var number))
        {
            return number;
        }

        return null;
    }

    // ============================================================
    // EXTRACT CALL ID
    //
    // Example:
    //
    // /app/calls/{callId}
    //
    // /communications/calls/{callId}/operations/{operationId}
    // ============================================================

    private static string? ExtractCallId(
        string? resource)
    {
        if (string.IsNullOrWhiteSpace(
                resource))
        {
            return null;
        }

        var parts =
            resource.Split(
                '/',
                StringSplitOptions.RemoveEmptyEntries);

        for (var i = 0;
             i < parts.Length - 1;
             i++)
        {
            if (string.Equals(
                    parts[i],
                    "calls",
                    StringComparison.OrdinalIgnoreCase))
            {
                return parts[i + 1];
            }
        }

        return null;
    }
}
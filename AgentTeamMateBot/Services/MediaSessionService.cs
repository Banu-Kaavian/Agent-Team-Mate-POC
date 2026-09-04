using System.Collections.Concurrent;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using AgentTeamMateBot.Media;
using Microsoft.Graph.Communications.Calls;
using Microsoft.Graph.Communications.Client;
using Microsoft.Graph.Communications.Common.Telemetry;
using Microsoft.Graph.Communications.Resources;
using Microsoft.Graph.Contracts;
using Microsoft.Graph.Models;

namespace AgentTeamMateBot.Services;

public class MediaSessionService
{
    private readonly IConfiguration _configuration;
    private readonly GraphAuthService _graphAuthService;
    private readonly AudioHandler _audioHandler;
    private readonly SpeechRecognitionService _speechService;
    private readonly AiResponseService _aiResponseService;
    private readonly SpeechSynthesisService _speechSynthesisService;
    private readonly MeetingContextService _meetingContextService;
    private readonly MeetingExportService _meetingExportService;

    private readonly ConcurrentDictionary<string, ICall> _calls = new();
    private readonly ConcurrentDictionary<string, byte> _recordingStarted = new();
    private readonly ConcurrentDictionary<string, byte> _playbackInProgress = new();

    private readonly HttpClient _httpClient = new();

    private readonly object _initLock = new();

    private IGraphLogger? _graphLogger;
    private ICommunicationsClient? _client;

    private bool _initialized;
    private string? _initError;

    public MediaSessionService(
        IConfiguration configuration,
        GraphAuthService graphAuthService,
        AudioHandler audioHandler,
        SpeechRecognitionService speechService,
        AiResponseService aiResponseService,
        SpeechSynthesisService speechSynthesisService,
        MeetingContextService meetingContextService,
        MeetingExportService meetingExportService)
    {
        _configuration = configuration;
        _graphAuthService = graphAuthService;
        _audioHandler = audioHandler;
        _speechService = speechService;
        _aiResponseService = aiResponseService;
        _speechSynthesisService = speechSynthesisService;
        _meetingContextService = meetingContextService;
        _meetingExportService = meetingExportService;
    }

    public ICommunicationsClient? Client => _client;

    public bool IsInitialized => _initialized;

    // ============================================================
    // INITIALIZE GRAPH CLIENT
    // ============================================================

    public void Initialize()
    {
        lock (_initLock)
        {
            if (_initialized)
            {
                return;
            }

            try
            {
                Console.WriteLine();
                Console.WriteLine("================================================");
                Console.WriteLine(" INITIALIZING GRAPH COMMUNICATIONS CLIENT");
                Console.WriteLine("================================================");

                var clientId =
                    _graphAuthService.ClientId;

                var callbackUri =
                    _configuration["Bot:CallbackUri"]
                    ?? "https://teammate-bot.westus3.cloudapp.azure.com/api/calling";

                _graphLogger =
                    new GraphLogger(
                        "AgentTeamMateBot",
                        redirectToTrace: true);

                var builder =
                    new CommunicationsClientBuilder(
                        "AgentTeamMateBot",
                        clientId,
                        _graphLogger);

#pragma warning disable CS0618

                builder.SetAuthenticationProvider(
                    _graphAuthService
                        .CreateAuthenticationProvider(
                            _graphLogger));

#pragma warning restore CS0618

                builder.SetNotificationUrl(
                    new Uri(callbackUri));

                builder.SetServiceBaseUrl(
                    new Uri(
                        "https://graph.microsoft.com/v1.0"));

                _client =
                    builder.Build();

                _client
                    .Calls()
                    .OnUpdated += OnCallsUpdated;

                _initialized = true;
                _initError = null;

                Console.WriteLine(
                    $"Application ID    : {clientId}");

                Console.WriteLine(
                    $"Notification URL  : {callbackUri}");

                Console.WriteLine(
                    "Media mode        : SERVICE HOSTED");

                Console.WriteLine(
                    "App MediaPlatform : DISABLED");

                Console.WriteLine(
                    "AudioSocket       : NOT USED");

                Console.WriteLine(
                    "Voice capture     : recordResponse");

                Console.WriteLine(
                    "GRAPH COMMUNICATIONS CLIENT INITIALIZED");

                Console.WriteLine(
                    "================================================");
            }
            catch (Exception ex)
            {
                _initialized = false;
                _initError = ex.Message;

                Console.WriteLine();
                Console.WriteLine("================================================");
                Console.WriteLine(" GRAPH CLIENT INITIALIZATION FAILURE");
                Console.WriteLine("================================================");

                Console.WriteLine(
                    ex.Message);

                Console.WriteLine(
                    ex);
            }
        }
    }

    // ============================================================
    // JOIN TEAMS MEETING
    // ============================================================

    public async Task<ICall> JoinMeetingAsync(
        string meetingId,
        string? passcode,
        string? organizerUserId = null,
        string? joinWebUrl = null)
    {
        if (!_initialized ||
            _client == null)
        {
            Initialize();
        }

        if (!_initialized ||
            _client == null)
        {
            throw new InvalidOperationException(
                "Graph communications client is not initialized. "
                + (_initError ?? "See initialization logs."));
        }

        try
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine("       JOIN MEETING REQUEST RECEIVED");
            Console.WriteLine("================================================");

            var tenantId =
                _graphAuthService
                    .TenantId
                    .Trim();

            var normalizedMeetingId =
                NormalizeMeetingId(
                    meetingId);

            var normalizedPasscode =
                string.IsNullOrWhiteSpace(
                    passcode)
                    ? null
                    : passcode.Trim();

            if (string.IsNullOrWhiteSpace(
                    normalizedMeetingId))
            {
                throw new ArgumentException(
                    "Meeting ID is required.",
                    nameof(meetingId));
            }

            Console.WriteLine(
                $"Meeting ID : {normalizedMeetingId}");

            Console.WriteLine(
                $"Tenant ID  : {tenantId}");

            Console.WriteLine(
                "Media mode : SERVICE HOSTED");

            var applicationIdentity =
                new Identity
                {
                    Id =
                        _graphAuthService.ClientId,

                    DisplayName =
                        "Agent Team Mate"
                };

            applicationIdentity.SetTenantId(
                tenantId);

            var serviceHostedMediaConfig =
                new ServiceHostedMediaConfig
                {
                    OdataType =
                        "#microsoft.graph.serviceHostedMediaConfig"
                };

            var call =
                new Call
                {
                    OdataType =
                        "#microsoft.graph.call",

                    CallbackUri =
                        _configuration["Bot:CallbackUri"]
                        ?? "https://teammate-bot.westus3.cloudapp.azure.com/api/calling",

                    TenantId =
                        tenantId,

                    Source =
                        new ParticipantInfo
                        {
                            Identity =
                                new IdentitySet
                                {
                                    Application =
                                        applicationIdentity
                                }
                        },

                    RequestedModalities =
                        new List<Modality?>
                        {
                            Modality.Audio
                        },

                    MediaConfig =
                        serviceHostedMediaConfig,

                    MeetingInfo =
                        new JoinMeetingIdMeetingInfo
                        {
                            OdataType =
                                "#microsoft.graph.joinMeetingIdMeetingInfo",

                            JoinMeetingId =
                                normalizedMeetingId,

                            Passcode =
                                normalizedPasscode
                        }
                };

            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" JOINING TEAMS MEETING WITH SERVICE-HOSTED MEDIA");
            Console.WriteLine("================================================");

            Console.WriteLine(
                $"Meeting ID : {normalizedMeetingId}");

            Console.WriteLine(
                $"Tenant ID  : {tenantId}");

            Console.WriteLine(
                "[MEDIA] Microsoft hosts the media.");

            Console.WriteLine(
                "[MEDIA] No LocalMediaSession.");

            Console.WriteLine(
                "[MEDIA] No AudioSocket.");

            Console.WriteLine(
                "[MEDIA] No MediaPlatform.");

            var statefulCall =
                await _client
                    .Calls()
                    .AddAsync(call);

            TrackCall(
                statefulCall);

            _meetingContextService.RegisterCall(
                statefulCall.Id,
                tenantId,
                normalizedMeetingId,
                organizerUserId,
                joinWebUrl);

            _meetingContextService.EnrichFromCallResource(
                statefulCall.Id,
                statefulCall.Resource);

            Console.WriteLine(
                $"Call ID : {statefulCall.Id}");

            Console.WriteLine(
                $"State   : {statefulCall.Resource?.State}");

            Console.WriteLine(
                "================================================");

            Console.WriteLine();
            Console.WriteLine(
                "JOIN REQUEST SENT TO MICROSOFT GRAPH");

            BotLog.Info($"Joined meeting {normalizedMeetingId}.");

            Console.WriteLine(
                $"Call ID : {statefulCall.Id}");

            Console.WriteLine(
                $"State   : {statefulCall.Resource?.State}");

            return statefulCall;
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" SERVICE-HOSTED MEETING JOIN FAILURE");
            Console.WriteLine("================================================");

            Console.WriteLine(
                ex.Message);

            Console.WriteLine(
                ex);

            throw;
        }
    }

    // ============================================================
    // CALL ESTABLISHED
    // START SHORT VOICE RECORDING
    // ============================================================

    public async Task StartMediaSessionAsync(
        string callId)
    {
        if (string.IsNullOrWhiteSpace(
                callId))
        {
            return;
        }

        if (_playbackInProgress.ContainsKey(
                callId))
        {
            Console.WriteLine(
                $"[RECORD] Skipping recordResponse for {callId} because playPrompt is still in progress.");

            return;
        }

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" SERVICE-HOSTED CALL ESTABLISHED");
        Console.WriteLine("================================================");

        Console.WriteLine(
            $"Call ID : {callId}");

        Console.WriteLine(
            "Media   : Hosted by Microsoft");

        Console.WriteLine(
            "Playing welcome prompt, then listening...");

        Console.WriteLine(
            "================================================");

        try
        {
            BotLog.Info("Playing welcome...");

            var welcomeUrl =
                await _speechSynthesisService
                    .SynthesizeSpeechAsync(
                        "Hi, I am Agent Nova. I am listening. Say your request.");

            if (!string.IsNullOrWhiteSpace(welcomeUrl))
            {
                await PlayPromptAsync(
                    callId,
                    welcomeUrl);

                return;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(
                $"[WELCOME] Could not play greeting: {ex.Message}");
        }

        if (!_recordingStarted.TryAdd(
                callId,
                0))
        {
            Console.WriteLine(
                $"[RECORD] recordResponse already started for {callId}");

            return;
        }

        await StartRecordResponseAsync(
            callId);
    }

    // ============================================================
    // NEXT CONVERSATION TURN
    // Immediately restarts recording. playPrompt will barge in
    // if an AI response arrives while recording.
    // ============================================================

    public async Task StartNextConversationTurnAsync(
        string callId)
    {
        if (string.IsNullOrWhiteSpace(
                callId))
        {
            Console.WriteLine(
                "[LISTEN LOOP] Restart skipped: call ID is missing.");

            return;
        }

        _playbackInProgress.TryRemove(
            callId,
            out _);

        if (!_recordingStarted.TryAdd(
                callId,
                0))
        {
            Console.WriteLine(
                $"[LISTEN LOOP] recordResponse already active for {callId}, skipping.");

            return;
        }

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" NEXT CONVERSATION TURN");
        Console.WriteLine("================================================");

        Console.WriteLine(
            $"Call ID : {callId}");

        Console.WriteLine(
            "Listening for user...");

        Console.WriteLine(
            "================================================");

        await StartRecordResponseAsync(
            callId);
    }

    // ============================================================
    // START recordResponse
    // ============================================================

    public async Task StartRecordResponseAsync(
        string callId)
    {
        try
        {
            var token =
                await _graphAuthService
                    .GetAccessTokenAsync();

            var url =
                $"https://graph.microsoft.com/v1.0/communications/calls/{callId}/recordResponse";

            var clientContext =
                Guid.NewGuid()
                    .ToString();

            var payload =
                new
                {
                    bargeInAllowed =
                        true,

                    clientContext =
                        clientContext,

                    maxRecordDurationInSeconds =
                        _configuration.GetValue(
                            "Recording:MaxDurationSeconds", 5),

                    initialSilenceTimeoutInSeconds =
                        _configuration.GetValue(
                            "Recording:InitialSilenceTimeoutSeconds", 600),

                    maxSilenceTimeoutInSeconds =
                        _configuration.GetValue(
                            "Recording:SilenceTimeoutSeconds", 1),

                    playBeep =
                        _configuration.GetValue(
                            "Recording:PlayBeep", false),

                    stopTones =
                        new[]
                        {
                            "#"
                        }
                };

            var json =
                JsonSerializer.Serialize(
                    payload);

            using var request =
                new HttpRequestMessage(
                    HttpMethod.Post,
                    url);

            request.Headers.Authorization =
                new AuthenticationHeaderValue(
                    "Bearer",
                    token);

            request.Content =
                new StringContent(
                    json,
                    Encoding.UTF8,
                    "application/json");

            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" STARTING RECORD RESPONSE");
            Console.WriteLine("================================================");

            Console.WriteLine(
                $"Call ID : {callId}");

            var response =
                await _httpClient
                    .SendAsync(request);

            var responseBody =
                await response.Content
                    .ReadAsStringAsync();

            Console.WriteLine(
                $"Status : {(int)response.StatusCode} {response.StatusCode}");

            if (!string.IsNullOrWhiteSpace(
                    responseBody))
            {
                Console.WriteLine(
                    responseBody);
            }

            if (!response.IsSuccessStatusCode)
            {
                _recordingStarted.TryRemove(
                    callId,
                    out _);

                throw new Exception(
                    $"recordResponse failed: {(int)response.StatusCode} {responseBody}");
            }

            Console.WriteLine();
            Console.WriteLine(
                "recordResponse started successfully.");

            BotLog.Info("Listening...");

            Console.WriteLine(
                "================================================");
        }
        catch (Exception ex)
        {
            _recordingStarted.TryRemove(
                callId,
                out _);

            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" RECORD RESPONSE FAILURE");
            Console.WriteLine("================================================");

            Console.WriteLine(
                ex.Message);

            Console.WriteLine(
                ex);
        }
    }

    // ============================================================
    // DOWNLOAD RECORDING
    // STT → AI → TTS → playPrompt
    // ============================================================

    public async Task ProcessCompletedRecordingAsync(
        string callId,
        string recordingLocation,
        string recordingAccessToken)
    {
        try
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" RECORD RESPONSE COMPLETED");
            Console.WriteLine("================================================");

            Console.WriteLine(
                $"Call ID : {callId}");

            Console.WriteLine(
                "Downloading recorded audio...");

            using var request =
                new HttpRequestMessage(
                    HttpMethod.Get,
                    recordingLocation);

            request.Headers.Authorization =
                new AuthenticationHeaderValue(
                    "Bearer",
                    recordingAccessToken);

            var response =
                await _httpClient
                    .SendAsync(request);

            if (!response.IsSuccessStatusCode)
            {
                var error =
                    await response.Content
                        .ReadAsStringAsync();

                Console.WriteLine(
                    $"Recording download failed: {(int)response.StatusCode}");

                Console.WriteLine(
                    error);

                return;
            }

            var recordingBytes =
                await response.Content
                    .ReadAsByteArrayAsync();

            Console.WriteLine(
                $"Recording downloaded : {recordingBytes.Length} bytes");

            if (recordingBytes.Length == 0)
            {
                Console.WriteLine(
                    "[RECORD] Recording contains no audio.");

                return;
            }

            _ = Task.Run(async () =>
            {
                try
                {
                    await RecognizeAndRespondAsync(
                        callId,
                        recordingBytes);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(
                        $"[BACKGROUND AI] Error: {ex.Message}");
                    BotLog.Info($"Error: {ex.Message}");
                }
            });
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" RECORDING PROCESSING FAILURE");
            Console.WriteLine("================================================");

            Console.WriteLine(
                ex.Message);

            Console.WriteLine(
                ex);
        }
        finally
        {
            _recordingStarted.TryRemove(
                callId,
                out _);

            // Immediately restart recording regardless of outcome.
            // playPrompt will barge in if AI responds while recording.
            await StartNextConversationTurnAsync(
                callId);
        }
    }

    // ============================================================
    // BACKGROUND: AI → TTS → PLAY PROMPT
    // ============================================================

    private async Task RecognizeAndRespondAsync(
        string callId,
        byte[] recordingBytes)
    {
        Console.WriteLine();
        Console.WriteLine(
            "Sending recording to Azure Speech...");

        var recognizedText =
            await _speechService
                .RecognizeRecordingAsync(
                    recordingBytes);

        if (string.IsNullOrWhiteSpace(
                recognizedText))
        {
            Console.WriteLine();
            Console.WriteLine(
                "[SPEECH] No recognizable speech.");

            return;
        }

        var transcript =
            _meetingContextService.AppendLiveTranscript(
                callId,
                recognizedText);

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" MEETING TRANSCRIPT");
        Console.WriteLine("================================================");
        Console.WriteLine(
            $"Call ID : {callId}");
        Console.WriteLine(
            transcript);
        Console.WriteLine(
            "================================================");

        var requireWakeWord =
            _configuration.GetValue(
                "Recording:RequireWakeWord",
                false);

        var invoked =
            WakeWordDetector.IsAgentInvocation(
                recognizedText);

        if (requireWakeWord && !invoked)
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" PASSIVE MEETING SPEECH");
            Console.WriteLine("================================================");
            Console.WriteLine(
                $"Heard: {recognizedText}");
            Console.WriteLine(
                "Say 'Agent Nova' plus your question.");
            Console.WriteLine(
                "================================================");
            return;
        }

        if ((!requireWakeWord || invoked) &&
            WakeWordDetector.IsSummaryExportRequest(recognizedText))
        {
            BotLog.Info($"User: {recognizedText}");
            BotLog.Info("Exporting meeting summary...");
            var spoken = await _meetingExportService.ExportMeetingSummaryAsync(callId);
            await SpeakAsync(callId, spoken);

            if (WakeWordDetector.IsLeaveMeetingRequest(recognizedText))
            {
                await LeaveMeetingAsync(callId);
            }

            return;
        }

        if (invoked &&
            WakeWordDetector.IsLeaveMeetingRequest(recognizedText))
        {
            BotLog.Info($"User: {recognizedText}");
            BotLog.Info("Leaving meeting...");
            await LeaveMeetingAsync(callId);
            return;
        }

        if (requireWakeWord &&
            invoked &&
            !WakeWordDetector.IsActionableRequest(recognizedText))
        {
            Console.WriteLine(
                $"[LISTEN] Ignoring casual Agent Nova mention: {recognizedText}");
            return;
        }

        var question =
            invoked
                ? WakeWordDetector.RemoveActivationPhrase(
                    recognizedText)
                : recognizedText;

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" AGENT NOVA INVOCATION DETECTED");
        Console.WriteLine("================================================");
        Console.WriteLine(
            $"Original speech : {recognizedText}");
        Console.WriteLine(
            $"Question        : {question}");
        Console.WriteLine(
            "================================================");

        BotLog.Info($"User: {recognizedText}");
        BotLog.Info("Processing...");

        await ProcessAgentResponseAsync(
            callId,
            question);
    }

    private async Task SpeakAsync(string callId, string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        var audioUrl = await _speechSynthesisService.SynthesizeSpeechAsync(text);
        if (string.IsNullOrWhiteSpace(audioUrl))
        {
            return;
        }

        await PlayPromptAsync(callId, audioUrl);
    }

    private async Task LeaveMeetingAsync(string callId)
    {
        try
        {
            if (!_calls.TryGetValue(callId, out var call))
            {
                BotLog.Info("Error: Call not found; cannot leave meeting.");
                return;
            }

            try
            {
                var goodbyeUrl =
                    await _speechSynthesisService.SynthesizeSpeechAsync(
                        "Goodbye. I am leaving the meeting now.");
                if (!string.IsNullOrWhiteSpace(goodbyeUrl))
                {
                    await PlayPromptAsync(callId, goodbyeUrl);
                    await Task.Delay(2500);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[LEAVE] Goodbye audio failed: {ex.Message}");
            }

            await call.DeleteAsync();
            BotLog.Info("Left the meeting.");
        }
        catch (Exception ex)
        {
            BotLog.Info($"Error: Could not leave meeting. {ex.Message}");
            Console.WriteLine($"[LEAVE] {ex}");
        }
    }

    private async Task ProcessAgentResponseAsync(
        string callId,
        string question)
    {
        Console.WriteLine();
        Console.WriteLine(
            "Sending recognized speech to Azure OpenAI...");

        var aiResponse =
            await _aiResponseService
                .GetResponseAsync(
                    callId,
                    question);

        if (string.IsNullOrWhiteSpace(
                aiResponse))
        {
            Console.WriteLine();
            Console.WriteLine(
                "[AI] No response received.");

            BotLog.Info("Error: No AI response text.");
            return;
        }

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" AGENT TEAM MATE RESPONSE");
        Console.WriteLine("================================================");

        Console.WriteLine(
            aiResponse);

        Console.WriteLine(
            "================================================");

        BotLog.Info($"Nova: {aiResponse}");

        Console.WriteLine();
        Console.WriteLine(
            "Generating Agent Team Mate voice...");

        var audioUrl =
            await _speechSynthesisService
                .SynthesizeSpeechAsync(
                    aiResponse);

        if (string.IsNullOrWhiteSpace(
                    audioUrl))
        {
            Console.WriteLine();
            Console.WriteLine(
                "[TTS] No audio URL generated.");

            BotLog.Info("Error: Could not generate voice audio.");
            return;
        }

        Console.WriteLine();
        Console.WriteLine(
            $"Generated audio URL : {audioUrl}");

        await PlayPromptAsync(
            callId,
            audioUrl);
    }

    // ============================================================
    // PLAY TTS AUDIO INSIDE TEAMS
    // ============================================================

    private async Task PlayPromptAsync(
        string callId,
        string audioUrl)
    {
        try
        {
            var token =
                await _graphAuthService
                    .GetAccessTokenAsync();

            var url =
                $"https://graph.microsoft.com/v1.0/communications/calls/{callId}/playPrompt";

            var mediaInfo =
                new Dictionary<string, object?>
                {
                    ["@odata.type"] =
                        "#microsoft.graph.mediaInfo",

                    ["uri"] =
                        audioUrl,

                    ["resourceId"] =
                        Guid.NewGuid()
                            .ToString()
                };

            var mediaPrompt =
                new Dictionary<string, object?>
                {
                    ["@odata.type"] =
                        "#microsoft.graph.mediaPrompt",

                    ["mediaInfo"] =
                        mediaInfo
                };

            var payload =
                new Dictionary<string, object?>
                {
                    ["prompts"] =
                        new object[]
                        {
                            mediaPrompt
                        },

                    ["clientContext"] =
                        Guid.NewGuid()
                            .ToString()
                };

            var json =
                JsonSerializer.Serialize(
                    payload);

            using var request =
                new HttpRequestMessage(
                    HttpMethod.Post,
                    url);

            request.Headers.Authorization =
                new AuthenticationHeaderValue(
                    "Bearer",
                    token);

            request.Content =
                new StringContent(
                    json,
                    Encoding.UTF8,
                    "application/json");

            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" PLAYING AI RESPONSE IN TEAMS");
            Console.WriteLine("================================================");

            Console.WriteLine(
                $"Audio URL : {audioUrl}");

            var response =
                await _httpClient
                    .SendAsync(request);

            var responseBody =
                await response.Content
                    .ReadAsStringAsync();

            Console.WriteLine(
                $"Status : {(int)response.StatusCode} {response.StatusCode}");

            if (!string.IsNullOrWhiteSpace(
                    responseBody))
            {
                Console.WriteLine(
                    responseBody);
            }

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception(
                    $"playPrompt failed: {(int)response.StatusCode} {responseBody}");
            }

            Console.WriteLine();
            Console.WriteLine(
                "playPrompt started successfully.");

            Console.WriteLine(
                "================================================");

            _playbackInProgress.TryAdd(
                callId,
                0);
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" PLAY PROMPT FAILURE");
            Console.WriteLine("================================================");

            Console.WriteLine(
                ex.Message);

            Console.WriteLine(
                ex);

            Console.WriteLine(
                "================================================");

            throw;
        }
    }

    // ============================================================
    // PROCESS GRAPH CALLBACK
    // ============================================================

    public async Task<HttpResponseMessage>
        ProcessNotificationAsync(
            HttpRequestMessage requestMessage)
    {
        if (_client == null)
        {
            return new HttpResponseMessage(
                HttpStatusCode.ServiceUnavailable);
        }

        try
        {
            return await _client
                .ProcessNotificationAsync(
                    requestMessage);
        }
        catch (Exception ex)
        {
            Console.WriteLine(
                "[GRAPH] Notification processing failure:");

            Console.WriteLine(
                ex);

            return new HttpResponseMessage(
                HttpStatusCode.InternalServerError);
        }
    }

    // ============================================================
    // TRACK CALL
    // ============================================================

    private void TrackCall(
        ICall call)
    {
        if (_calls.TryAdd(
                call.Id,
                call))
        {
            call.OnUpdated +=
                OnCallResourceUpdated;
        }
    }

    // ============================================================
    // CALL COLLECTION EVENTS
    // ============================================================

    private void OnCallsUpdated(
        ICallCollection sender,
        CollectionEventArgs<ICall> args)
    {
        foreach (var call
                 in args.AddedResources)
        {
            Console.WriteLine(
                $"[CALL] Added {call.Id} state={call.Resource?.State}");

            TrackCall(
                call);

            _meetingContextService.EnrichFromCallResource(
                call.Id,
                call.Resource);
        }

        foreach (var call
                 in args.UpdatedResources)
        {
            HandleCallStateChange(
                call);
        }

        foreach (var call
                 in args.RemovedResources)
        {
            Console.WriteLine(
                $"[CALL] Removed {call.Id}");

            _recordingStarted.TryRemove(
                call.Id,
                out _);

            _playbackInProgress.TryRemove(
                call.Id,
                out _);

            _aiResponseService.ClearConversation(
                call.Id);

            _meetingContextService.Clear(
                call.Id);

            if (_calls.TryRemove(
                    call.Id,
                    out var existing))
            {
                existing.OnUpdated -=
                    OnCallResourceUpdated;
            }
        }
    }

    // ============================================================
    // INDIVIDUAL CALL UPDATE
    // ============================================================

    private void OnCallResourceUpdated(
        ICall sender,
        ResourceEventArgs<Call> args)
    {
        HandleCallStateChange(
            sender);
    }

    // ============================================================
    // HANDLE CALL STATE
    // ============================================================

    private void HandleCallStateChange(
        ICall call)
    {
        var state =
            call.Resource?.State;

        Console.WriteLine(
            $"[CALL] {call.Id} state={state}");

        _meetingContextService.EnrichFromCallResource(
            call.Id,
            call.Resource);

        if (state ==
            CallState.Established)
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine("     TEAMS CALL ESTABLISHED");
            Console.WriteLine("================================================");

            Console.WriteLine(
                $"Call ID : {call.Id}");

            Console.WriteLine(
                "Media mode : SERVICE HOSTED");

            Console.WriteLine(
                "Agent Team Mate has joined the meeting.");

            Console.WriteLine(
                "================================================");
        }

        if (state ==
            CallState.Terminated)
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine("       TEAMS CALL TERMINATED");
            Console.WriteLine("================================================");

            Console.WriteLine(
                $"Call ID : {call.Id}");

            var resultInfo =
                call.Resource?.ResultInfo;

            if (resultInfo != null)
            {
                Console.WriteLine(
                    $"Code    : {resultInfo.Code}");

                Console.WriteLine(
                    $"Subcode : {resultInfo.Subcode}");

                Console.WriteLine(
                    $"Message : {resultInfo.Message}");
            }

            _recordingStarted.TryRemove(
                call.Id,
                out _);

            _playbackInProgress.TryRemove(
                call.Id,
                out _);

            _aiResponseService.ClearConversation(
                call.Id);

            _meetingContextService.Clear(
                call.Id);

            Console.WriteLine(
                "================================================");
        }
    }

    // ============================================================
    // NORMALIZE MEETING ID
    // ============================================================

    private static string NormalizeMeetingId(
        string meetingId)
    {
        if (string.IsNullOrWhiteSpace(
                meetingId))
        {
            return string.Empty;
        }

        return new string(
            meetingId
                .Where(char.IsDigit)
                .ToArray());
    }
}
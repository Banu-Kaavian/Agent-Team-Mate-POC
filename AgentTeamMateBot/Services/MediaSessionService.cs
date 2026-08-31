using System.Collections.Concurrent;
using System.Net;
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

    // Keep these dependencies for now so we do not disturb
    // the rest of the existing project/DI configuration.
    private readonly AudioHandler _audioHandler;
    private readonly SpeechRecognitionService _speechService;

    private readonly ConcurrentDictionary<string, ICall> _calls =
        new();

    private readonly object _initLock =
        new();

    private IGraphLogger? _graphLogger;
    private ICommunicationsClient? _client;

    private bool _initialized;
    private string? _initError;

    public MediaSessionService(
        IConfiguration configuration,
        GraphAuthService graphAuthService,
        AudioHandler audioHandler,
        SpeechRecognitionService speechService)
    {
        _configuration =
            configuration;

        _graphAuthService =
            graphAuthService;

        _audioHandler =
            audioHandler;

        _speechService =
            speechService;
    }

    public ICommunicationsClient? Client =>
        _client;

    public bool IsInitialized =>
        _initialized;

    // ============================================================
    // INITIALIZE GRAPH COMMUNICATIONS CLIENT
    //
    // IMPORTANT:
    // This version DOES NOT initialize the app-hosted MediaPlatform.
    // Microsoft hosts the media for the call.
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
                Console.WriteLine(
                    "================================================");

                Console.WriteLine(
                    " INITIALIZING GRAPH COMMUNICATIONS CLIENT");

                Console.WriteLine(
                    "================================================");

                var clientId =
                    _graphAuthService.ClientId;

                var callbackUri =
                    _configuration["Bot:CallbackUri"]
                    ?? "https://teammate-bot.westus3.cloudapp.azure.com/api/calling";

                // ============================================================
                // GRAPH LOGGER
                // ============================================================

                _graphLogger =
                    new GraphLogger(
                        "AgentTeamMateBot",
                        redirectToTrace: true);

                // ============================================================
                // GRAPH COMMUNICATIONS CLIENT
                // ============================================================

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
                    new Uri(
                        callbackUri));

                builder.SetServiceBaseUrl(
                    new Uri(
                        "https://graph.microsoft.com/v1.0"));

                /*
                 * IMPORTANT
                 * ============================================================
                 *
                 * We intentionally DO NOT call:
                 *
                 * builder.SetMediaPlatformSettings(...)
                 *
                 * because this implementation uses:
                 *
                 * #microsoft.graph.serviceHostedMediaConfig
                 *
                 * Microsoft will host the call media.
                 *
                 * Therefore we do NOT require:
                 *
                 * - MediaPlatform
                 * - AudioSocket
                 * - VideoSocket
                 * - NativeMedia
                 * - UDP media session
                 * - CreateMediaSession()
                 *
                 * for the join operation.
                 */

                _client =
                    builder.Build();

                _client
                    .Calls()
                    .OnUpdated +=
                    OnCallsUpdated;

                _initialized =
                    true;

                _initError =
                    null;

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
                    "GRAPH COMMUNICATIONS CLIENT INITIALIZED");

                Console.WriteLine(
                    "================================================");
            }
            catch (Exception ex)
            {
                _initialized =
                    false;

                _initError =
                    ex.Message;

                Console.WriteLine();
                Console.WriteLine(
                    "================================================");

                Console.WriteLine(
                    " GRAPH CLIENT INITIALIZATION FAILURE");

                Console.WriteLine(
                    "================================================");

                Console.WriteLine(
                    ex.Message);

                Console.WriteLine(
                    ex);
            }
        }
    }

    // ============================================================
    // JOIN TEAMS MEETING
    // SERVICE-HOSTED MEDIA
    // ============================================================

    public async Task<ICall> JoinMeetingAsync(
        string meetingId,
        string? passcode)
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
                + (_initError
                   ?? "See initialization logs."));
        }

        try
        {
            Console.WriteLine();
            Console.WriteLine(
                "================================================");

            Console.WriteLine(
                "       JOIN MEETING REQUEST RECEIVED");

            Console.WriteLine(
                "================================================");

            // ============================================================
            // NORMALIZE MEETING DETAILS
            // ============================================================

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

            // ============================================================
            // APPLICATION IDENTITY
            // ============================================================

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

            // ============================================================
            // SERVICE-HOSTED MEDIA CONFIGURATION
            // ============================================================

            var serviceHostedMediaConfig =
                new ServiceHostedMediaConfig
                {
                    OdataType =
                        "#microsoft.graph.serviceHostedMediaConfig"
                };

            // ============================================================
            // GRAPH CALL
            // ============================================================

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
            Console.WriteLine(
                "================================================");

            Console.WriteLine(
                " JOINING TEAMS MEETING WITH SERVICE-HOSTED MEDIA");

            Console.WriteLine(
                "================================================");

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

            // ============================================================
            // CREATE CALL
            // ============================================================

            var statefulCall =
                await _client
                    .Calls()
                    .AddAsync(
                        call);

            TrackCall(
                statefulCall);

            Console.WriteLine(
                $"Call ID : {statefulCall.Id}");

            Console.WriteLine(
                $"State   : {statefulCall.Resource?.State}");

            Console.WriteLine(
                "================================================");

            Console.WriteLine();
            Console.WriteLine(
                "JOIN REQUEST SENT TO MICROSOFT GRAPH");

            Console.WriteLine(
                $"Call ID : {statefulCall.Id}");

            Console.WriteLine(
                $"State   : {statefulCall.Resource?.State}");

            return statefulCall;
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine(
                "================================================");

            Console.WriteLine(
                " SERVICE-HOSTED MEETING JOIN FAILURE");

            Console.WriteLine(
                "================================================");

            Console.WriteLine(
                ex.Message);

            Console.WriteLine(
                ex);

            throw;
        }
    }

    // ============================================================
    // EXISTING CALLBACK COMPATIBILITY METHOD
    //
    // Your MeetingMediaHandler may still call this method when
    // a call becomes established.
    //
    // In service-hosted mode there is NO AudioSocket to start.
    // ============================================================

    public Task StartMediaSessionAsync(
        string callId)
    {
        if (string.IsNullOrWhiteSpace(
                callId))
        {
            return Task.CompletedTask;
        }

        Console.WriteLine();
        Console.WriteLine(
            "================================================");

        Console.WriteLine(
            " SERVICE-HOSTED CALL ESTABLISHED");

        Console.WriteLine(
            "================================================");

        Console.WriteLine(
            $"Call ID : {callId}");

        Console.WriteLine(
            "Media   : Hosted by Microsoft");

        Console.WriteLine(
            "AudioSocket : NOT APPLICABLE");

        Console.WriteLine(
            "Next step   : recordResponse / playPrompt");

        Console.WriteLine(
            "================================================");

        return Task.CompletedTask;
    }

    // ============================================================
    // GRAPH CALLBACK PROCESSING
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
    // CALL STATE
    // ============================================================

    private void HandleCallStateChange(
        ICall call)
    {
        var state =
            call.Resource?.State;

        Console.WriteLine(
            $"[CALL] {call.Id} state={state}");

        if (state ==
            CallState.Established)
        {
            Console.WriteLine();
            Console.WriteLine(
                "================================================");

            Console.WriteLine(
                "     TEAMS CALL ESTABLISHED");

            Console.WriteLine(
                "================================================");

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
            Console.WriteLine(
                "================================================");

            Console.WriteLine(
                "       TEAMS CALL TERMINATED");

            Console.WriteLine(
                "================================================");

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

            Console.WriteLine(
                "================================================");
        }
    }

    // ============================================================
    // MEETING ID NORMALIZER
    //
    // Example:
    //
    // 245 589 142 346 08
    //
    // becomes:
    //
    // 24558914234608
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
using System.Collections.Concurrent;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using AgentTeamMateBot.Media;
using Microsoft.Graph.Communications.Calls;
using Microsoft.Graph.Communications.Calls.Media;
using Microsoft.Graph.Communications.Client;
using Microsoft.Graph.Communications.Common.Telemetry;
using Microsoft.Graph.Communications.Resources;
using Microsoft.Graph.Contracts;
using Microsoft.Graph.Models;
using Microsoft.Skype.Bots.Media;

namespace AgentTeamMateBot.Services;

public class MediaSessionService
{
    private readonly IConfiguration _configuration;
    private readonly GraphAuthService _graphAuthService;
    private readonly AudioHandler _audioHandler;
    private readonly SpeechRecognitionService _speechService;

    private readonly ConcurrentDictionary<string, CallMediaState> _calls = new();
    private readonly ConcurrentDictionary<Guid, AudioSocketBinding> _audioBindings = new();
    private readonly object _initLock = new();

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
        _configuration = configuration;
        _graphAuthService = graphAuthService;
        _audioHandler = audioHandler;
        _speechService = speechService;
    }

    public ICommunicationsClient? Client => _client;

    public bool IsInitialized => _initialized;

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
                Console.WriteLine(" INITIALIZING GRAPH MEDIA PLATFORM");
                Console.WriteLine("================================================");

                var clientId =
                    _graphAuthService.ClientId;

                var callbackUri =
                    _configuration["Bot:CallbackUri"]
                    ?? "https://teammate-bot.westus3.cloudapp.azure.com/api/calling";

                var serviceFqdn =
                    _configuration["Bot:ServiceFqdn"]
                    ?? _configuration["Media:ServiceFqdn"]
                    ?? new Uri(callbackUri).Host;

                var internalPort =
                    _configuration.GetValue(
                        "Media:InternalPort",
                        8445);

                var publicPort =
                    _configuration.GetValue(
                        "Media:PublicPort",
                        8445);

                var portMin =
                    _configuration.GetValue(
                        "Media:PortRangeMin",
                        20000);

                var portMax =
                    _configuration.GetValue(
                        "Media:PortRangeMax",
                        20999);

                var certificate =
                    LoadCertificate();

                var mediaBindAddress =
                    IPAddress.Any;

                var instanceSettings =
                    new MediaPlatformInstanceSettings
                    {
                        ServiceFqdn =
                            serviceFqdn,

                        InstancePublicIPAddress =
                            mediaBindAddress,

                        InstanceInternalPort =
                            internalPort,

                        InstancePublicPort =
                            publicPort,

                        MediaPortRange =
                            new PortRange(
                                (uint)portMin,
                                (uint)portMax)
                    };

                if (certificate != null)
                {
                    instanceSettings.Certificate =
                        certificate;

                    Console.WriteLine(
                        $"Certificate subject : {certificate.Subject}");

                    Console.WriteLine(
                        $"Certificate thumbprint : {certificate.Thumbprint}");
                }
                else
                {
                    throw new InvalidOperationException(
                        "Media platform certificate is missing.");
                }

                // ============================================================
                // MEDIA SDK LOGGER
                // ============================================================

                var loggerFactory =
                    LoggerFactory.Create(
                        logging =>
                        {
                            logging.AddConsole();

                            logging.SetMinimumLevel(
                                Microsoft.Extensions.Logging.LogLevel.Trace);
                        });

                var mediaLogger =
                    new BotMediaLogger(
                        loggerFactory.CreateLogger<BotMediaLogger>());

                var mediaPlatformSettings =
                    new MediaPlatformSettings
                    {
                        ApplicationId =
                            clientId,

                        MediaPlatformInstanceSettings =
                            instanceSettings,

                        MediaPlatformLogger =
                            mediaLogger
                    };

                // ============================================================
                // GRAPH LOGGER
                // ============================================================

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

                builder.SetMediaPlatformSettings(
                    mediaPlatformSettings);

                _client =
                    builder.Build();

                _client.Calls().OnUpdated +=
                    OnCallsUpdated;

                _initialized =
                    true;

                _initError =
                    null;

                Console.WriteLine(
                    $"Service FQDN     : {serviceFqdn}");

                Console.WriteLine(
                    $"Media bind IP    : {mediaBindAddress} (IPAddress.Any)");

                Console.WriteLine(
                    $"Control port     : {internalPort} (internal) / {publicPort} (public)");

                Console.WriteLine(
                    $"Media UDP ports  : {portMin}-{portMax}");

                Console.WriteLine(
                    $"Notification URL : {callbackUri}");

                Console.WriteLine(
                    "Media SDK logger  : ENABLED");

                Console.WriteLine(
                    "MEDIA PLATFORM INITIALIZED");

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
                    " MEDIA PLATFORM INITIALIZATION FAILURE");

                Console.WriteLine(
                    "================================================");

                Console.WriteLine(
                    ex.Message);

                Console.WriteLine(
                    ex);
            }
        }
    }

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
                "Media platform is not initialized. " +
                (_initError ??
                 "See initialization logs."));
        }

        ILocalMediaSession? mediaSession =
            null;

        try
        {
            mediaSession =
                CreateLocalMediaSession();

            var audioBinding =
                BindAudioSocket(
                    mediaSession);

            var mediaConfiguration =
                mediaSession
                    .GetMediaConfiguration();

            Console.WriteLine();
            Console.WriteLine(
                "================================================");

            Console.WriteLine(
                " GENERATED MEDIA CONFIGURATION");

            Console.WriteLine(
                "================================================");

            Console.WriteLine(
                mediaConfiguration.ToString(
                    Newtonsoft.Json.Formatting.Indented));

            Console.WriteLine(
                "================================================");

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
                        new AppHostedMediaConfig
                        {
                            OdataType =
                                "#microsoft.graph.appHostedMediaConfig",

                            Blob =
                                mediaConfiguration.ToString(
                                    Newtonsoft.Json.Formatting.None)
                        },

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
                " JOINING TEAMS MEETING WITH APP-HOSTED MEDIA");

            Console.WriteLine(
                "================================================");

            Console.WriteLine(
                $"Meeting ID : {normalizedMeetingId}");

            Console.WriteLine(
                $"Tenant ID  : {tenantId}");

            var statefulCall =
                await _client
                    .Calls()
                    .AddAsync(
                        call,
                        mediaSession);

            audioBinding.CallId =
                statefulCall.Id;

            TrackCall(
                statefulCall,
                mediaSession,
                audioBinding);

            Console.WriteLine(
                $"Call ID    : {statefulCall.Id}");

            Console.WriteLine(
                "Media session created and AudioSocket subscribed.");

            Console.WriteLine(
                "================================================");

            return statefulCall;
        }
        catch (Exception ex)
        {
            mediaSession?.Dispose();

            Console.WriteLine();
            Console.WriteLine(
                "================================================");

            Console.WriteLine(
                " MEDIA SESSION CREATION / JOIN FAILURE");

            Console.WriteLine(
                "================================================");

            Console.WriteLine(
                ex.Message);

            Console.WriteLine(
                ex);

            throw;
        }
    }

    public async Task StartMediaSessionAsync(
        string callId)
    {
        if (string.IsNullOrWhiteSpace(
                callId))
        {
            return;
        }

        if (!_initialized ||
            _client == null)
        {
            Console.WriteLine(
                "[MEDIA] Media platform is not initialized.");

            return;
        }

        if (_calls.TryGetValue(
                callId,
                out var state))
        {
            await OnCallEstablishedAsync(
                state);

            return;
        }

        ICall? existingCall =
            null;

        try
        {
            existingCall =
                _client.Calls()[callId];
        }
        catch (Exception ex)
        {
            Console.WriteLine(
                $"[MEDIA] Cannot find call {callId}: {ex.Message}");
        }

        if (existingCall?.MediaSession
            is ILocalMediaSession localSession)
        {
            TrackCall(
                existingCall,
                localSession,
                BindAudioSocket(
                    localSession,
                    callId));

            await OnCallEstablishedAsync(
                _calls[callId]);

            return;
        }

        Console.WriteLine(
            "[MEDIA] Media session cannot be attached after call creation.");
    }

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
                "[MEDIA] Notification processing failure:");

            Console.WriteLine(
                ex);

            return new HttpResponseMessage(
                HttpStatusCode.InternalServerError);
        }
    }

    private ILocalMediaSession
        CreateLocalMediaSession()
    {
        if (_client == null)
        {
            throw new InvalidOperationException(
                "Communications client is not initialized.");
        }

        try
        {
            var audioSocketSettings =
                new AudioSocketSettings
                {
                    StreamDirections =
                        StreamDirection.Sendrecv,

                    SupportedAudioFormat =
                        AudioFormat.Pcm16K,

                    ReceiveUnmixedMeetingAudio =
                        false
                };

            var videoSocketSettings =
                new VideoSocketSettings
                {
                    StreamDirections =
                        StreamDirection.Inactive
                };

            Console.WriteLine(
                "[MEDIA] Audio direction : Sendrecv");

            Console.WriteLine(
                "[MEDIA] Video direction : Inactive");

            return _client.CreateMediaSession(
                audioSocketSettings,
                videoSocketSettings);
        }
        catch (Exception ex)
        {
            Console.WriteLine(
                "[MEDIA] Media session creation failure:");

            Console.WriteLine(
                ex);

            throw;
        }
    }

    private AudioSocketBinding
        BindAudioSocket(
            ILocalMediaSession mediaSession,
            string? callId = null)
    {
        if (_audioBindings.TryGetValue(
                mediaSession.MediaSessionId,
                out var existing))
        {
            if (!string.IsNullOrWhiteSpace(
                    callId))
            {
                existing.CallId =
                    callId;
            }

            return existing;
        }

        var audioSocket =
            mediaSession.AudioSocket
            ?? throw new InvalidOperationException(
                "Media session does not contain an AudioSocket.");

        var binding =
            new AudioSocketBinding
            {
                CallId =
                    callId
                    ?? mediaSession
                        .MediaSessionId
                        .ToString()
            };

        binding.AudioReceivedHandler =
            (_, args) =>
                OnAudioMediaReceived(
                    binding.CallId,
                    args);

        binding.MediaFailureHandler =
            (_, args) =>
            {
                Console.WriteLine();
                Console.WriteLine(
                    "================================================");

                Console.WriteLine(
                    " AUDIO SOCKET FAILURE");

                Console.WriteLine(
                    "================================================");

                Console.WriteLine(
                    $"Call / session : {binding.CallId}");

                Console.WriteLine(
                    args?.ToString());
            };

        audioSocket.AudioMediaReceived +=
            binding.AudioReceivedHandler;

        audioSocket.MediaStreamFailure +=
            binding.MediaFailureHandler;

        _audioBindings[
            mediaSession.MediaSessionId] =
            binding;

        Console.WriteLine(
            "AudioSocket event AudioMediaReceived subscribed.");

        return binding;
    }

    private void TrackCall(
        ICall call,
        ILocalMediaSession mediaSession,
        AudioSocketBinding audioBinding)
    {
        audioBinding.CallId =
            call.Id;

        if (_calls.TryGetValue(
                call.Id,
                out var existing))
        {
            existing.AudioBinding =
                audioBinding;

            return;
        }

        var state =
            new CallMediaState(
                call,
                mediaSession)
            {
                AudioBinding =
                    audioBinding
            };

        _calls[
            call.Id] =
            state;

        call.OnUpdated +=
            OnCallResourceUpdated;
    }

    private void OnCallsUpdated(
        ICallCollection sender,
        CollectionEventArgs<ICall> args)
    {
        foreach (var call
                 in args.AddedResources)
        {
            Console.WriteLine(
                $"[CALL] Added {call.Id} state={call.Resource?.State}");

            if (!_calls.ContainsKey(
                    call.Id))
            {
                call.OnUpdated +=
                    OnCallResourceUpdated;
            }
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
                    out var state))
            {
                _audioBindings.TryRemove(
                    state.MediaSession.MediaSessionId,
                    out _);

                state.Dispose();
            }
        }
    }

    private void OnCallResourceUpdated(
        ICall sender,
        ResourceEventArgs<Call> args)
    {
        HandleCallStateChange(
            sender);
    }

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
            if (_calls.TryGetValue(
                    call.Id,
                    out var mediaState))
            {
                _ =
                    OnCallEstablishedAsync(
                        mediaState);
            }
            else if (call.MediaSession
                     is ILocalMediaSession localSession)
            {
                TrackCall(
                    call,
                    localSession,
                    BindAudioSocket(
                        localSession,
                        call.Id));

                _ =
                    OnCallEstablishedAsync(
                        _calls[call.Id]);
            }
        }
    }

    private async Task
        OnCallEstablishedAsync(
            CallMediaState state)
    {
        if (state.EstablishedHandled)
        {
            return;
        }

        state.EstablishedHandled =
            true;

        Console.WriteLine();
        Console.WriteLine(
            "================================================");

        Console.WriteLine(
            " MEDIA SESSION STARTED");

        Console.WriteLine(
            "================================================");

        Console.WriteLine(
            $"Call ID : {state.Call.Id}");

        Console.WriteLine(
            state.MediaSession.AudioSocket != null
                ? "AudioSocket connected"
                : "AudioSocket missing");

        try
        {
            await _speechService
                .StartAsync();
        }
        catch (Exception ex)
        {
            Console.WriteLine(
                "[SPEECH] Startup failure:");

            Console.WriteLine(
                ex);
        }
    }

    private void OnAudioMediaReceived(
        string callId,
        AudioMediaReceivedEventArgs args)
    {
        var buffer =
            args.Buffer;

        if (buffer == null)
        {
            return;
        }

        try
        {
            var length =
                (int)buffer.Length;

            if (length <= 0 ||
                buffer.Data == IntPtr.Zero)
            {
                return;
            }

            var audioBytes =
                new byte[length];

            Marshal.Copy(
                buffer.Data,
                audioBytes,
                0,
                length);

            _audioHandler.ProcessAudio(
                audioBytes,
                buffer.AudioFormat,
                buffer.IsSilence,
                callId);
        }
        catch (Exception ex)
        {
            Console.WriteLine(
                "[AUDIO SOCKET] " +
                ex.Message);
        }
        finally
        {
            buffer.Dispose();
        }
    }

    private static string NormalizeMeetingId(
        string meetingId)
    {
        return new string(
            meetingId
                .Where(char.IsDigit)
                .ToArray());
    }

    private X509Certificate2?
        LoadCertificate()
    {
        var path =
            _configuration[
                "Media:CertificatePath"];

        var password =
            _configuration[
                "Media:CertificatePassword"];

        var thumbprint =
            _configuration[
                "Media:CertificateThumbprint"];

        if (!string.IsNullOrWhiteSpace(
                path))
        {
            return new X509Certificate2(
                path,
                password);
        }

        if (!string.IsNullOrWhiteSpace(
                thumbprint))
        {
            thumbprint =
                thumbprint.Replace(
                    " ",
                    string.Empty,
                    StringComparison.Ordinal);

            foreach (var location
                     in new[]
                     {
                         StoreLocation.LocalMachine,
                         StoreLocation.CurrentUser
                     })
            {
                using var store =
                    new X509Store(
                        StoreName.My,
                        location);

                store.Open(
                    OpenFlags.ReadOnly);

                var certs =
                    store.Certificates.Find(
                        X509FindType.FindByThumbprint,
                        thumbprint,
                        validOnly: false);

                if (certs.Count > 0)
                {
                    Console.WriteLine(
                        $"Loaded media certificate from {location}\\My");

                    return certs[0];
                }
            }
        }

        var host =
            _configuration[
                "Bot:ServiceFqdn"]
            ?? _configuration[
                "Media:ServiceFqdn"]
            ?? "teammate-bot.westus3.cloudapp.azure.com";

        return FindCertificateByHost(
            host);
    }

    private static X509Certificate2?
        FindCertificateByHost(
            string host)
    {
        foreach (var location
                 in new[]
                 {
                     StoreLocation.LocalMachine,
                     StoreLocation.CurrentUser
                 })
        {
            using var store =
                new X509Store(
                    StoreName.My,
                    location);

            store.Open(
                OpenFlags.ReadOnly);

            foreach (var cert
                     in store.Certificates)
            {
                if (cert.HasPrivateKey &&
                    CertificateMatchesHost(
                        cert,
                        host))
                {
                    Console.WriteLine(
                        $"Loaded media certificate for {host} from {location}\\My");

                    return cert;
                }
            }
        }

        return null;
    }

    private static bool CertificateMatchesHost(
        X509Certificate2 cert,
        string host)
    {
        if (cert.Subject.Contains(
                host,
                StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        foreach (var extension
                 in cert.Extensions)
        {
            if (extension.Oid?.Value
                != "2.5.29.17")
            {
                continue;
            }

            var formatted =
                extension.Format(
                    true);

            if (formatted.Contains(
                    host,
                    StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private sealed class CallMediaState
        : IDisposable
    {
        public CallMediaState(
            ICall call,
            ILocalMediaSession mediaSession)
        {
            Call =
                call;

            MediaSession =
                mediaSession;
        }

        public ICall Call
        {
            get;
        }

        public ILocalMediaSession MediaSession
        {
            get;
        }

        public AudioSocketBinding? AudioBinding
        {
            get;
            set;
        }

        public bool EstablishedHandled
        {
            get;
            set;
        }

        public void Dispose()
        {
            var audioSocket =
                MediaSession.AudioSocket;

            if (audioSocket != null &&
                AudioBinding != null)
            {
                if (AudioBinding.AudioReceivedHandler != null)
                {
                    audioSocket.AudioMediaReceived -=
                        AudioBinding.AudioReceivedHandler;
                }

                if (AudioBinding.MediaFailureHandler != null)
                {
                    audioSocket.MediaStreamFailure -=
                        AudioBinding.MediaFailureHandler;
                }
            }

            MediaSession.Dispose();
        }
    }

    private sealed class AudioSocketBinding
    {
        public string CallId
        {
            get;
            set;
        } = string.Empty;

        public EventHandler<AudioMediaReceivedEventArgs>?
            AudioReceivedHandler
        {
            get;
            set;
        }

        public EventHandler<MediaStreamFailureEventArgs>?
            MediaFailureHandler
        {
            get;
            set;
        }
    }
}
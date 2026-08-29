using System.Collections.Concurrent;
using System.Net;
using System.Net.Sockets;
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

                var clientId = _graphAuthService.ClientId;
                var callbackUri = _configuration["Bot:CallbackUri"]
                    ?? "https://teammate-bot.westus3.cloudapp.azure.com/api/calling";
                var serviceFqdn = _configuration["Bot:ServiceFqdn"]
                    ?? _configuration["Media:ServiceFqdn"]
                    ?? new Uri(callbackUri).Host;

                var internalPort = _configuration.GetValue("Media:InternalPort", 8445);
                var publicPort = _configuration.GetValue("Media:PublicPort", 8445);
                var portMin = _configuration.GetValue("Media:PortRangeMin", 20000);
                var portMax = _configuration.GetValue("Media:PortRangeMax", 20039);

                var publicIp = ResolvePublicIp(serviceFqdn);
                var certificate = LoadCertificate();

                var instanceSettings = new MediaPlatformInstanceSettings
                {
                    ServiceFqdn = serviceFqdn,
                    InstancePublicIPAddress = publicIp,
                    InstanceInternalPort = internalPort,
                    InstancePublicPort = publicPort,
                    MediaPortRange = new PortRange((uint)portMin, (uint)portMax)
                };

                if (certificate != null)
                {
                    instanceSettings.Certificate = certificate;
                    Console.WriteLine($"Certificate subject : {certificate.Subject}");
                    Console.WriteLine($"Certificate thumbprint : {certificate.Thumbprint}");
                }
                else
                {
                    throw new InvalidOperationException(
                        "Media platform certificate is missing. Set Media:CertificateThumbprint or Media:CertificatePath, " +
                        "or install an SSL certificate for the service FQDN in LocalMachine\\My. " +
                        "The Graph Media SDK requires an SSL certificate for MTLS with Teams.");
                }

                var mediaPlatformSettings = new MediaPlatformSettings
                {
                    ApplicationId = clientId,
                    MediaPlatformInstanceSettings = instanceSettings
                };

                _graphLogger = new GraphLogger("AgentTeamMateBot", redirectToTrace: true);

                var builder = new CommunicationsClientBuilder(
                    "AgentTeamMateBot",
                    clientId,
                    _graphLogger);

#pragma warning disable CS0618
                builder.SetAuthenticationProvider(
                    _graphAuthService.CreateAuthenticationProvider(_graphLogger));
#pragma warning restore CS0618
                builder.SetNotificationUrl(new Uri(callbackUri));
                builder.SetServiceBaseUrl(new Uri("https://graph.microsoft.com/v1.0"));
                builder.SetMediaPlatformSettings(mediaPlatformSettings);

                _client = builder.Build();
                _client.Calls().OnUpdated += OnCallsUpdated;

                _initialized = true;
                _initError = null;

                Console.WriteLine($"Service FQDN     : {serviceFqdn}");
                Console.WriteLine($"Public IP        : {publicIp}");
                Console.WriteLine($"Control port     : {internalPort} (internal) / {publicPort} (public)");
                Console.WriteLine($"Media UDP ports  : {portMin}-{portMax}");
                Console.WriteLine("Notification URL  : " + callbackUri);
                Console.WriteLine("MEDIA PLATFORM INITIALIZED");
                Console.WriteLine("================================================");
            }
            catch (Exception ex)
            {
                _initialized = false;
                _initError = ex.Message;

                Console.WriteLine();
                Console.WriteLine("================================================");
                Console.WriteLine(" MEDIA PLATFORM INITIALIZATION FAILURE");
                Console.WriteLine("================================================");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex);
                Console.WriteLine();
                Console.WriteLine("The bot cannot receive real Teams audio until MediaPlatform starts.");
                Console.WriteLine("Required: Windows Server x64, VC++ x64 runtime, SSL cert,");
                Console.WriteLine("and NSG rules for the media control TCP port plus the UDP media port range.");
            }
        }
    }

    public async Task<ICall> JoinMeetingAsync(string meetingId, string? passcode)
    {
        if (!_initialized || _client == null)
        {
            Initialize();
        }

        if (!_initialized || _client == null)
        {
            throw new InvalidOperationException(
                "Media platform is not initialized. " + (_initError ?? "See MEDIA PLATFORM INITIALIZATION FAILURE logs."));
        }

        ILocalMediaSession? mediaSession = null;

        try
        {
            mediaSession = CreateLocalMediaSession();
            var audioBinding = BindAudioSocket(mediaSession);

            var mediaConfiguration = mediaSession.GetMediaConfiguration();

            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" GENERATED MEDIA CONFIGURATION");
            Console.WriteLine("================================================");
            Console.WriteLine(
                mediaConfiguration.ToString(Newtonsoft.Json.Formatting.Indented));
            Console.WriteLine("================================================");

            var tenantId = _graphAuthService.TenantId.Trim();
            var normalizedMeetingId = NormalizeMeetingId(meetingId);
            var normalizedPasscode = string.IsNullOrWhiteSpace(passcode) ? null : passcode.Trim();

            var applicationIdentity = new Identity
            {
                Id = _graphAuthService.ClientId,
                DisplayName = "Agent Team Mate"
            };
            applicationIdentity.SetTenantId(tenantId);

            var call = new Call
            {
                OdataType = "#microsoft.graph.call",
                CallbackUri = _configuration["Bot:CallbackUri"]
                    ?? "https://teammate-bot.westus3.cloudapp.azure.com/api/calling",
                TenantId = tenantId,
                Source = new ParticipantInfo
                {
                    Identity = new IdentitySet
                    {
                        Application = applicationIdentity
                    }
                },
                RequestedModalities = new List<Modality?> { Modality.Audio },
                MediaConfig = new AppHostedMediaConfig
                {
                    OdataType = "#microsoft.graph.appHostedMediaConfig",
                    Blob = mediaConfiguration.ToString(Newtonsoft.Json.Formatting.None)
                },
                MeetingInfo = new JoinMeetingIdMeetingInfo
                {
                    OdataType = "#microsoft.graph.joinMeetingIdMeetingInfo",
                    JoinMeetingId = normalizedMeetingId,
                    Passcode = normalizedPasscode
                }
            };

            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" JOINING TEAMS MEETING WITH APP-HOSTED MEDIA");
            Console.WriteLine("================================================");
            Console.WriteLine($"Meeting ID : {normalizedMeetingId}");
            Console.WriteLine($"Tenant ID  : {tenantId}");

            var statefulCall = await _client.Calls().AddAsync(call, mediaSession);
            audioBinding.CallId = statefulCall.Id;
            TrackCall(statefulCall, mediaSession, audioBinding);

            Console.WriteLine($"Call ID    : {statefulCall.Id}");
            Console.WriteLine("Media session created and AudioSocket subscribed before call negotiation.");
            Console.WriteLine("================================================");

            return statefulCall;
        }
        catch (Exception ex)
        {
            mediaSession?.Dispose();

            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" MEDIA SESSION CREATION / JOIN FAILURE");
            Console.WriteLine("================================================");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex);
            throw;
        }
    }

    public async Task StartMediaSessionAsync(string callId)
    {
        if (string.IsNullOrWhiteSpace(callId))
        {
            Console.WriteLine("[MEDIA] Cannot start media session. Call ID is missing.");
            return;
        }

        if (!_initialized || _client == null)
        {
            Console.WriteLine("[MEDIA] Media platform is not initialized. Real AudioSocket audio is unavailable.");
            return;
        }

        if (_calls.TryGetValue(callId, out var state))
        {
            await OnCallEstablishedAsync(state);
            return;
        }

        ICall? existingCall = null;
        try
        {
            existingCall = _client.Calls()[callId];
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[MEDIA] Call {callId} is not in the CommunicationsClient collection: {ex.Message}");
        }

        if (existingCall?.MediaSession is ILocalMediaSession localSession)
        {
            TrackCall(existingCall, localSession, BindAudioSocket(localSession, callId));
            await OnCallEstablishedAsync(_calls[callId]);
            return;
        }

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" MEDIA SESSION CANNOT BE ATTACHED AFTER THE FACT");
        Console.WriteLine("================================================");
        Console.WriteLine($"Call ID : {callId}");
        Console.WriteLine("AudioSocket must be created before Graph negotiates the call.");
        Console.WriteLine("Join through POST /api/join so the bot uses AppHostedMediaConfig.");
        Console.WriteLine("Service-hosted media never delivers PCM bytes to this VM.");
    }

    public async Task<HttpResponseMessage> ProcessNotificationAsync(HttpRequestMessage requestMessage)
    {
        if (_client == null)
        {
            Console.WriteLine("[MEDIA] Ignoring calling notification because CommunicationsClient is not initialized.");
            return new HttpResponseMessage(HttpStatusCode.ServiceUnavailable);
        }

        try
        {
            return await _client.ProcessNotificationAsync(requestMessage);
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" CALLING NOTIFICATION PROCESSING FAILURE");
            Console.WriteLine("================================================");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex);
            return new HttpResponseMessage(HttpStatusCode.InternalServerError);
        }
    }

    private ILocalMediaSession CreateLocalMediaSession()
    {
        if (_client == null)
        {
            throw new InvalidOperationException("Communications client is not initialized.");
        }

        try
        {
            var audioSocketSettings = new AudioSocketSettings
            {
                StreamDirections = StreamDirection.Sendrecv,
                SupportedAudioFormat = AudioFormat.Pcm16K,
                ReceiveUnmixedMeetingAudio = false
            };

            return _client.CreateMediaSession(audioSocketSettings);
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" MEDIA SESSION CREATION FAILURE");
            Console.WriteLine("================================================");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex);
            throw;
        }
    }

    private AudioSocketBinding BindAudioSocket(ILocalMediaSession mediaSession, string? callId = null)
    {
        if (_audioBindings.TryGetValue(mediaSession.MediaSessionId, out var existing))
        {
            if (!string.IsNullOrWhiteSpace(callId))
            {
                existing.CallId = callId;
            }

            return existing;
        }

        var audioSocket = mediaSession.AudioSocket
            ?? throw new InvalidOperationException("Media session does not contain an AudioSocket.");

        var binding = new AudioSocketBinding
        {
            CallId = callId ?? mediaSession.MediaSessionId.ToString()
        };

        binding.AudioReceivedHandler = (_, args) => OnAudioMediaReceived(binding.CallId, args);
        binding.MediaFailureHandler = (_, args) =>
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" AUDIO SOCKET FAILURE");
            Console.WriteLine("================================================");
            Console.WriteLine($"Call / session : {binding.CallId}");
            Console.WriteLine(args?.ToString());
        };

        audioSocket.AudioMediaReceived += binding.AudioReceivedHandler;
        audioSocket.MediaStreamFailure += binding.MediaFailureHandler;
        _audioBindings[mediaSession.MediaSessionId] = binding;

        Console.WriteLine("AudioSocket event AudioMediaReceived subscribed.");
        return binding;
    }

    private void TrackCall(ICall call, ILocalMediaSession mediaSession, AudioSocketBinding audioBinding)
    {
        audioBinding.CallId = call.Id;

        if (_calls.TryGetValue(call.Id, out var existing))
        {
            existing.AudioBinding = audioBinding;
            return;
        }

        var state = new CallMediaState(call, mediaSession)
        {
            AudioBinding = audioBinding
        };

        _calls[call.Id] = state;
        call.OnUpdated += OnCallResourceUpdated;
    }

    private void OnCallsUpdated(ICallCollection sender, CollectionEventArgs<ICall> args)
    {
        foreach (var call in args.AddedResources)
        {
            Console.WriteLine($"[CALL] Added {call.Id} state={call.Resource?.State}");
            if (!_calls.ContainsKey(call.Id))
            {
                call.OnUpdated += OnCallResourceUpdated;
            }
        }

        foreach (var call in args.UpdatedResources)
        {
            HandleCallStateChange(call);
        }

        foreach (var call in args.RemovedResources)
        {
            Console.WriteLine($"[CALL] Removed {call.Id}");
            if (_calls.TryRemove(call.Id, out var state))
            {
                _audioBindings.TryRemove(state.MediaSession.MediaSessionId, out _);
                state.Dispose();
            }
        }
    }

    private void OnCallResourceUpdated(ICall sender, ResourceEventArgs<Call> args)
    {
        HandleCallStateChange(sender);
    }

    private void HandleCallStateChange(ICall call)
    {
        var state = call.Resource?.State;
        Console.WriteLine($"[CALL] {call.Id} state={state}");

        if (state == CallState.Established)
        {
            if (_calls.TryGetValue(call.Id, out var mediaState))
            {
                _ = OnCallEstablishedAsync(mediaState);
            }
            else if (call.MediaSession is ILocalMediaSession localSession)
            {
                TrackCall(call, localSession, BindAudioSocket(localSession, call.Id));
                _ = OnCallEstablishedAsync(_calls[call.Id]);
            }
            else
            {
                _ = StartMediaSessionAsync(call.Id);
            }
        }
    }

    private async Task OnCallEstablishedAsync(CallMediaState state)
    {
        if (state.EstablishedHandled)
        {
            return;
        }

        state.EstablishedHandled = true;

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine("MEDIA SESSION STARTED");
        Console.WriteLine("================================================");
        Console.WriteLine();
        Console.WriteLine("Call ID:");
        Console.WriteLine(state.Call.Id);
        Console.WriteLine();
        Console.WriteLine(
            state.MediaSession.AudioSocket != null
                ? "AudioSocket connected"
                : "AudioSocket missing");
        Console.WriteLine();

        try
        {
            await _speechService.StartAsync();
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" SPEECH SDK FAILURE");
            Console.WriteLine("================================================");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex);
        }
    }

    private void OnAudioMediaReceived(string callId, AudioMediaReceivedEventArgs args)
    {
        var buffer = args.Buffer;
        if (buffer == null)
        {
            return;
        }

        try
        {
            var length = (int)buffer.Length;
            if (length <= 0 || buffer.Data == IntPtr.Zero)
            {
                return;
            }

            var audioBytes = new byte[length];
            Marshal.Copy(buffer.Data, audioBytes, 0, length);

            _audioHandler.ProcessAudio(
                audioBytes,
                buffer.AudioFormat,
                buffer.IsSilence,
                callId);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[AUDIO SOCKET] Failed to read Teams media buffer: {ex.Message}");
        }
        finally
        {
            buffer.Dispose();
        }
    }

    private IPAddress ResolvePublicIp(string serviceFqdn)
    {
        var configured = _configuration["Media:PublicIpAddress"];
        if (!string.IsNullOrWhiteSpace(configured))
        {
            return IPAddress.Parse(configured);
        }

        try
        {
            var addresses = Dns.GetHostAddresses(serviceFqdn);
            var ipv4 = Array.Find(addresses, a => a.AddressFamily == AddressFamily.InterNetwork);
            if (ipv4 != null)
            {
                return ipv4;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[MEDIA] DNS lookup for {serviceFqdn} failed: {ex.Message}");
        }

        throw new InvalidOperationException(
            "Media:PublicIpAddress is required when the service FQDN cannot be resolved to a public IPv4 address.");
    }

    private static string NormalizeMeetingId(string meetingId)
    {
        return new string(meetingId.Where(char.IsDigit).ToArray());
    }

    private X509Certificate2? LoadCertificate()
    {
        var path = _configuration["Media:CertificatePath"];
        var password = _configuration["Media:CertificatePassword"];
        var thumbprint = _configuration["Media:CertificateThumbprint"];

        if (!string.IsNullOrWhiteSpace(path))
        {
            Console.WriteLine($"Loading media certificate from {path}");
            return new X509Certificate2(path, password);
        }

        if (!string.IsNullOrWhiteSpace(thumbprint))
        {
            thumbprint = thumbprint.Replace(" ", string.Empty, StringComparison.Ordinal);

            foreach (var location in new[] { StoreLocation.LocalMachine, StoreLocation.CurrentUser })
            {
                using var store = new X509Store(StoreName.My, location);
                store.Open(OpenFlags.ReadOnly);
                var certs = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, validOnly: false);
                if (certs.Count > 0)
                {
                    Console.WriteLine($"Loaded media certificate from {location}\\My");
                    return certs[0];
                }
            }

            throw new InvalidOperationException(
                $"Certificate with thumbprint {thumbprint} was not found in LocalMachine\\My or CurrentUser\\My.");
        }

        var host = _configuration["Bot:ServiceFqdn"]
            ?? _configuration["Media:ServiceFqdn"]
            ?? "teammate-bot.westus3.cloudapp.azure.com";

        return FindCertificateByHost(host);
    }

    private static X509Certificate2? FindCertificateByHost(string host)
    {
        foreach (var location in new[] { StoreLocation.LocalMachine, StoreLocation.CurrentUser })
        {
            using var store = new X509Store(StoreName.My, location);
            store.Open(OpenFlags.ReadOnly);

            foreach (var cert in store.Certificates)
            {
                if (cert.HasPrivateKey && CertificateMatchesHost(cert, host))
                {
                    Console.WriteLine($"Loaded media certificate for {host} from {location}\\My");
                    return cert;
                }
            }
        }

        return null;
    }

    private static bool CertificateMatchesHost(X509Certificate2 cert, string host)
    {
        if (cert.Subject.Contains(host, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        foreach (var extension in cert.Extensions)
        {
            if (extension.Oid?.Value != "2.5.29.17")
            {
                continue;
            }

            var formatted = extension.Format(true);
            if (formatted.Contains(host, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private sealed class CallMediaState : IDisposable
    {
        public CallMediaState(ICall call, ILocalMediaSession mediaSession)
        {
            Call = call;
            MediaSession = mediaSession;
        }

        public ICall Call { get; }

        public ILocalMediaSession MediaSession { get; }

        public AudioSocketBinding? AudioBinding { get; set; }

        public bool EstablishedHandled { get; set; }

        public void Dispose()
        {
            var audioSocket = MediaSession.AudioSocket;
            if (audioSocket != null && AudioBinding != null)
            {
                if (AudioBinding.AudioReceivedHandler != null)
                {
                    audioSocket.AudioMediaReceived -= AudioBinding.AudioReceivedHandler;
                }

                if (AudioBinding.MediaFailureHandler != null)
                {
                    audioSocket.MediaStreamFailure -= AudioBinding.MediaFailureHandler;
                }
            }

            MediaSession.Dispose();
        }
    }

    private sealed class AudioSocketBinding
    {
        public string CallId { get; set; } = string.Empty;

        public EventHandler<AudioMediaReceivedEventArgs>? AudioReceivedHandler { get; set; }

        public EventHandler<MediaStreamFailureEventArgs>? MediaFailureHandler { get; set; }
    }
}

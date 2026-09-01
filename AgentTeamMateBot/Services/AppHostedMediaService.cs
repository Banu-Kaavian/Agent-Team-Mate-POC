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

public class AppHostedMediaService
{
    private readonly IConfiguration _configuration;
    private readonly GraphAuthService _graphAuthService;
    private readonly AudioHandler _audioHandler;
    private readonly SpeechRecognitionService _speechService;
    private readonly IBotMediaLogger _mediaLogger;

    private readonly ConcurrentDictionary<string, ICall> _calls = new();
    private readonly ConcurrentDictionary<string, ILocalMediaSession> _mediaSessions = new();
    private readonly ConcurrentDictionary<Guid, AudioSocketBinding> _audioBindings = new();
    private readonly object _initLock = new();

    private IGraphLogger? _graphLogger;
    private ICommunicationsClient? _client;
    private bool _initialized;
    private bool _mediaPlatformReady;
    private string? _initError;

    public AppHostedMediaService(
        IConfiguration configuration,
        GraphAuthService graphAuthService,
        AudioHandler audioHandler,
        SpeechRecognitionService speechService,
        IBotMediaLogger mediaLogger)
    {
        _configuration = configuration;
        _graphAuthService = graphAuthService;
        _audioHandler = audioHandler;
        _speechService = speechService;
        _mediaLogger = mediaLogger;
    }

    public bool IsInitialized => _initialized && _mediaPlatformReady;

    public bool IsAppHostedCall(string? callId)
    {
        return !string.IsNullOrWhiteSpace(callId) &&
               _calls.ContainsKey(callId);
    }

    public void TryInitialize()
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
                Console.WriteLine(" MEDIA PLATFORM INITIALIZATION");
                Console.WriteLine("================================================");

                var clientId = _graphAuthService.ClientId;
                var callbackUri = _configuration["Bot:CallbackUri"]
                    ?? "https://teammate-bot.westus3.cloudapp.azure.com/api/calling";
                var serviceFqdn = _configuration["Bot:ServiceFqdn"]
                    ?? _configuration["Media:ServiceFqdn"]
                    ?? "teammate-bot.westus3.cloudapp.azure.com";

                WarnIfTunnelCallback(callbackUri);

                var internalPort = _configuration.GetValue("Media:InternalPort", 8445);
                var publicPort = _configuration.GetValue("Media:PublicPort", 8445);
                var portMin = _configuration.GetValue("Media:PortRangeMin", 20000);
                var portMax = _configuration.GetValue("Media:PortRangeMax", 20999);

                var publicIp = ResolvePublicIp(serviceFqdn);
                var privateIp = ResolvePrivateIp();
                var certificate = LoadCertificate(serviceFqdn);

                var instanceSettings = new MediaPlatformInstanceSettings
                {
                    ServiceFqdn = serviceFqdn,
                    InstancePublicIPAddress = publicIp,
                    InstanceInternalPort = internalPort,
                    InstancePublicPort = publicPort,
                    MediaPortRange = new PortRange((uint)portMin, (uint)portMax),
                    Certificate = certificate
                };

                var mediaPlatformSettings = new MediaPlatformSettings
                {
                    ApplicationId = clientId,
                    MediaPlatformInstanceSettings = instanceSettings,
                    MediaPlatformLogger = _mediaLogger
                };

                Console.WriteLine($"Service FQDN     : {serviceFqdn}");
                Console.WriteLine($"Public IP        : {publicIp}");
                Console.WriteLine($"Private IP       : {(privateIp?.ToString() ?? "(not used; SDK property not implemented in this package)")}");
                Console.WriteLine($"Control port     : {internalPort} internal / {publicPort} public");
                Console.WriteLine($"Media UDP ports  : {portMin}-{portMax}");
                Console.WriteLine($"Certificate      : {certificate.Subject}");
                Console.WriteLine($"Thumbprint       : {certificate.Thumbprint}");
                Console.WriteLine("This MediaPlatform must run on the Azure Windows VM.");
                Console.WriteLine("Teams media servers connect to the VM public IP and media ports.");
                Console.WriteLine("Dev Tunnel cannot carry application-hosted RTP/media.");

                _graphLogger = new GraphLogger("AgentTeamMateBot.AppHosted", redirectToTrace: true);

                var builder = new CommunicationsClientBuilder(
                    "AgentTeamMateBot.AppHosted",
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

                _mediaPlatformReady = true;
                _initialized = true;
                _initError = null;

                Console.WriteLine("MEDIA PLATFORM INITIALIZED");
                Console.WriteLine("================================================");
            }
            catch (Exception ex)
            {
                _initialized = false;
                _mediaPlatformReady = false;
                _initError = ex.Message;

                Console.WriteLine();
                Console.WriteLine("================================================");
                Console.WriteLine(" MEDIA PLATFORM INITIALIZATION FAILURE");
                Console.WriteLine("================================================");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex);
                Console.WriteLine();
                Console.WriteLine("Phase 2 app-hosted listening is BLOCKED until MediaPlatform starts.");
                Console.WriteLine("Required: Azure Windows Server x64 VM, VC++ x64 runtime,");
                Console.WriteLine("SSL cert for Bot:ServiceFqdn, Media:PublicIpAddress,");
                Console.WriteLine("NSG inbound TCP 8445 and UDP media port range.");
                Console.WriteLine("Service-hosted POST /api/join is unaffected.");
                Console.WriteLine("================================================");
            }
        }
    }

    public async Task<ICall> JoinMeetingAsync(string meetingId, string? passcode)
    {
        if (!_initialized || !_mediaPlatformReady || _client == null)
        {
            TryInitialize();
        }

        if (!_initialized || !_mediaPlatformReady || _client == null)
        {
            throw new InvalidOperationException(
                "App-hosted MediaPlatform is not initialized. " +
                (_initError ?? "See MEDIA PLATFORM INITIALIZATION FAILURE logs. Continuous audio is NOT proven."));
        }

        ILocalMediaSession? mediaSession = null;

        try
        {
            mediaSession = CreateLocalMediaSession();
            var audioBinding = BindAudioSocket(mediaSession);
            await _speechService.StartAsync();

            var tenantId = _graphAuthService.TenantId.Trim();
            var normalizedMeetingId = NormalizeMeetingId(meetingId);
            var normalizedPasscode = string.IsNullOrWhiteSpace(passcode) ? null : passcode.Trim();

            if (string.IsNullOrWhiteSpace(normalizedMeetingId))
            {
                throw new ArgumentException("Meeting ID is required.", nameof(meetingId));
            }

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
                    Blob = mediaSession.GetMediaConfiguration().ToString()
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
            Console.WriteLine(" APP-HOSTED JOIN REQUEST");
            Console.WriteLine("================================================");
            Console.WriteLine($"Meeting ID : {normalizedMeetingId}");
            Console.WriteLine($"Tenant ID  : {tenantId}");
            Console.WriteLine("Media      : AppHostedMediaConfig");
            Console.WriteLine("AudioSocket.AudioMediaReceived subscribed before negotiation.");

            var statefulCall = await _client.Calls().AddAsync(call, mediaSession);
            audioBinding.CallId = statefulCall.Id;
            TrackCall(statefulCall, mediaSession, audioBinding);

            Console.WriteLine($"Call ID    : {statefulCall.Id}");
            Console.WriteLine($"State      : {statefulCall.Resource?.State}");
            Console.WriteLine("================================================");

            return statefulCall;
        }
        catch (Exception ex)
        {
            mediaSession?.Dispose();

            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" APP-HOSTED JOIN FAILURE");
            Console.WriteLine("================================================");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex);
            throw;
        }
    }

    public async Task<HttpResponseMessage> ProcessNotificationAsync(HttpRequestMessage requestMessage)
    {
        if (_client == null)
        {
            return new HttpResponseMessage(HttpStatusCode.ServiceUnavailable);
        }

        try
        {
            return await _client.ProcessNotificationAsync(requestMessage);
        }
        catch (Exception ex)
        {
            Console.WriteLine("[APP-HOSTED GRAPH] Notification processing failure:");
            Console.WriteLine(ex);
            return new HttpResponseMessage(HttpStatusCode.InternalServerError);
        }
    }

    private ILocalMediaSession CreateLocalMediaSession()
    {
        if (_client == null)
        {
            throw new InvalidOperationException("App-hosted communications client is not initialized.");
        }

        var audioSocketSettings = new AudioSocketSettings
        {
            StreamDirections = StreamDirection.Sendrecv,
            SupportedAudioFormat = AudioFormat.Pcm16K,
            ReceiveUnmixedMeetingAudio = false
        };

        var mediaSession = _client.CreateMediaSession(
            audioSocketSettings,
            videoSocketSettings: (VideoSocketSettings?)null,
            vbssSocketSettings: null,
            dataSocketSettings: null);

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" LOCAL MEDIA SESSION CREATED");
        Console.WriteLine("================================================");
        Console.WriteLine($"MediaSessionId : {mediaSession.MediaSessionId}");
        Console.WriteLine("Audio format    : Pcm16K");
        Console.WriteLine("Direction       : Sendrecv");
        Console.WriteLine("================================================");

        return mediaSession;
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
            Console.WriteLine("Continuous audio is NOT proven for this call.");
            Console.WriteLine("================================================");
        };
        binding.SendStatusHandler = (_, args) =>
        {
            Console.WriteLine($"[AUDIO SOCKET] SendStatus={args.MediaSendStatus} call={binding.CallId}");
        };

        audioSocket.AudioMediaReceived += binding.AudioReceivedHandler;
        audioSocket.MediaStreamFailure += binding.MediaFailureHandler;
        audioSocket.AudioSendStatusChanged += binding.SendStatusHandler;
        _audioBindings[mediaSession.MediaSessionId] = binding;

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" AUDIO SOCKET CREATED");
        Console.WriteLine("================================================");
        Console.WriteLine("Event : IAudioSocket.AudioMediaReceived");
        Console.WriteLine("================================================");

        return binding;
    }

    private void TrackCall(ICall call, ILocalMediaSession mediaSession, AudioSocketBinding audioBinding)
    {
        audioBinding.CallId = call.Id;
        _calls[call.Id] = call;
        _mediaSessions[call.Id] = mediaSession;
        call.OnUpdated += OnCallResourceUpdated;
    }

    private void OnCallsUpdated(ICallCollection sender, CollectionEventArgs<ICall> args)
    {
        foreach (var call in args.AddedResources)
        {
            Console.WriteLine($"[APP-HOSTED CALL] Added {call.Id} state={call.Resource?.State}");
            LogCallState(call);
            if (!_calls.ContainsKey(call.Id))
            {
                call.OnUpdated += OnCallResourceUpdated;
                _calls[call.Id] = call;
            }
        }

        foreach (var call in args.UpdatedResources)
        {
            LogCallState(call);
        }

        foreach (var call in args.RemovedResources)
        {
            Console.WriteLine($"[APP-HOSTED CALL] Removed {call.Id}");
            _calls.TryRemove(call.Id, out _);
            if (_mediaSessions.TryRemove(call.Id, out var session))
            {
                try
                {
                    session.Dispose();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[APP-HOSTED] Media session dispose failed: {ex.Message}");
                }
            }
        }
    }

    private void OnCallResourceUpdated(ICall sender, ResourceEventArgs<Call> args)
    {
        LogCallState(sender);
    }

    private static void LogCallState(ICall call)
    {
        var state = call.Resource?.State;

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" CALL STATE");
        Console.WriteLine("================================================");
        Console.WriteLine($"Call ID : {call.Id}");
        Console.WriteLine($"State   : {state}");

        if (state == CallState.Establishing)
        {
            Console.WriteLine("Establishing");
        }
        else if (state == CallState.Established)
        {
            Console.WriteLine("Established");
            Console.WriteLine("Waiting for IAudioSocket.AudioMediaReceived (real Teams PCM).");
        }
        else if (state == CallState.Terminated)
        {
            Console.WriteLine("Terminated");

            var resultInfo = call.Resource?.ResultInfo;
            if (resultInfo != null)
            {
                Console.WriteLine($"resultInfo.code    : {resultInfo.Code}");
                Console.WriteLine($"resultInfo.subcode : {resultInfo.Subcode}");
                Console.WriteLine($"resultInfo.message : {resultInfo.Message}");

                if (resultInfo.Subcode == 1203002)
                {
                    Console.WriteLine();
                    Console.WriteLine("BLOCKED: Graph media negotiation failed (1203002).");
                    Console.WriteLine("Continuous audio is NOT proven.");
                    Console.WriteLine("Check VM public IP, TCP 8445, UDP media ports,");
                    Console.WriteLine("and that ServiceFqdn matches the media certificate.");
                    Console.WriteLine("Do not use Dev Tunnel for application-hosted media.");
                }
            }
        }

        Console.WriteLine("================================================");
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
                callId,
                buffer.Timestamp);
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
            return IPAddress.Parse(configured.Trim());
        }

        try
        {
            var addresses = Dns.GetHostAddresses(serviceFqdn);
            var ipv4 = Array.Find(addresses, a => a.AddressFamily == AddressFamily.InterNetwork);
            if (ipv4 != null && !IPAddress.IsLoopback(ipv4))
            {
                return ipv4;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[MEDIA] DNS lookup for {serviceFqdn} failed: {ex.Message}");
        }

        throw new InvalidOperationException(
            "Media:PublicIpAddress is required for application-hosted media. " +
            "Set it to the Azure VM public IPv4. Dev Tunnel IPs cannot receive Teams RTP.");
    }

    private IPAddress? ResolvePrivateIp()
    {
        var configured = _configuration["Media:PrivateIpAddress"];
        if (string.IsNullOrWhiteSpace(configured))
        {
            return null;
        }

        return IPAddress.Parse(configured.Trim());
    }

    private X509Certificate2 LoadCertificate(string host)
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

        throw new InvalidOperationException(
            $"No SSL certificate with private key found for {host}. " +
            "Set Media:CertificateThumbprint or Media:CertificatePath.");
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

            if (extension.Format(true).Contains(host, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static void WarnIfTunnelCallback(string callbackUri)
    {
        if (callbackUri.Contains("devtunnels.ms", StringComparison.OrdinalIgnoreCase) ||
            callbackUri.Contains("ngrok", StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine();
            Console.WriteLine("WARNING: Bot:CallbackUri looks like a Dev Tunnel.");
            Console.WriteLine("Graph signaling may work through a tunnel, but Teams media");
            Console.WriteLine("will still target Media:PublicIpAddress / Bot:ServiceFqdn.");
            Console.WriteLine("For Phase 2, run on the Azure VM and prefer:");
            Console.WriteLine("https://teammate-bot.westus3.cloudapp.azure.com/api/calling");
        }
    }

    private static string NormalizeMeetingId(string meetingId)
    {
        if (string.IsNullOrWhiteSpace(meetingId))
        {
            return string.Empty;
        }

        return new string(meetingId.Where(char.IsDigit).ToArray());
    }

    private sealed class AudioSocketBinding
    {
        public string CallId { get; set; } = string.Empty;

        public EventHandler<AudioMediaReceivedEventArgs>? AudioReceivedHandler { get; set; }

        public EventHandler<MediaStreamFailureEventArgs>? MediaFailureHandler { get; set; }

        public EventHandler<AudioSendStatusChangedEventArgs>? SendStatusHandler { get; set; }
    }
}

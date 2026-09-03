using System.Collections.Concurrent;
using System.Net;
using System.Net.NetworkInformation;
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
    private readonly AiResponseService _aiResponseService;
    private readonly SpeechSynthesisService _speechSynthesisService;
    private readonly MeetingContextService _meetingContextService;
    private readonly IBotMediaLogger _mediaLogger;

    private readonly ConcurrentDictionary<string, ICall> _calls = new();
    private readonly ConcurrentDictionary<string, ILocalMediaSession> _mediaSessions = new();
    private readonly ConcurrentDictionary<Guid, AudioSocketBinding> _audioBindings = new();
    private readonly ConcurrentDictionary<string, byte> _welcomePlayed = new();
    private readonly object _initLock = new();

    // Track the active call ID for the continuous recognizer callback
    private string? _activeCallId;

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
        AiResponseService aiResponseService,
        SpeechSynthesisService speechSynthesisService,
        MeetingContextService meetingContextService,
        IBotMediaLogger mediaLogger)
    {
        _configuration = configuration;
        _graphAuthService = graphAuthService;
        _audioHandler = audioHandler;
        _speechService = speechService;
        _aiResponseService = aiResponseService;
        _speechSynthesisService = speechSynthesisService;
        _meetingContextService = meetingContextService;
        _mediaLogger = mediaLogger;

        // Subscribe to continuous speech recognition events
        _speechService.OnSpeechRecognized += OnLiveSpeechRecognized;
    }

    public bool IsInitialized => _initialized && _mediaPlatformReady;

    public string? InitError => _initError;

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
                BotLog.Info("Initializing MediaPlatform...");

                var clientId = _graphAuthService.ClientId;
                var callbackUri = _configuration["Bot:CallbackUri"]
                    ?? "https://teammate-bot.westus3.cloudapp.azure.com/api/calling";
                var serviceFqdn = _configuration["Bot:ServiceFqdn"]
                    ?? _configuration["Media:ServiceFqdn"]
                    ?? "teammate-bot.westus3.cloudapp.azure.com";

                WarnIfTunnelCallback(callbackUri);
                EnsureNativeMediaPresent();

                var internalPort = _configuration.GetValue("Media:InternalPort", 8445);
                var publicPort = _configuration.GetValue("Media:PublicPort", 8445);
                var portMin = _configuration.GetValue("Media:PortRangeMin", 20000);
                var portMax = _configuration.GetValue("Media:PortRangeMax", 20999);

                var publicIp = ResolvePublicIp(serviceFqdn);
                var privateIp = ResolvePrivateIp();
                var certificate = LoadCertificate(serviceFqdn);

                if (!certificate.HasPrivateKey)
                {
                    throw new InvalidOperationException(
                        $"Certificate {certificate.Thumbprint} has no private key accessible to this process. " +
                        "Import the PFX into LocalMachine\\My and grant the bot user Read on the private key.");
                }

                // EchoBot / Media SDK loads the cert from LocalMachine by thumbprint.
                var instanceSettings = new MediaPlatformInstanceSettings
                {
                    ServiceFqdn = serviceFqdn,
                    CertificateThumbprint = certificate.Thumbprint,
                    InstancePublicIPAddress = publicIp,
                    InstanceInternalPort = internalPort,
                    InstancePublicPort = publicPort,
                    MediaPortRange = new PortRange((uint)portMin, (uint)portMax)
                };

                var mediaPlatformSettings = new MediaPlatformSettings
                {
                    ApplicationId = clientId,
                    MediaPlatformInstanceSettings = instanceSettings,
                    MediaPlatformLogger = _mediaLogger
                };

                BotLog.Info(
                    $"MediaPlatform FQDN={serviceFqdn} PublicIP={publicIp} PrivateIP={privateIp?.ToString() ?? "(detected-none)"} Cert={certificate.Thumbprint}");
                Console.WriteLine($"Service FQDN     : {serviceFqdn}");
                Console.WriteLine($"Public IP        : {publicIp}");
                Console.WriteLine($"Private IP       : {(privateIp?.ToString() ?? "(not required by this SDK package)")}");
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

                BotLog.Info("MediaPlatform ready.");
                Console.WriteLine("MEDIA PLATFORM INITIALIZED");
                Console.WriteLine("================================================");
            }
            catch (Exception ex)
            {
                _initialized = false;
                _mediaPlatformReady = false;
                _initError = FormatExceptionChain(ex);

                BotLog.Info($"MediaPlatform FAILED: {_initError}");
                Console.WriteLine();
                Console.WriteLine("================================================");
                Console.WriteLine(" MEDIA PLATFORM INITIALIZATION FAILURE");
                Console.WriteLine("================================================");
                Console.WriteLine(_initError);
                Console.WriteLine(ex);
                Console.WriteLine();
                Console.WriteLine("Phase 2 app-hosted listening is BLOCKED until MediaPlatform starts.");
                Console.WriteLine("Required: full Windows Server (not Core), VC++ x64 runtime,");
                Console.WriteLine("Windows Media Foundation, SSL cert in LocalMachine\\My,");
                Console.WriteLine("Media:PublicIpAddress + Media:PrivateIpAddress,");
                Console.WriteLine("NSG inbound TCP 443/8445 and UDP 20000-20999.");
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
                "MediaPlatform was not initialized. " +
                (_initError ?? "Missing Media:PublicIpAddress, SSL cert for Bot:ServiceFqdn, or VC++ x64 runtime."));
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
            _activeCallId = statefulCall.Id;
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

            if (args.MediaSendStatus == MediaSendStatus.Active &&
                !string.IsNullOrWhiteSpace(binding.CallId))
            {
                var callId = binding.CallId;
                _ = Task.Run(async () =>
                {
                    try
                    {
                        await WaitForCallEstablishedAsync(callId);
                        await PlayWelcomeAsync(callId);
                    }
                    catch (Exception ex)
                    {
                        BotLog.Info($"Error: Welcome playback failed. {ex.Message}");
                    }
                });
            }
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
            _welcomePlayed.TryRemove(call.Id, out _);
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

    // ============================================================
    // LIVE SPEECH → AGENT NOVA → AI → TTS → AUDIOSOCKET
    // ============================================================

    private void OnLiveSpeechRecognized(string recognizedText)
    {
        var callId = _activeCallId;
        if (string.IsNullOrWhiteSpace(callId))
        {
            return;
        }

        _meetingContextService.AppendLiveTranscript(
            callId, recognizedText);

        if (!WakeWordDetector.IsAgentInvocation(recognizedText))
        {
            return;
        }

        var question =
            WakeWordDetector.RemoveActivationPhrase(recognizedText);

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" [APP-HOSTED] AGENT NOVA INVOCATION DETECTED");
        Console.WriteLine("================================================");
        Console.WriteLine($"Speech   : {recognizedText}");
        Console.WriteLine($"Question : {question}");
        Console.WriteLine("================================================");

        BotLog.Info($"User: {recognizedText}");
        BotLog.Info("Processing...");

        _ = Task.Run(async () =>
        {
            try
            {
                var aiResponse =
                    await _aiResponseService
                        .GetResponseAsync(callId, question);

                if (string.IsNullOrWhiteSpace(aiResponse))
                {
                    BotLog.Info("Error: No AI response text.");
                    Console.WriteLine("[APP-HOSTED AI] No response received.");
                    return;
                }

                Console.WriteLine();
                Console.WriteLine("================================================");
                Console.WriteLine(" [APP-HOSTED] AGENT RESPONSE");
                Console.WriteLine("================================================");
                Console.WriteLine(aiResponse);
                Console.WriteLine("================================================");

                BotLog.Info($"Nova: {aiResponse}");

                var pcmAudio =
                    await SynthesizeSpeechToPcmAsync(aiResponse);

                if (pcmAudio == null || pcmAudio.Length == 0)
                {
                    BotLog.Info("Error: Could not generate voice audio.");
                    Console.WriteLine("[APP-HOSTED TTS] No audio generated.");
                    return;
                }

                await SendPcmToAudioSocketAsync(callId, pcmAudio);
            }
            catch (Exception ex)
            {
                BotLog.Info($"Error: {ex.Message}");
                Console.WriteLine(
                    $"[APP-HOSTED] AI/TTS pipeline error: {ex.Message}");
            }
        });
    }

    private async Task PlayWelcomeAsync(string callId)
    {
        if (!_welcomePlayed.TryAdd(callId, 0))
        {
            return;
        }

        try
        {
            BotLog.Info("Playing welcome...");

            var pcmAudio = await SynthesizeSpeechToPcmAsync(
                "Hi, I am Agent Nova. I am listening. Say your request.");

            if (pcmAudio == null || pcmAudio.Length == 0)
            {
                BotLog.Info("Error: Could not generate welcome audio.");
                return;
            }

            await SendPcmToAudioSocketAsync(callId, pcmAudio);
        }
        catch (Exception ex)
        {
            BotLog.Info($"Error: Welcome playback failed. {ex.Message}");
            Console.WriteLine($"[APP-HOSTED WELCOME] {ex.Message}");
        }
    }

    private async Task WaitForCallEstablishedAsync(string callId)
    {
        for (var i = 0; i < 75; i++)
        {
            if (_calls.TryGetValue(callId, out var call) &&
                call.Resource?.State == CallState.Established)
            {
                return;
            }

            await Task.Delay(200);
        }

        Console.WriteLine(
            $"[APP-HOSTED WELCOME] Timed out waiting for Established on {callId}. Playing anyway.");
    }

    // ============================================================
    // SYNTHESIZE SPEECH TO RAW PCM (16kHz 16-bit mono)
    // ============================================================

    private async Task<byte[]?> SynthesizeSpeechToPcmAsync(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        var key = _configuration["Speech:Key"]
            ?? _configuration["AZURE_SPEECH_KEY"];

        var region = _configuration["Speech:Region"]
            ?? _configuration["AZURE_SPEECH_REGION"];

        var voice = _configuration["Speech:Voice"]
            ?? "en-US-AvaMultilingualNeural";

        if (string.IsNullOrWhiteSpace(key) ||
            string.IsNullOrWhiteSpace(region))
        {
            Console.WriteLine("[APP-HOSTED TTS] Speech key or region missing.");
            return null;
        }

        var speechConfig =
            Microsoft.CognitiveServices.Speech.SpeechConfig
                .FromSubscription(key, region);

        speechConfig.SpeechSynthesisVoiceName = voice;
        speechConfig.SetSpeechSynthesisOutputFormat(
            Microsoft.CognitiveServices.Speech.SpeechSynthesisOutputFormat
                .Raw16Khz16BitMonoPcm);

        using var synthesizer =
            new Microsoft.CognitiveServices.Speech.SpeechSynthesizer(
                speechConfig, null);

        var result = await synthesizer.SpeakTextAsync(text);

        if (result.Reason ==
            Microsoft.CognitiveServices.Speech.ResultReason
                .SynthesizingAudioCompleted)
        {
            Console.WriteLine(
                $"[APP-HOSTED TTS] Synthesized {result.AudioData.Length} bytes of raw PCM.");

            return result.AudioData;
        }

        Console.WriteLine(
            $"[APP-HOSTED TTS] Synthesis failed: {result.Reason}");

        return null;
    }

    // ============================================================
    // SEND RAW PCM FRAMES VIA AUDIOSOCKET
    // Media SDK keeps using each buffer after Send returns. Do not
    // free the unmanaged memory here — AudioSendBuffer.Dispose does.
    // ============================================================

    private Task SendPcmToAudioSocketAsync(string callId, byte[] pcmData)
    {
        return Task.Run(() => SendPcmToAudioSocket(callId, pcmData));
    }

    private void SendPcmToAudioSocket(
        string callId, byte[] pcmData)
    {
        if (!_mediaSessions.TryGetValue(callId, out var mediaSession))
        {
            Console.WriteLine(
                $"[APP-HOSTED SEND] No media session for call {callId}.");
            return;
        }

        var audioSocket = mediaSession.AudioSocket;
        if (audioSocket == null)
        {
            Console.WriteLine(
                "[APP-HOSTED SEND] AudioSocket is null.");
            return;
        }

        // PCM 16kHz 16-bit mono: 20ms frames = 640 bytes
        const int frameSize = 640;
        const long frameDurationTicks = 20 * 10000;
        var timestamp = DateTime.UtcNow.Ticks;
        var totalFrames = pcmData.Length / frameSize;

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" [APP-HOSTED] SENDING AUDIO TO TEAMS");
        Console.WriteLine("================================================");
        Console.WriteLine($"Call ID      : {callId}");
        Console.WriteLine($"PCM bytes    : {pcmData.Length}");
        Console.WriteLine($"Frames (20ms): {totalFrames}");
        Console.WriteLine("================================================");

        var stopwatch = System.Diagnostics.Stopwatch.StartNew();

        for (var i = 0; i < totalFrames; i++)
        {
            var unmanagedBuffer = Marshal.AllocHGlobal(frameSize);
            try
            {
                Marshal.Copy(pcmData, i * frameSize, unmanagedBuffer, frameSize);

                // Ownership of unmanagedBuffer transfers to AudioSendBuffer.
                // Media platform calls Dispose later and frees the memory.
                var buffer = new AudioSendBuffer(
                    unmanagedBuffer,
                    frameSize,
                    AudioFormat.Pcm16K,
                    timestamp);

                unmanagedBuffer = IntPtr.Zero;
                audioSocket.Send(buffer);
            }
            catch (Exception ex)
            {
                if (unmanagedBuffer != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(unmanagedBuffer);
                }

                Console.WriteLine(
                    $"[APP-HOSTED SEND] Frame {i} send error: {ex.Message}");
                break;
            }

            timestamp += frameDurationTicks;

            var targetMs = (i + 1) * 20;
            var delayMs = targetMs - (int)stopwatch.ElapsedMilliseconds;
            if (delayMs > 0)
            {
                Thread.Sleep(delayMs);
            }
        }

        Console.WriteLine(
            $"[APP-HOSTED SEND] Sent audio for call {callId}.");
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
        if (!string.IsNullOrWhiteSpace(configured))
        {
            return IPAddress.Parse(configured.Trim());
        }

        try
        {
            foreach (var nic in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (nic.OperationalStatus != OperationalStatus.Up)
                {
                    continue;
                }

                if (nic.NetworkInterfaceType is NetworkInterfaceType.Loopback or NetworkInterfaceType.Tunnel)
                {
                    continue;
                }

                foreach (var address in nic.GetIPProperties().UnicastAddresses)
                {
                    if (address.Address.AddressFamily != AddressFamily.InterNetwork)
                    {
                        continue;
                    }

                    if (IPAddress.IsLoopback(address.Address))
                    {
                        continue;
                    }

                    // Prefer Azure VNet private ranges.
                    var bytes = address.Address.GetAddressBytes();
                    var isPrivate =
                        bytes[0] == 10 ||
                        (bytes[0] == 172 && bytes[1] >= 16 && bytes[1] <= 31) ||
                        (bytes[0] == 192 && bytes[1] == 168);

                    if (isPrivate)
                    {
                        return address.Address;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[MEDIA] Private IP auto-detect failed: {ex.Message}");
        }

        return null;
    }

    private static void EnsureNativeMediaPresent()
    {
        var nativeMedia = Path.Combine(AppContext.BaseDirectory, "NativeMedia.dll");
        if (!File.Exists(nativeMedia))
        {
            throw new InvalidOperationException(
                $"NativeMedia.dll was not found next to the app at {AppContext.BaseDirectory}. " +
                "Rebuild so Microsoft.Skype.Bots.Media native binaries are copied to the output.");
        }

        var required = new[]
        {
            "RtmPal.dll",
            "RtmCodecs.dll",
            "skypert.dll",
            "Ijwhost.dll"
        };

        var missing = required
            .Where(name => !File.Exists(Path.Combine(AppContext.BaseDirectory, name)))
            .ToArray();

        if (missing.Length > 0)
        {
            throw new InvalidOperationException(
                "Media native dependencies missing from output: " +
                string.Join(", ", missing) +
                $". Output folder: {AppContext.BaseDirectory}");
        }
    }

    private static string FormatExceptionChain(Exception ex)
    {
        var parts = new List<string>();
        for (var current = ex; current != null; current = current.InnerException)
        {
            parts.Add($"{current.GetType().Name}: {current.Message}");
        }

        return string.Join(" => ", parts);
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

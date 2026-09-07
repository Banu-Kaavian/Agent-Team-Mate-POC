using Microsoft.CognitiveServices.Speech;
using Microsoft.CognitiveServices.Speech.Audio;

namespace AgentTeamMateBot.Services;

public class SpeechRecognitionService
{
    private readonly IConfiguration _configuration;

    public const int SpeechSampleRate = 16000;
    public const int SpeechBitsPerSample = 16;
    public const int SpeechChannels = 1;

    // Quiet-meeting idle before Azure Speech drops the websocket.
    private const string LiveIdleTimeoutMs = "300000";
    // Azure max for end-of-phrase silence (was 1500). Longer pauses stay one point.
    private const string LiveEndSilenceMs = "5000";
    // Default phrase cap is ~20s; +20s so one spoken point can run to 40s.
    private const string LiveMaxPhraseMs = "40000";

    public SpeechRecognitionService(
        IConfiguration configuration)
    {
        _configuration =
            configuration;
    }

    private readonly SemaphoreSlim _liveGate = new(1, 1);
    private SpeechRecognizer? _liveRecognizer;
    private PushAudioInputStream? _livePushStream;
    private AudioConfig? _liveAudioConfig;
    private volatile bool _wantLive;
    private volatile bool _acceptingAudio;
    private int _restartQueued;
    private int _restartFailures;

    public event Action<string>? OnSpeechRecognized;

    public bool IsListening => _acceptingAudio;

    // ============================================================
    // PHASE 2: CONTINUOUS LIVE PCM FROM AUDIOSOCKET
    // RecognizeRecordingAsync below is unchanged for service-hosted.
    // ============================================================

    public async Task StartAsync()
    {
        _wantLive = true;
        await _liveGate.WaitAsync();
        try
        {
            if (_liveRecognizer != null && _acceptingAudio)
            {
                return;
            }

            if (_liveRecognizer != null)
            {
                await TearDownLiveAsync();
            }

            await CreateAndStartLiveAsync();
        }
        catch (Exception ex)
        {
            _acceptingAudio = false;
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" LIVE SPEECH SDK FAILURE");
            Console.WriteLine("================================================");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex);
            throw;
        }
        finally
        {
            _liveGate.Release();
        }
    }

    public void ProcessAudio(
        byte[] audioData)
    {
        if (!_acceptingAudio ||
            _livePushStream == null ||
            audioData == null ||
            audioData.Length == 0)
        {
            return;
        }

        try
        {
            _livePushStream.Write(audioData);
        }
        catch (Exception ex)
        {
            _acceptingAudio = false;
            Console.WriteLine(
                $"[LIVE SPEECH] Push stream write failed: {ex.Message}");
            QueueLiveRestart();
        }
    }

    public async Task StopAsync()
    {
        _wantLive = false;
        _acceptingAudio = false;
        await _liveGate.WaitAsync();
        try
        {
            await TearDownLiveAsync();
        }
        finally
        {
            _liveGate.Release();
        }
    }

    private async Task CreateAndStartLiveAsync()
    {
        var key =
            _configuration["Speech:Key"]
            ?? _configuration["AZURE_SPEECH_KEY"];

        var region =
            _configuration["Speech:Region"]
            ?? _configuration["AZURE_SPEECH_REGION"];

        if (string.IsNullOrWhiteSpace(key))
        {
            throw new Exception("Speech:Key missing");
        }

        if (string.IsNullOrWhiteSpace(region))
        {
            throw new Exception("Speech:Region missing");
        }

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" LIVE AZURE SPEECH STARTING");
        Console.WriteLine("================================================");
        Console.WriteLine($"Speech region : {region}");
        Console.WriteLine($"Idle timeout : {int.Parse(LiveIdleTimeoutMs) / 1000} seconds");
        Console.WriteLine($"End silence  : {int.Parse(LiveEndSilenceMs)} ms");
        Console.WriteLine($"Max phrase   : {int.Parse(LiveMaxPhraseMs) / 1000} seconds");
        Console.WriteLine("================================================");

        var speechConfig = CreateLiveRecognitionConfig(key, region);
        var format = AudioStreamFormat.GetWaveFormatPCM(
            SpeechSampleRate,
            SpeechBitsPerSample,
            SpeechChannels);

        _livePushStream = AudioInputStream.CreatePushStream(format);
        _liveAudioConfig = AudioConfig.FromStreamInput(_livePushStream);
        var recognizer = new SpeechRecognizer(speechConfig, _liveAudioConfig);
        _liveRecognizer = recognizer;

        recognizer.SessionStarted += (_, _) =>
        {
            if (!ReferenceEquals(_liveRecognizer, recognizer))
            {
                return;
            }

            _acceptingAudio = true;
            _restartFailures = 0;
            Console.WriteLine("Live Azure Speech session started.");
        };

        recognizer.SessionStopped += (_, _) =>
        {
            if (!ReferenceEquals(_liveRecognizer, recognizer))
            {
                return;
            }

            _acceptingAudio = false;
            Console.WriteLine("Live Azure Speech session stopped.");
            QueueLiveRestart();
        };

        recognizer.Recognized += (_, e) =>
        {
            if (!ReferenceEquals(_liveRecognizer, recognizer))
            {
                return;
            }

            if (e.Result.Reason != ResultReason.RecognizedSpeech ||
                string.IsNullOrWhiteSpace(e.Result.Text))
            {
                return;
            }

            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" LIVE SPEECH");
            Console.WriteLine("================================================");
            Console.WriteLine(e.Result.Text);
            Console.WriteLine("================================================");

            try
            {
                OnSpeechRecognized?.Invoke(e.Result.Text);
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    $"[LIVE SPEECH] Callback error: {ex.Message}");
            }
        };

        recognizer.Canceled += (_, e) =>
        {
            if (!ReferenceEquals(_liveRecognizer, recognizer))
            {
                return;
            }

            _acceptingAudio = false;
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" LIVE SPEECH SDK FAILURE");
            Console.WriteLine("================================================");
            Console.WriteLine($"Reason        : {e.Reason}");
            Console.WriteLine($"Error code    : {e.ErrorCode}");
            Console.WriteLine($"Error details : {e.ErrorDetails}");
            Console.WriteLine("================================================");

            if (e.Reason == CancellationReason.Error)
            {
                QueueLiveRestart();
            }
        };

        await recognizer.StartContinuousRecognitionAsync();
        if (!ReferenceEquals(_liveRecognizer, recognizer))
        {
            return;
        }
        _acceptingAudio = true;
        Console.WriteLine(
            "Live Azure Speech continuous recognition started.");
    }

    private void QueueLiveRestart()
    {
        if (!_wantLive)
        {
            return;
        }

        if (Interlocked.Exchange(ref _restartQueued, 1) == 1)
        {
            return;
        }

        _ = Task.Run(async () =>
        {
            try
            {
                var delayMs = Math.Min(
                    8000,
                    500 * (1 << Math.Min(_restartFailures, 4)));
                await Task.Delay(delayMs);
                if (!_wantLive)
                {
                    return;
                }

                await _liveGate.WaitAsync();
                try
                {
                    if (!_wantLive)
                    {
                        return;
                    }

                    Console.WriteLine(
                        "[LIVE SPEECH] Restarting continuous recognition after timeout.");
                    await TearDownLiveAsync();
                    await CreateAndStartLiveAsync();
                }
                catch (Exception ex)
                {
                    _restartFailures++;
                    _acceptingAudio = false;
                    Console.WriteLine(
                        $"[LIVE SPEECH] Restart failed: {ex.Message}");
                    try
                    {
                        await TearDownLiveAsync();
                    }
                    catch
                    {
                        // Already logged in TearDownLiveAsync.
                    }
                }
                finally
                {
                    _liveGate.Release();
                }
            }
            finally
            {
                Interlocked.Exchange(ref _restartQueued, 0);
                if (_wantLive && !_acceptingAudio)
                {
                    QueueLiveRestart();
                }
            }
        });
    }

    private async Task TearDownLiveAsync()
    {
        _acceptingAudio = false;
        var recognizer = _liveRecognizer;
        var pushStream = _livePushStream;
        var audioConfig = _liveAudioConfig;
        _liveRecognizer = null;
        _livePushStream = null;
        _liveAudioConfig = null;

        if (recognizer != null)
        {
            try
            {
                await recognizer.StopContinuousRecognitionAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    $"[LIVE SPEECH] Stop failed: {ex.Message}");
            }

            recognizer.Dispose();
        }

        try
        {
            pushStream?.Close();
        }
        catch
        {
            // Ignore close after a dead session.
        }

        pushStream?.Dispose();
        audioConfig?.Dispose();
    }

    private static SpeechConfig CreateLiveRecognitionConfig(
        string key,
        string region)
    {
        var speechConfig = SpeechConfig.FromSubscription(key, region);
        speechConfig.SpeechRecognitionLanguage = "en-US";
        speechConfig.SetProperty(
            PropertyId.SpeechServiceConnection_InitialSilenceTimeoutMs,
            LiveIdleTimeoutMs);
        speechConfig.SetProperty(
            PropertyId.SpeechServiceConnection_EndSilenceTimeoutMs,
            LiveEndSilenceMs);
        speechConfig.SetProperty(
            PropertyId.Speech_SegmentationStrategy,
            "Time");
        speechConfig.SetProperty(
            PropertyId.Speech_SegmentationSilenceTimeoutMs,
            LiveEndSilenceMs);
        speechConfig.SetProperty(
            PropertyId.Speech_SegmentationMaximumTimeMs,
            LiveMaxPhraseMs);
        return speechConfig;
    }

    // ============================================================
    // RECOGNIZE recordResponse AUDIO
    // ============================================================

    public async Task<string?> RecognizeRecordingAsync(
        byte[] recordingBytes)
    {
        if (recordingBytes == null ||
            recordingBytes.Length == 0)
        {
            Console.WriteLine(
                "[SPEECH] Recording is empty.");

            return null;
        }

        var key =
            _configuration["Speech:Key"]
            ?? _configuration["AZURE_SPEECH_KEY"];

        var region =
            _configuration["Speech:Region"]
            ?? _configuration["AZURE_SPEECH_REGION"];

        if (string.IsNullOrWhiteSpace(
                key))
        {
            throw new Exception(
                "Speech:Key missing");
        }

        if (string.IsNullOrWhiteSpace(
                region))
        {
            throw new Exception(
                "Speech:Region missing");
        }

        var tempFile =
            Path.Combine(
                Path.GetTempPath(),
                $"agent-teammate-{Guid.NewGuid():N}.wav");

        try
        {
            await File.WriteAllBytesAsync(
                tempFile,
                recordingBytes);

            TrySaveDebugRecording(
                recordingBytes);

            var wav =
                AnalyzeWav(recordingBytes);

            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" AZURE SPEECH RECOGNITION");
            Console.WriteLine("================================================");

            Console.WriteLine(
                $"Audio bytes   : {recordingBytes.Length}");

            Console.WriteLine(
                $"Speech region : {region}");

            Console.WriteLine(
                $"WAV format    : {wav.SampleRate} Hz / {wav.BitsPerSample}-bit / {wav.Channels} ch");

            Console.WriteLine(
                $"PCM energy    : rms={wav.Rms:F1} peak={wav.Peak}");

            if (wav.Rms < 50)
            {
                Console.WriteLine(
                    "WARNING: Recording is nearly silent. Unmute your Teams mic,");
                Console.WriteLine(
                    "speak closer to the microphone, and say 'Agent Nova' clearly.");
            }

            var speechConfig =
                CreateRecognitionConfig(
                    key,
                    region);

            using var audioConfig =
                AudioConfig.FromWavFileInput(
                    tempFile);

            using var recognizer =
                new SpeechRecognizer(
                    speechConfig,
                    audioConfig);

            Console.WriteLine(
                "Recognizing speech...");

            var result =
                await recognizer
                    .RecognizeOnceAsync();

            if (result.Reason == ResultReason.NoMatch &&
                wav.Rms >= 50)
            {
                Console.WriteLine(
                    "[SPEECH] WAV recognize returned NoMatch. Retrying raw PCM push stream...");

                var retry =
                    await RecognizePcmPushAsync(
                        speechConfig,
                        wav);

                if (!string.IsNullOrWhiteSpace(retry))
                {
                    return retry;
                }
            }

            // ============================================================
            // SUCCESS
            // ============================================================

            if (result.Reason ==
                ResultReason.RecognizedSpeech)
            {
                Console.WriteLine();
                Console.WriteLine("================================================");
                Console.WriteLine(" SPEECH RECOGNIZED");
                Console.WriteLine("================================================");

                Console.WriteLine(
                    result.Text);

                Console.WriteLine(
                    "================================================");

                return result.Text;
            }

            // ============================================================
            // NO MATCH
            // ============================================================

            if (result.Reason ==
                ResultReason.NoMatch)
            {
                Console.WriteLine();
                Console.WriteLine("================================================");
                Console.WriteLine(" SPEECH NOT RECOGNIZED");
                Console.WriteLine("================================================");

                var noMatchDetails =
                    NoMatchDetails.FromResult(
                        result);

                Console.WriteLine(
                    $"Reason : {noMatchDetails.Reason}");

                Console.WriteLine(
                    "================================================");

                return null;
            }

            // ============================================================
            // CANCELLED
            // ============================================================

            if (result.Reason ==
                ResultReason.Canceled)
            {
                var cancellation =
                    CancellationDetails.FromResult(
                        result);

                Console.WriteLine();
                Console.WriteLine("================================================");
                Console.WriteLine(" AZURE SPEECH FAILURE");
                Console.WriteLine("================================================");

                Console.WriteLine(
                    $"Reason  : {cancellation.Reason}");

                Console.WriteLine(
                    $"Code    : {cancellation.ErrorCode}");

                Console.WriteLine(
                    $"Details : {cancellation.ErrorDetails}");

                Console.WriteLine(
                    "================================================");

                return null;
            }

            Console.WriteLine(
                $"[SPEECH] Unexpected result: {result.Reason}");

            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" SPEECH SDK FAILURE");
            Console.WriteLine("================================================");

            Console.WriteLine(
                ex.Message);

            Console.WriteLine(
                ex);

            throw;
        }
        finally
        {
            try
            {
                if (File.Exists(
                        tempFile))
                {
                    File.Delete(
                        tempFile);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    $"[SPEECH] Could not delete temp file: {ex.Message}");
            }
        }
    }

    private static SpeechConfig CreateRecognitionConfig(
        string key,
        string region)
    {
        var speechConfig =
            SpeechConfig.FromSubscription(
                key,
                region);

        speechConfig.SpeechRecognitionLanguage =
            "en-US";

        speechConfig.SetProperty(
            PropertyId.SpeechServiceConnection_InitialSilenceTimeoutMs,
            "15000");

        speechConfig.SetProperty(
            PropertyId.SpeechServiceConnection_EndSilenceTimeoutMs,
            LiveEndSilenceMs);

        return speechConfig;
    }

    private static async Task<string?> RecognizePcmPushAsync(
        SpeechConfig speechConfig,
        WavAnalysis wav)
    {
        if (wav.Pcm.Length == 0)
        {
            return null;
        }

        var format =
            AudioStreamFormat.GetWaveFormatPCM(
                (uint)wav.SampleRate,
                (byte)wav.BitsPerSample,
                (byte)wav.Channels);

        using var pushStream =
            AudioInputStream.CreatePushStream(
                format);

        using var audioConfig =
            AudioConfig.FromStreamInput(
                pushStream);

        using var recognizer =
            new SpeechRecognizer(
                speechConfig,
                audioConfig);

        var recognizeTask =
            recognizer.RecognizeOnceAsync();

        pushStream.Write(wav.Pcm);
        pushStream.Close();

        var result =
            await recognizeTask;

        if (result.Reason == ResultReason.RecognizedSpeech &&
            !string.IsNullOrWhiteSpace(result.Text))
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" SPEECH RECOGNIZED");
            Console.WriteLine("================================================");
            Console.WriteLine(result.Text);
            Console.WriteLine("================================================");
            return result.Text;
        }

        Console.WriteLine(
            $"[SPEECH] PCM retry result: {result.Reason}");

        return null;
    }

    private void TrySaveDebugRecording(
        byte[] recordingBytes)
    {
        try
        {
            var debugDirectory =
                Path.Combine(
                    AppContext.BaseDirectory,
                    "TempAudio");

            Directory.CreateDirectory(debugDirectory);

            File.WriteAllBytes(
                Path.Combine(debugDirectory, "last-recording.wav"),
                recordingBytes);
        }
        catch (Exception ex)
        {
            Console.WriteLine(
                $"[SPEECH] Could not save debug recording: {ex.Message}");
        }
    }

    private readonly record struct WavAnalysis(
        int SampleRate,
        int BitsPerSample,
        int Channels,
        byte[] Pcm,
        double Rms,
        int Peak);

    private static WavAnalysis AnalyzeWav(
        byte[] bytes)
    {
        if (bytes.Length < 44 ||
            bytes[0] != (byte)'R' ||
            bytes[1] != (byte)'I')
        {
            return MeasurePcm(
                bytes,
                16000,
                16,
                1,
                0);
        }

        var channels =
            BitConverter.ToInt16(bytes, 22);

        var sampleRate =
            BitConverter.ToInt32(bytes, 24);

        var bits =
            BitConverter.ToInt16(bytes, 34);

        var dataOffset = 44;
        for (var i = 12; i < bytes.Length - 8; i++)
        {
            if (bytes[i] == (byte)'d' &&
                bytes[i + 1] == (byte)'a' &&
                bytes[i + 2] == (byte)'t' &&
                bytes[i + 3] == (byte)'a')
            {
                dataOffset = i + 8;
                break;
            }
        }

        return MeasurePcm(
            bytes,
            sampleRate > 0 ? sampleRate : 16000,
            bits > 0 ? bits : 16,
            channels > 0 ? channels : 1,
            dataOffset);
    }

    private static WavAnalysis MeasurePcm(
        byte[] bytes,
        int sampleRate,
        int bits,
        int channels,
        int offset)
    {
        if (offset >= bytes.Length)
        {
            return new WavAnalysis(
                sampleRate,
                bits,
                channels,
                Array.Empty<byte>(),
                0,
                0);
        }

        var pcm =
            bytes[offset..];

        long sumSquares = 0;
        var peak = 0;
        var samples = 0;

        if (bits == 16)
        {
            for (var i = 0; i + 1 < pcm.Length; i += 2)
            {
                var sample =
                    Math.Abs(
                        BitConverter.ToInt16(pcm, i));

                sumSquares += (long)sample * sample;
                if (sample > peak)
                {
                    peak = sample;
                }

                samples++;
            }
        }

        var rms =
            samples == 0
                ? 0
                : Math.Sqrt(sumSquares / (double)samples);

        return new WavAnalysis(
            sampleRate,
            bits,
            channels,
            pcm,
            rms,
            peak);
    }
}
using Microsoft.CognitiveServices.Speech;
using Microsoft.CognitiveServices.Speech.Audio;

namespace AgentTeamMateBot.Services;

public class SpeechRecognitionService
{
    private readonly IConfiguration _configuration;

    public const int SpeechSampleRate = 16000;
    public const int SpeechBitsPerSample = 16;
    public const int SpeechChannels = 1;

    public SpeechRecognitionService(
        IConfiguration configuration)
    {
        _configuration =
            configuration;
    }

    private readonly object _liveLock = new();
    private SpeechRecognizer? _liveRecognizer;
    private PushAudioInputStream? _livePushStream;
    private bool _liveStarted;

    // ============================================================
    // PHASE 2: CONTINUOUS LIVE PCM FROM AUDIOSOCKET
    // RecognizeRecordingAsync below is unchanged for service-hosted.
    // ============================================================

    public async Task StartAsync()
    {
        lock (_liveLock)
        {
            if (_liveStarted)
            {
                return;
            }

            _liveStarted = true;
        }

        try
        {
            var key =
                _configuration["Speech:Key"];

            var region =
                _configuration["Speech:Region"];

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
            Console.WriteLine("Input         : 16 kHz 16-bit mono PCM push stream");
            Console.WriteLine("================================================");

            var speechConfig =
                SpeechConfig.FromSubscription(
                    key,
                    region);

            speechConfig.SpeechRecognitionLanguage =
                "en-US";

            var format =
                AudioStreamFormat.GetWaveFormatPCM(
                    SpeechSampleRate,
                    SpeechBitsPerSample,
                    SpeechChannels);

            _livePushStream =
                AudioInputStream.CreatePushStream(
                    format);

            var audioConfig =
                AudioConfig.FromStreamInput(
                    _livePushStream);

            _liveRecognizer =
                new SpeechRecognizer(
                    speechConfig,
                    audioConfig);

            _liveRecognizer.Recognized +=
                (_, e) =>
                {
                    if (e.Result.Reason ==
                            ResultReason.RecognizedSpeech &&
                        !string.IsNullOrWhiteSpace(
                            e.Result.Text))
                    {
                        Console.WriteLine();
                        Console.WriteLine("================================================");
                        Console.WriteLine(" LIVE SPEECH");
                        Console.WriteLine("================================================");
                        Console.WriteLine(e.Result.Text);
                        Console.WriteLine("================================================");
                    }
                };

            _liveRecognizer.Canceled +=
                (_, e) =>
                {
                    Console.WriteLine();
                    Console.WriteLine("================================================");
                    Console.WriteLine(" LIVE SPEECH SDK FAILURE");
                    Console.WriteLine("================================================");
                    Console.WriteLine($"Reason        : {e.Reason}");
                    Console.WriteLine($"Error code    : {e.ErrorCode}");
                    Console.WriteLine($"Error details : {e.ErrorDetails}");
                    Console.WriteLine("================================================");
                };

            await _liveRecognizer
                .StartContinuousRecognitionAsync();

            Console.WriteLine(
                "Live Azure Speech continuous recognition started.");
        }
        catch (Exception ex)
        {
            lock (_liveLock)
            {
                _liveStarted = false;
            }

            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" LIVE SPEECH SDK FAILURE");
            Console.WriteLine("================================================");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex);
            throw;
        }
    }

    public void ProcessAudio(
        byte[] audioData)
    {
        if (_livePushStream == null)
        {
            return;
        }

        if (audioData == null ||
            audioData.Length == 0)
        {
            return;
        }

        _livePushStream.Write(audioData);
    }

    public async Task StopAsync()
    {
        if (_liveRecognizer != null)
        {
            await _liveRecognizer
                .StopContinuousRecognitionAsync();
        }

        _livePushStream?.Close();
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
            _configuration["Speech:Key"];

        var region =
            _configuration["Speech:Region"];

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

            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" AZURE SPEECH RECOGNITION");
            Console.WriteLine("================================================");

            Console.WriteLine(
                $"Audio bytes   : {recordingBytes.Length}");

            Console.WriteLine(
                $"Speech region : {region}");

            var speechConfig =
                SpeechConfig.FromSubscription(
                    key,
                    region);

            speechConfig.SpeechRecognitionLanguage =
                "en-US";

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
}
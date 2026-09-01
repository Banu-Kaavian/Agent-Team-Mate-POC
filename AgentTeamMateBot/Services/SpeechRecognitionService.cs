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

    // ============================================================
    // OLD APP-HOSTED COMPATIBILITY
    // ============================================================

    public Task StartAsync()
    {
        Console.WriteLine(
            "[SPEECH] Raw PCM streaming disabled in service-hosted mode.");

        return Task.CompletedTask;
    }

    public void ProcessAudio(
        byte[] audioData)
    {
        Console.WriteLine(
            "[SPEECH] ProcessAudio is not used in service-hosted mode.");
    }

    public Task StopAsync()
    {
        return Task.CompletedTask;
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
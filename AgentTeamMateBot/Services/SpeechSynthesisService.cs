using Microsoft.CognitiveServices.Speech;
using Microsoft.CognitiveServices.Speech.Audio;

namespace AgentTeamMateBot.Services;

public class SpeechSynthesisService
{
    private readonly IConfiguration _configuration;
    private readonly IWebHostEnvironment _environment;

    public SpeechSynthesisService(
        IConfiguration configuration,
        IWebHostEnvironment environment)
    {
        _configuration = configuration;
        _environment = environment;
    }

    // ============================================================
    // CONVERT AI TEXT TO WAV AND RETURN PUBLIC URL
    // ============================================================

    public async Task<string?> SynthesizeSpeechAsync(
        string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            Console.WriteLine(
                "[TTS] Text is empty.");

            return null;
        }

        var key =
            _configuration["Speech:Key"];

        var region =
            _configuration["Speech:Region"];

        var voice =
            _configuration["Speech:Voice"]
            ?? "en-US-AvaMultilingualNeural";

        var callbackUri =
            _configuration["Bot:CallbackUri"];

        if (string.IsNullOrWhiteSpace(key))
        {
            throw new Exception(
                "Speech:Key missing");
        }

        if (string.IsNullOrWhiteSpace(region))
        {
            throw new Exception(
                "Speech:Region missing");
        }

        if (string.IsNullOrWhiteSpace(callbackUri))
        {
            throw new Exception(
                "Bot:CallbackUri missing");
        }

        // ============================================================
        // TEMP AUDIO DIRECTORY
        // ============================================================

        var audioDirectory =
            Path.Combine(
                _environment.ContentRootPath,
                "TempAudio");

        Directory.CreateDirectory(
            audioDirectory);

        var fileName =
            $"agent-response-{Guid.NewGuid():N}.wav";

        var filePath =
            Path.Combine(
                audioDirectory,
                fileName);

        try
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" AZURE SPEECH SYNTHESIS");
            Console.WriteLine("================================================");

            Console.WriteLine(
                $"Text   : {text}");

            Console.WriteLine(
                $"Region : {region}");

            Console.WriteLine(
                $"Voice  : {voice}");

            // ============================================================
            // SPEECH CONFIG
            // ============================================================

            var speechConfig =
                SpeechConfig.FromSubscription(
                    key,
                    region);

            speechConfig.SpeechSynthesisVoiceName =
                voice;

            speechConfig.SetSpeechSynthesisOutputFormat(
                SpeechSynthesisOutputFormat
                    .Riff16Khz16BitMonoPcm);

            // ============================================================
            // SAVE WAV
            // ============================================================

            using var audioConfig =
                AudioConfig.FromWavFileOutput(
                    filePath);

            using var synthesizer =
                new SpeechSynthesizer(
                    speechConfig,
                    audioConfig);

            Console.WriteLine();
            Console.WriteLine(
                "Generating speech...");

            var result =
                await synthesizer
                    .SpeakTextAsync(text);

            if (result.Reason ==
                ResultReason.SynthesizingAudioCompleted)
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine(
                        "[TTS] WAV file was not created.");

                    return null;
                }

                var fileInfo =
                    new FileInfo(filePath);

                Console.WriteLine();
                Console.WriteLine("================================================");
                Console.WriteLine(" SPEECH SYNTHESIS COMPLETED");
                Console.WriteLine("================================================");

                Console.WriteLine(
                    $"File        : {fileName}");

                Console.WriteLine(
                    $"Audio bytes : {fileInfo.Length}");

                Console.WriteLine(
                    "Format      : WAV 16 kHz / 16-bit / Mono");

                // ============================================================
                // BUILD PUBLIC URL
                // ============================================================

                var callbackBase =
                    callbackUri
                        .Replace(
                            "/api/calling",
                            "",
                            StringComparison.OrdinalIgnoreCase)
                        .TrimEnd('/');

                var publicAudioUrl =
                    $"{callbackBase}/api/audio/{fileName}";

                Console.WriteLine(
                    $"Audio URL   : {publicAudioUrl}");

                Console.WriteLine(
                    "================================================");

                return publicAudioUrl;
            }

            if (result.Reason ==
                ResultReason.Canceled)
            {
                var cancellation =
                    SpeechSynthesisCancellationDetails
                        .FromResult(result);

                Console.WriteLine();
                Console.WriteLine("================================================");
                Console.WriteLine(" AZURE TTS FAILURE");
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
                $"[TTS] Unexpected result: {result.Reason}");

            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" SPEECH SYNTHESIS SDK FAILURE");
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
}
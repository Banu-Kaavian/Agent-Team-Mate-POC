using Microsoft.CognitiveServices.Speech;

namespace AgentTeamMateBot.Services;

public class SpeechSynthesisService : IDisposable
{
    private readonly IConfiguration _configuration;
    private readonly IWebHostEnvironment _environment;
    private readonly SemaphoreSlim _synthesizeLock = new(1, 1);
    private readonly object _initLock = new();

    private SpeechSynthesizer? _synthesizer;
    private bool _warmed;

    public SpeechSynthesisService(
        IConfiguration configuration,
        IWebHostEnvironment environment)
    {
        _configuration = configuration;
        _environment = environment;
    }

    public Task WarmupAsync()
    {
        return Task.Run(async () =>
        {
            try
            {
                await _synthesizeLock.WaitAsync();
                try
                {
                    var synthesizer = EnsureSynthesizer();
                    using var result = await synthesizer.SpeakTextAsync(".");
                    _warmed = result.Reason == ResultReason.SynthesizingAudioCompleted;
                }
                finally
                {
                    _synthesizeLock.Release();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[TTS] Warmup failed: {ex.Message}");
            }
        });
    }

    public async Task<string?> SynthesizeSpeechAsync(
        string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            Console.WriteLine("[TTS] Text is empty.");
            return null;
        }

        var callbackUri = _configuration["Bot:CallbackUri"];
        if (string.IsNullOrWhiteSpace(callbackUri))
        {
            throw new Exception("Bot:CallbackUri missing");
        }

        var audioDirectory = Path.Combine(
            _environment.ContentRootPath,
            "TempAudio");

        Directory.CreateDirectory(audioDirectory);

        var fileName = $"agent-response-{Guid.NewGuid():N}.wav";
        var filePath = Path.Combine(audioDirectory, fileName);

        await _synthesizeLock.WaitAsync();
        try
        {
            var synthesizer = EnsureSynthesizer();

            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" AZURE SPEECH SYNTHESIS");
            Console.WriteLine("================================================");
            Console.WriteLine($"Text   : {text}");
            Console.WriteLine($"Warmed : {_warmed}");

            var result = await synthesizer.SpeakTextAsync(text);

            if (result.Reason == ResultReason.SynthesizingAudioCompleted)
            {
                await File.WriteAllBytesAsync(filePath, result.AudioData);

                var callbackBase = callbackUri
                    .Replace("/api/calling", "", StringComparison.OrdinalIgnoreCase)
                    .TrimEnd('/');

                var publicAudioUrl = $"{callbackBase}/api/audio/{fileName}";

                Console.WriteLine($"Audio bytes : {result.AudioData.Length}");
                Console.WriteLine($"Audio URL   : {publicAudioUrl}");
                Console.WriteLine("================================================");

                return publicAudioUrl;
            }

            if (result.Reason == ResultReason.Canceled)
            {
                var cancellation =
                    SpeechSynthesisCancellationDetails.FromResult(result);

                Console.WriteLine();
                Console.WriteLine("================================================");
                Console.WriteLine(" AZURE TTS FAILURE");
                Console.WriteLine("================================================");
                Console.WriteLine($"Reason  : {cancellation.Reason}");
                Console.WriteLine($"Code    : {cancellation.ErrorCode}");
                Console.WriteLine($"Details : {cancellation.ErrorDetails}");
                Console.WriteLine("================================================");

                ResetSynthesizer();
                return null;
            }

            Console.WriteLine($"[TTS] Unexpected result: {result.Reason}");
            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" SPEECH SYNTHESIS SDK FAILURE");
            Console.WriteLine("================================================");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex);
            Console.WriteLine("================================================");

            ResetSynthesizer();
            throw;
        }
        finally
        {
            _synthesizeLock.Release();
        }
    }

    private SpeechSynthesizer EnsureSynthesizer()
    {
        lock (_initLock)
        {
            if (_synthesizer != null)
            {
                return _synthesizer;
            }

            var key = _configuration["Speech:Key"]
                ?? _configuration["AZURE_SPEECH_KEY"];

            var region = _configuration["Speech:Region"]
                ?? _configuration["AZURE_SPEECH_REGION"];

            var voice = _configuration["Speech:Voice"]
                ?? "en-US-JennyNeural";

            if (string.IsNullOrWhiteSpace(key))
            {
                throw new Exception("Speech:Key missing");
            }

            if (string.IsNullOrWhiteSpace(region))
            {
                throw new Exception("Speech:Region missing");
            }

            var speechConfig = SpeechConfig.FromSubscription(key, region);
            speechConfig.SpeechSynthesisVoiceName = voice;
            speechConfig.SetSpeechSynthesisOutputFormat(
                SpeechSynthesisOutputFormat.Riff16Khz16BitMonoPcm);

            Console.WriteLine($"[TTS] Connecting synthesizer. Region={region} Voice={voice}");

            _synthesizer = new SpeechSynthesizer(speechConfig, audioConfig: null);
            return _synthesizer;
        }
    }

    private void ResetSynthesizer()
    {
        lock (_initLock)
        {
            _synthesizer?.Dispose();
            _synthesizer = null;
            _warmed = false;
        }
    }

    public void Dispose()
    {
        ResetSynthesizer();
        _synthesizeLock.Dispose();
    }
}

using Microsoft.CognitiveServices.Speech;
using Microsoft.CognitiveServices.Speech.Audio;

namespace AgentTeamMateBot.Services;

public class SpeechRecognitionService
{
    private readonly IConfiguration _configuration;
    private readonly object _startLock = new();

    private SpeechRecognizer? _recognizer;
    private PushAudioInputStream? _pushStream;
    private bool _started;

    public const int SpeechSampleRate = 16000;
    public const int SpeechBitsPerSample = 16;
    public const int SpeechChannels = 1;

    public SpeechRecognitionService(IConfiguration configuration)
    {
        _configuration = configuration;
    }

    public async Task StartAsync()
    {
        lock (_startLock)
        {
            if (_started)
            {
                return;
            }

            _started = true;
        }

        try
        {
            Console.WriteLine();
            Console.WriteLine("=================================");
            Console.WriteLine(" Azure Speech Service Starting ");
            Console.WriteLine("=================================");

            var key = _configuration["Speech:Key"];
            var region = _configuration["Speech:Region"];

            if (string.IsNullOrEmpty(key))
            {
                throw new Exception("Speech Key missing");
            }

            if (string.IsNullOrEmpty(region))
            {
                throw new Exception("Speech Region missing");
            }

            var speechConfig = SpeechConfig.FromSubscription(key, region);
            speechConfig.SpeechRecognitionLanguage = "en-US";

            var format = AudioStreamFormat.GetWaveFormatPCM(
                SpeechSampleRate,
                SpeechBitsPerSample,
                SpeechChannels);

            _pushStream = AudioInputStream.CreatePushStream(format);

            var audioConfig = AudioConfig.FromStreamInput(_pushStream);

            _recognizer = new SpeechRecognizer(speechConfig, audioConfig);

            _recognizer.Recognizing += (_, e) =>
            {
                if (!string.IsNullOrWhiteSpace(e.Result.Text))
                {
                    Console.WriteLine($"[Partial] {e.Result.Text}");
                }
            };

            _recognizer.Recognized += (_, e) =>
            {
                if (e.Result.Reason == ResultReason.RecognizedSpeech &&
                    !string.IsNullOrWhiteSpace(e.Result.Text))
                {
                    Console.WriteLine();
                    Console.WriteLine("================================================");
                    Console.WriteLine("TRANSCRIPT");
                    Console.WriteLine("================================================");
                    Console.WriteLine();
                    Console.WriteLine(e.Result.Text);
                    Console.WriteLine();
                    Console.WriteLine("================================================");
                }
                else if (e.Result.Reason == ResultReason.NoMatch)
                {
                    Console.WriteLine("[Speech] NoMatch for this audio segment");
                }
            };

            _recognizer.Canceled += (_, e) =>
            {
                Console.WriteLine();
                Console.WriteLine("================================================");
                Console.WriteLine(" SPEECH SDK FAILURE");
                Console.WriteLine("================================================");
                Console.WriteLine($"Reason        : {e.Reason}");
                Console.WriteLine($"Error code    : {e.ErrorCode}");
                Console.WriteLine($"Error details : {e.ErrorDetails}");
            };

            _recognizer.SessionStarted += (_, _) =>
            {
                Console.WriteLine("Azure Speech continuous recognition session started");
            };

            await _recognizer.StartContinuousRecognitionAsync();

            Console.WriteLine("Azure Speech ready (16 kHz 16-bit mono PCM push stream)");
        }
        catch (Exception ex)
        {
            lock (_startLock)
            {
                _started = false;
            }

            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" SPEECH SDK FAILURE");
            Console.WriteLine("================================================");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex);
            throw;
        }
    }

    public void ProcessAudio(byte[] audioData)
    {
        if (_pushStream == null)
        {
            Console.WriteLine("Speech stream not initialized");
            return;
        }

        if (audioData == null || audioData.Length == 0)
        {
            return;
        }

        _pushStream.Write(audioData);
    }

    public async Task StopAsync()
    {
        if (_recognizer != null)
        {
            await _recognizer.StopContinuousRecognitionAsync();
        }

        _pushStream?.Close();
    }
}

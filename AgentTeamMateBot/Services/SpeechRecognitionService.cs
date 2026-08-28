using Microsoft.CognitiveServices.Speech;
using Microsoft.CognitiveServices.Speech.Audio;

namespace AgentTeamMateBot.Services;


public class SpeechRecognitionService
{

    private readonly IConfiguration _configuration;

    private SpeechRecognizer? _recognizer;

    private PushAudioInputStream? _pushStream;



    public SpeechRecognitionService(
        IConfiguration configuration)
    {
        _configuration = configuration;
    }



    public async Task StartAsync()
    {

        Console.WriteLine();
        Console.WriteLine("=================================");
        Console.WriteLine(" Azure Speech Service Starting ");
        Console.WriteLine("=================================");


        var key =
            _configuration["Speech:Key"];


        var region =
            _configuration["Speech:Region"];



        if(string.IsNullOrEmpty(key))
        {
            throw new Exception(
                "Speech Key missing");
        }


        if(string.IsNullOrEmpty(region))
        {
            throw new Exception(
                "Speech Region missing");
        }



        var speechConfig =
            SpeechConfig.FromSubscription(
                key,
                region);



        speechConfig.SpeechRecognitionLanguage =
            "en-US";



        //
        // Input stream receives
        // REAL Teams audio packets
        //

        _pushStream =
            AudioInputStream.CreatePushStream();



        var audioConfig =
            AudioConfig.FromStreamInput(
                _pushStream);



        _recognizer =
            new SpeechRecognizer(
                speechConfig,
                audioConfig);



        _recognizer.Recognizing +=
        (sender, e) =>
        {

            Console.WriteLine(
                $"[Partial] {e.Result.Text}");

        };



        _recognizer.Recognized +=
        (sender, e) =>
        {

            if(e.Result.Reason ==
                ResultReason.RecognizedSpeech)
            {

                Console.WriteLine();
                Console.WriteLine(
                "==============================");

                Console.WriteLine(
                " TRANSCRIPT ");

                Console.WriteLine(
                e.Result.Text);

                Console.WriteLine(
                "==============================");

            }

        };



        _recognizer.Canceled +=
        (sender, e) =>
        {

            Console.WriteLine(
                $"Speech cancelled : {e.ErrorDetails}");

        };



        await _recognizer
            .StartContinuousRecognitionAsync();



        Console.WriteLine(
            "Azure Speech ready");

    }





    public async Task ProcessAudioAsync(
        byte[] audioData)
    {


        if(_pushStream == null)
        {
            Console.WriteLine(
                "Speech stream not initialized");

            return;
        }



        Console.WriteLine(
            $"Sending audio to Speech : {audioData.Length} bytes");



        _pushStream.Write(audioData);



        await Task.CompletedTask;

    }

}
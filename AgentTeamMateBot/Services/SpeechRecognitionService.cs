using Microsoft.CognitiveServices.Speech;

public class SpeechRecognitionService
{
    private readonly IConfiguration _configuration;

    public SpeechRecognitionService(
        IConfiguration configuration)
    {
        _configuration = configuration;
    }


    public async Task StartAsync()
    {
        Console.WriteLine();
        Console.WriteLine("=================================");
        Console.WriteLine(" Azure Speech Service Started");
        Console.WriteLine("=================================");


        var key =
            _configuration["Speech:Key"];

        var region =
            _configuration["Speech:Region"];


        if (string.IsNullOrEmpty(key))
        {
            Console.WriteLine(
                "ERROR: Speech Key missing");

            return;
        }


        if (string.IsNullOrEmpty(region))
        {
            Console.WriteLine(
                "ERROR: Speech Region missing");

            return;
        }


        Console.WriteLine(
            $"Region : {region}");


        var speechConfig =
            SpeechConfig.FromSubscription(
                key,
                region);


        speechConfig.SpeechRecognitionLanguage =
            "en-US";


        using var recognizer =
            new SpeechRecognizer(
                speechConfig);


        recognizer.Recognizing +=
            (sender, e) =>
            {
                Console.WriteLine(
                    $"[Listening] {e.Result.Text}");
            };


        recognizer.Recognized +=
            (sender, e) =>
            {

                if (!string.IsNullOrEmpty(
                    e.Result.Text))
                {
                    Console.WriteLine();
                    Console.WriteLine(
                    "------------------------------");

                    Console.WriteLine(
                    " FINAL TEXT");

                    Console.WriteLine(
                    e.Result.Text);

                    Console.WriteLine(
                    "------------------------------");
                }

            };


        recognizer.Canceled +=
            (sender, e) =>
            {
                Console.WriteLine(
                    $"Speech Error: {e.ErrorDetails}");
            };


        await recognizer
            .StartContinuousRecognitionAsync();


        Console.WriteLine(
            "Speech recognition is listening...");
    }
}
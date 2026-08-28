using AgentTeamMateBot.Services;

namespace AgentTeamMateBot.Media;


public class AudioHandler
{

    private readonly SpeechRecognitionService _speechService;


    public AudioHandler(
        SpeechRecognitionService speechService)
    {
        _speechService = speechService;
    }



    public async Task ProcessAudio(
        byte[] data)
    {

        if(data == null || data.Length == 0)
        {
            Console.WriteLine(
                "[AUDIO] Empty packet received");

            return;
        }


        Console.WriteLine();
        Console.WriteLine("--------------------------------");
        Console.WriteLine("TEAMS AUDIO PACKET RECEIVED");
        Console.WriteLine(
            $"Size : {data.Length} bytes");
        Console.WriteLine("--------------------------------");



        // Send real Teams audio stream
        // to Azure Speech Service

        await _speechService
            .ProcessAudioAsync(data);

    }

}
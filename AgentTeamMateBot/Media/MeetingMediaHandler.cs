namespace AgentTeamMateBot.Media;

public class MeetingMediaHandler
{
    public void OnAudioReceived(byte[] audioData)
    {
        Console.WriteLine();
        Console.WriteLine("==============================");
        Console.WriteLine("       TEAMS AUDIO RECEIVED");
        Console.WriteLine("==============================");

        Console.WriteLine(
            $"Audio packet size : {audioData.Length} bytes");

        Console.WriteLine("==============================");
    }
}
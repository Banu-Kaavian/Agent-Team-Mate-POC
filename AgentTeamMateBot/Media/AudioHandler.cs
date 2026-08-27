namespace AgentTeamMateBot.Media;

public class AudioHandler
{
    public void ProcessAudio(byte[] data)
    {
        Console.WriteLine(
            $"[AUDIO] Received {data.Length} bytes");
    }
}
using Microsoft.Graph.Communications.Calls.Media;

namespace AgentTeamMateBot.Media;

public class MeetingMediaHandler
{
    public void OnAudioReceived(byte[] audioData)
    {
        Console.WriteLine(
            $"Audio received bytes: {audioData.Length}"
        );
    }
}
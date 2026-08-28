using AgentTeamMateBot.Media;

namespace AgentTeamMateBot.Services;


public class MediaSessionService
{

    private readonly AudioHandler _audioHandler;


    public MediaSessionService(
        AudioHandler audioHandler)
    {
        _audioHandler = audioHandler;
    }



    public async Task StartMediaSessionAsync(
        string callId)
    {

        Console.WriteLine();
        Console.WriteLine("=================================");
        Console.WriteLine(" MEDIA SESSION START REQUEST");
        Console.WriteLine("=================================");


        Console.WriteLine(
            $"Call ID : {callId}");



        /*
        
        REAL IMPLEMENTATION FLOW:

        1. Get active Graph Call object
        2. Attach MediaSession
        3. Create AudioSocket
        4. Subscribe AudioReceived event
        5. Push audio bytes to AudioHandler


        AudioSocket
             |
             |
        AudioHandler.ProcessAudio()
             |
             |
        SpeechRecognitionService
        

        */



        await Task.CompletedTask;

    }




    public async Task OnAudioReceived(
        byte[] audioBytes)
    {

        Console.WriteLine(
            "REAL AUDIO PACKET RECEIVED");


        Console.WriteLine(
            $"Bytes : {audioBytes.Length}");



        await _audioHandler
            .ProcessAudio(audioBytes);

    }

}
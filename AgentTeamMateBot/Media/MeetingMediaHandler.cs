using System.Text.Json;
using AgentTeamMateBot.Services;

namespace AgentTeamMateBot.Media;

public class MeetingMediaHandler
{
    private readonly MediaSessionService _mediaSessionService;


    public MeetingMediaHandler(
        MediaSessionService mediaSessionService)
    {
        _mediaSessionService = mediaSessionService;
    }



    public async Task ProcessNotification(string notificationBody)
    {
        Console.WriteLine();
        Console.WriteLine("======================================");
        Console.WriteLine(" PROCESSING TEAMS NOTIFICATION");
        Console.WriteLine("======================================");


        try
        {
            using var json =
                JsonDocument.Parse(notificationBody);


            var root = json.RootElement;


            if (!root.TryGetProperty("value", out var values))
            {
                Console.WriteLine(
                    "No notification value found");

                return;
            }



            foreach(var item in values.EnumerateArray())
            {

                if(!item.TryGetProperty(
                    "resourceData",
                    out var resourceData))
                {
                    continue;
                }



                foreach(var participant 
                    in resourceData.EnumerateArray())
                {


                    if(participant.TryGetProperty(
                        "mediaStreams",
                        out var streams))
                    {

                        foreach(var stream in streams.EnumerateArray())
                        {

                            if(stream.TryGetProperty(
                                "mediaType",
                                out var mediaType)
                                &&
                                mediaType
                                .GetString()
                                ==
                                "audio")
                            {

                                Console.WriteLine(
                                    "Audio stream detected");


                                if(item.TryGetProperty(
                                    "resource",
                                    out var resource))
                                {

                                    var resourceValue =
                                        resource.GetString();


                                    Console.WriteLine(
                                    $"Resource : {resourceValue}");



                                    var callId =
                                        ExtractCallId(
                                            resourceValue);



                                    if(callId != null)
                                    {

                                        Console.WriteLine(
                                        $"Call ID : {callId}");



                                        await _mediaSessionService
                                            .StartMediaSessionAsync(
                                                callId);

                                    }

                                }

                            }

                        }

                    }

                }

            }


        }
        catch(Exception ex)
        {

            Console.WriteLine(
                $"Notification processing error : {ex.Message}");

        }


    }



    private string? ExtractCallId(
        string? resource)
    {

        if(string.IsNullOrEmpty(resource))
            return null;


        // Example:
        // /app/calls/{callId}/participants


        var parts =
            resource.Split('/');


        var index =
            Array.IndexOf(
                parts,
                "calls");


        if(index >= 0 &&
           parts.Length > index + 1)
        {
            return parts[index + 1];
        }


        return null;
    }
}
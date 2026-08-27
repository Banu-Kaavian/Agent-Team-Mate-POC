using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;


// ============================================================
// APPLICATION START
// ============================================================

var builder = WebApplication.CreateBuilder(args);


Console.WriteLine();
Console.WriteLine("================================================");
Console.WriteLine("        AGENT TEAM MATE BOT STARTING");
Console.WriteLine("================================================");


// ============================================================
// SERVICES
// ============================================================

builder.Services.AddHttpClient();


// ============================================================
// CONFIGURATION
// ============================================================


var clientId =
    builder.Configuration["Bot:ClientId"]
    ?? throw new Exception("Bot:ClientId missing");


var tenantId =
    builder.Configuration["Bot:TenantId"]
    ?? throw new Exception("Bot:TenantId missing");


var clientSecret =
    builder.Configuration["Bot:ClientSecret"]
    ?? throw new Exception("Bot:ClientSecret missing");


var callbackUri =
    builder.Configuration["Bot:CallbackUri"]
    ??
    "https://teammate-bot.westus3.cloudapp.azure.com/api/calling";



// ============================================================
// MICROSOFT GRAPH AUTH
// ============================================================


var confidentialClient =
    ConfidentialClientApplicationBuilder
    .Create(clientId)
    .WithClientSecret(clientSecret)
    .WithAuthority(
        $"https://login.microsoftonline.com/{tenantId}")
    .Build();



var app = builder.Build();



// ============================================================
// HEALTH CHECK
// ============================================================


app.MapGet("/", () =>
{

    Console.WriteLine("[HEALTH] Request received");


    return Results.Ok(new
    {
        Application = "Agent Team Mate",
        Status = "Running",
        Time = DateTime.UtcNow
    });

});



// ============================================================
// AUTH TEST
// ============================================================


app.MapGet("/auth-test",
async () =>
{

    try
    {

        Console.WriteLine(
        "[AUTH] Requesting Graph token");


        var result =
        await confidentialClient
        .AcquireTokenForClient(
        new[]
        {
            "https://graph.microsoft.com/.default"
        })
        .ExecuteAsync();



        Console.WriteLine(
        "[AUTH] SUCCESS");


        return Results.Ok(new
        {
            Message =
            "Microsoft Graph authentication successful",

            Expires =
            result.ExpiresOn
        });

    }
    catch(Exception ex)
    {

        Console.WriteLine(
        $"AUTH ERROR : {ex.Message}");


        return Results.BadRequest(ex.Message);

    }

});



// ============================================================
// TEAMS CALLBACK
// ============================================================


app.MapPost("/api/calling",
async(HttpRequest request)=>
{

    try
    {


        var body =
        await new StreamReader(
            request.Body)
        .ReadToEndAsync();



        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine("       TEAMS CALLING EVENT RECEIVED");
        Console.WriteLine("================================================");


        Console.WriteLine(
        $"Time : {DateTime.UtcNow}");



        var callId =
        ExtractCallId(body);



        if(callId != null)
        {

            Console.WriteLine();
            Console.WriteLine(
            $"CALL ID : {callId}");


            Console.WriteLine();
            Console.WriteLine(
            "MEDIA SESSION START POINT");

            /*
            
            NEXT IMPLEMENTATION:

            callId
              |
              |
              v

            Microsoft Graph Communications SDK

              |
              |
              v

            MediaSession

              |
              |
              v

            AudioSocket

              |
              |
              v

            Azure Speech


            */

        }



        Console.WriteLine();
        Console.WriteLine(
        "EVENT PAYLOAD");


        Console.WriteLine(body);


        Console.WriteLine();
        Console.WriteLine(
        "Webhook completed");


        Console.WriteLine(
        "================================================");



        // Important:
        // Teams requires HTTP 200
        return Results.Ok();

    }
    catch(Exception ex)
    {


        Console.WriteLine(
        $"CALLBACK ERROR : {ex.Message}");


        return Results.Ok();

    }

});



// ============================================================
// JOIN MEETING
// ============================================================


app.MapPost("/api/join",
async(
JoinRequest request,
IHttpClientFactory factory)=>
{


try
{


Console.WriteLine();
Console.WriteLine("==============================");
Console.WriteLine(" JOIN REQUEST ");
Console.WriteLine("==============================");



if(string.IsNullOrWhiteSpace(request.MeetingId))
{

return Results.BadRequest(
new
{
Message =
"MeetingId required"
});

}



var token =
await confidentialClient
.AcquireTokenForClient(
new[]
{
"https://graph.microsoft.com/.default"
})
.ExecuteAsync();



var payload =
new
{

@odata_type =
"#microsoft.graph.call",


callbackUri,


requestedModalities =
new[]
{
"audio"
},


mediaConfig =
new
{
@odata_type =
"#microsoft.graph.serviceHostedMediaConfig"
},


meetingInfo =
new
{

@odata_type =
"#microsoft.graph.joinMeetingIdMeetingInfo",

joinMeetingId =
request.MeetingId,

passcode =
request.Passcode

},


tenantId

};



var json =
JsonSerializer.Serialize(payload);



var client =
factory.CreateClient();



var graphRequest =
new HttpRequestMessage(
HttpMethod.Post,
"https://graph.microsoft.com/v1.0/communications/calls");



graphRequest.Headers.Authorization =
new AuthenticationHeaderValue(
"Bearer",
token.AccessToken);



graphRequest.Content =
new StringContent(
json,
Encoding.UTF8,
"application/json");



var response =
await client.SendAsync(graphRequest);



var responseText =
await response.Content.ReadAsStringAsync();



Console.WriteLine(
$"GRAPH STATUS : {response.StatusCode}");



return Results.Ok(
new
{
Status =
response.StatusCode,

Response =
responseText
});


}
catch(Exception ex)
{

return Results.BadRequest(
ex.Message);

}


});



// ============================================================
// HELPERS
// ============================================================


string? ExtractCallId(string payload)
{


var match =
System.Text.RegularExpressions.Regex.Match(
payload,
@"/app/calls/([^/]+)"
);


if(match.Success)
{
    return match.Groups[1].Value;
}


return null;

}




// ============================================================
// START SERVER
// ============================================================


Console.WriteLine();
Console.WriteLine("==============================");
Console.WriteLine(" Agent Team Mate Ready");
Console.WriteLine("==============================");

Console.WriteLine(
$"Callback : {callbackUri}");



app.Run();




// ============================================================
// REQUEST MODEL
// ============================================================


public record JoinRequest(
string MeetingId,
string? Passcode);
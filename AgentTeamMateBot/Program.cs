using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

using AgentTeamMateBot.Media;
using AgentTeamMateBot.Services;


// ============================================================
// APPLICATION START
// ============================================================

var builder = WebApplication.CreateBuilder(args);

Console.WriteLine(
    $"Environment : {builder.Environment.EnvironmentName}");

Console.WriteLine(
    $"ClientId : {builder.Configuration["Bot:ClientId"]}");

Console.WriteLine(
    $"TenantId : {builder.Configuration["Bot:TenantId"]}");


Console.WriteLine();
Console.WriteLine("================================================");
Console.WriteLine("        AGENT TEAM MATE BOT STARTING");
Console.WriteLine("================================================");


// ============================================================
// SERVICES
// ============================================================


builder.Services.AddHttpClient();


// Azure Speech Service
builder.Services.AddSingleton<SpeechRecognitionService>();


// Media Service (Real Teams Audio Handler)
builder.Services.AddSingleton<MediaSessionService>();


// Teams Media Handler
builder.Services.AddSingleton<MeetingMediaHandler>();


builder.Services.AddSingleton<AudioHandler>();
builder.Services.AddSingleton<GraphAuthService>();

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

    return Results.Ok(new
    {
        Application = "Agent Team Mate",
        Status = "Running",
        Time = DateTime.UtcNow
    });

});





// ============================================================
// GRAPH AUTH TEST
// ============================================================


app.MapGet("/auth-test",
async () =>
{

    try
    {

        var token =
            await confidentialClient
            .AcquireTokenForClient(
            new[]
            {
                "https://graph.microsoft.com/.default"
            })
            .ExecuteAsync();


        Console.WriteLine(
            "[AUTH] Graph authentication successful");


        return Results.Ok(new
        {
            Message =
            "Microsoft Graph authentication successful",

            Expires =
            token.ExpiresOn
        });

    }

    catch(Exception ex)
    {

        return Results.BadRequest(new
        {
            Message =
            "Authentication failed",

            Error =
            ex.Message
        });

    }

});





// ============================================================
// TEAMS CALLING CALLBACK
// ============================================================
//
// Microsoft Graph sends:
// - call events
// - participant events
// - media state changes
//
// Audio DOES NOT come here.
// Audio comes through Media SDK.
//

app.MapPost("/api/calling",
async (
HttpRequest request,
MeetingMediaHandler mediaHandler) =>
{

    try
    {

        var body =
            await new StreamReader(request.Body)
            .ReadToEndAsync();



        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine("       TEAMS EVENT RECEIVED");
        Console.WriteLine("================================================");


        Console.WriteLine(
            DateTime.UtcNow);


        Console.WriteLine(body);


        Console.WriteLine("================================================");



        await mediaHandler.ProcessNotification(body);



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
async (
JoinRequest request,
IHttpClientFactory factory) =>
{


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

        }

    };



    var client =
        factory.CreateClient();



    var http =
        new HttpRequestMessage(
            HttpMethod.Post,
            "https://graph.microsoft.com/v1.0/communications/calls");



    http.Headers.Authorization =
        new AuthenticationHeaderValue(
            "Bearer",
            token.AccessToken);



    http.Content =
        new StringContent(
            JsonSerializer.Serialize(payload),
            Encoding.UTF8,
            "application/json");



    var response =
        await client.SendAsync(http);



    return Results.Ok(
    new
    {
        Status =
        response.StatusCode,

        Response =
        await response.Content.ReadAsStringAsync()
    });


});





Console.WriteLine();
Console.WriteLine("================================================");
Console.WriteLine(" Agent Team Mate Bot Ready");
Console.WriteLine($" Callback : {callbackUri}");
Console.WriteLine("================================================");



app.Run();





public record JoinRequest(
    string MeetingId,
    string? Passcode
);
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
Console.WriteLine();


// ============================================================
// SERVICES
// ============================================================

// HTTP client for Microsoft Graph calls
builder.Services.AddHttpClient();


// Azure Speech Service
builder.Services.AddSingleton<SpeechRecognitionService>();



// ============================================================
// BOT CONFIGURATION
// ============================================================


var clientId =
    builder.Configuration["Bot:ClientId"]
    ?? throw new Exception(
        "Bot:ClientId missing");


var tenantId =
    builder.Configuration["Bot:TenantId"]
    ?? throw new Exception(
        "Bot:TenantId missing");


var clientSecret =
    builder.Configuration["Bot:ClientSecret"]
    ?? throw new Exception(
        "Bot:ClientSecret missing");



var callbackUri =
    builder.Configuration["Bot:CallbackUri"]
    ??
    "https://teammate-bot.westus3.cloudapp.azure.com/api/calling";



// ============================================================
// MICROSOFT ENTRA AUTH CLIENT
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

    Console.WriteLine(
        "[HEALTH] Application check received");


    return Results.Ok(new
    {
        Application =
            "Agent Team Mate",

        Status =
            "Running",

        Time =
            DateTime.UtcNow
    });

});



// ============================================================
// GRAPH AUTHENTICATION TEST
// ============================================================


app.MapGet("/auth-test",
async () =>
{

    try
    {

        Console.WriteLine(
            "[AUTH] Requesting Microsoft Graph token");


        var result =
            await confidentialClient
            .AcquireTokenForClient(
            new[]
            {
                "https://graph.microsoft.com/.default"
            })
            .ExecuteAsync();



        Console.WriteLine(
            "[AUTH] Microsoft Graph authentication successful");


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
            $"[AUTH ERROR] {ex.Message}");


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
// TEAMS CALLING WEBHOOK
// ============================================================
//
// Microsoft Teams / Graph sends call events here
//
// Current purpose:
// - Receive participant events
// - Receive call state changes
//
// Later:
// - Create media session
// - Receive audio stream
//
// ============================================================


app.MapPost("/api/calling",
async (HttpRequest request) =>
{

    try
    {

        var body =
            await new StreamReader(
                request.Body)
            .ReadToEndAsync();



        Console.WriteLine();
        Console.WriteLine(
        "================================================");

        Console.WriteLine(
        "       TEAMS CALLING EVENT RECEIVED");

        Console.WriteLine(
        $"Time : {DateTime.UtcNow}");

        Console.WriteLine(
        "================================================");


        Console.WriteLine(body);


        Console.WriteLine(
        "================================================");
        Console.WriteLine();


        return Results.Ok();

    }

    catch(Exception ex)
    {

        Console.WriteLine(
        $"[CALLBACK ERROR] {ex.Message}");


        return Results.Ok();

    }

});




// ============================================================
// JOIN TEAMS MEETING
// ============================================================


app.MapPost(
"/api/join",

async (
JoinRequest request,
IHttpClientFactory httpClientFactory) =>
{

    try
    {


        Console.WriteLine();
        Console.WriteLine(
        "================================================");

        Console.WriteLine(
        "       TEAMS JOIN REQUEST");

        Console.WriteLine(
        "================================================");



        if(string.IsNullOrWhiteSpace(request.MeetingId))
        {

            return Results.BadRequest(
            new
            {
                Message =
                "meetingId is required"
            });

        }



        var meetingId =
            request.MeetingId
            .Replace(" ","")
            .Trim();



        var passcode =
            request.Passcode?
            .Trim();



        Console.WriteLine(
        $"Meeting ID : {meetingId}");

        Console.WriteLine(
        $"Callback   : {callbackUri}");



        // Get Graph token

        var authResult =
            await confidentialClient
            .AcquireTokenForClient(
            new[]
            {
                "https://graph.microsoft.com/.default"
            })
            .ExecuteAsync();



        Console.WriteLine(
        "[GRAPH] Token generated");



        var payload =
        new Dictionary<string,object?>
        {


            ["@odata.type"] =
            "#microsoft.graph.call",


            ["callbackUri"] =
            callbackUri,


            ["requestedModalities"] =
            new[]
            {
                "audio"
            },


            ["mediaConfig"] =
            new Dictionary<string,string>
            {
                ["@odata.type"] =
                "#microsoft.graph.serviceHostedMediaConfig"
            },


            ["meetingInfo"] =
            new Dictionary<string,object?>
            {

                ["@odata.type"] =
                "#microsoft.graph.joinMeetingIdMeetingInfo",


                ["joinMeetingId"] =
                meetingId,


                ["passcode"] =
                passcode

            },


            ["tenantId"] =
            tenantId

        };



        var json =
            JsonSerializer.Serialize(payload);



        var client =
            httpClientFactory.CreateClient();



        using var graphRequest =
            new HttpRequestMessage(
                HttpMethod.Post,
                "https://graph.microsoft.com/v1.0/communications/calls");



        graphRequest.Headers.Authorization =
            new AuthenticationHeaderValue(
                "Bearer",
                authResult.AccessToken);



        graphRequest.Content =
            new StringContent(
                json,
                Encoding.UTF8,
                "application/json");



        Console.WriteLine(
        "[GRAPH] Sending join request");



        var response =
            await client.SendAsync(graphRequest);



        var responseBody =
            await response.Content
            .ReadAsStringAsync();



        Console.WriteLine(
        $"[GRAPH] Status : {(int)response.StatusCode}");



        if(!response.IsSuccessStatusCode)
        {

            Console.WriteLine(
            "[GRAPH] Join failed");


            return Results.Json(
            new
            {

                Message =
                "Microsoft Graph join failed",

                Status =
                response.StatusCode,

                Response =
                responseBody

            },

            statusCode:
            (int)response.StatusCode);

        }



        Console.WriteLine(
        "[GRAPH] Join accepted");


        return Results.Ok(
        new
        {

            Message =
            "Teams meeting join request accepted",

            Status =
            (int)response.StatusCode,

            Response =
            responseBody

        });


    }

    catch(Exception ex)
    {

        Console.WriteLine(
        $"[JOIN ERROR] {ex.Message}");


        return Results.BadRequest(
        new
        {
            Message =
            "Join failed",

            Error =
            ex.Message
        });

    }

});




// ============================================================
// AZURE SPEECH TEST
// ============================================================


app.MapGet(
"/speech-test",

async(SpeechRecognitionService speech)=>
{

    Console.WriteLine();

    Console.WriteLine(
    "[SPEECH] Test started");


    await speech.StartAsync();


    return Results.Ok(
    new
    {
        Message =
        "Azure Speech started"
    });

});




// ============================================================
// START SERVER
// ============================================================


Console.WriteLine();
Console.WriteLine(
"================================================");

Console.WriteLine(
" Agent Team Mate Bot Ready");

Console.WriteLine(
$" Callback : {callbackUri}");

Console.WriteLine(
"================================================");


app.Run();





// ============================================================
// REQUEST MODEL
// ============================================================


public record JoinRequest(
    string MeetingId,
    string? Passcode
);
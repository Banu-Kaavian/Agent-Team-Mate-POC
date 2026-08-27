using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

var builder = WebApplication.CreateBuilder(args);

// Used for calling Microsoft Graph
builder.Services.AddHttpClient();


// ============================================================
// 1. READ BOT / APP REGISTRATION CONFIGURATION
// ============================================================

var clientId = builder.Configuration["Bot:ClientId"]
    ?? throw new Exception("Bot:ClientId missing");

var tenantId = builder.Configuration["Bot:TenantId"]
    ?? throw new Exception("Bot:TenantId missing");

var clientSecret = builder.Configuration["Bot:ClientSecret"]
    ?? throw new Exception("Bot:ClientSecret missing");


// Public HTTPS endpoint where Microsoft Teams / Graph
// will send call status notifications.
var callbackUri =
    builder.Configuration["Bot:CallbackUri"]
    ?? "https://teammate-bot.westus3.cloudapp.azure.com/api/calling";


// ============================================================
// 2. CREATE MICROSOFT ENTRA AUTHENTICATION CLIENT
// ============================================================

var confidentialClient = ConfidentialClientApplicationBuilder
    .Create(clientId)
    .WithClientSecret(clientSecret)
    .WithAuthority(
        $"https://login.microsoftonline.com/{tenantId}")
    .Build();


var app = builder.Build();


// ============================================================
// 3. HEALTH CHECK
// ============================================================

app.MapGet("/", () =>
{
    return Results.Ok(new
    {
        message = "Agent Team Mate is running"
    });
});


// ============================================================
// 4. MICROSOFT GRAPH AUTHENTICATION TEST
// ============================================================

app.MapGet("/auth-test", async () =>
{
    try
    {
        var result = await confidentialClient
            .AcquireTokenForClient(
                new[]
                {
                    "https://graph.microsoft.com/.default"
                })
            .ExecuteAsync();

        return Results.Ok(new
        {
            message =
                "Microsoft Graph authentication successful",

            expiresOn = result.ExpiresOn
        });
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new
        {
            message = "Authentication failed",
            error = ex.Message
        });
    }
});


// ============================================================
// 5. MICROSOFT TEAMS CALLING CALLBACK
// ============================================================
//
// Microsoft Graph / Teams will POST call status notifications
// to this endpoint.
//
// For Phase 1 we simply log the notification and return 200.
// Later this will handle call state, media session etc.
// ============================================================

app.MapPost("/api/calling", async (HttpRequest request) =>
{
    try
    {
        using var reader =
            new StreamReader(request.Body);

        var body =
            await reader.ReadToEndAsync();

        Console.WriteLine(
            "====================================");

        Console.WriteLine(
            "Teams calling notification received");

        Console.WriteLine(body);

        Console.WriteLine(
            "====================================");

        return Results.Ok();
    }
    catch (Exception ex)
    {
        Console.WriteLine(
            $"Calling callback error: {ex.Message}");

        return Results.Ok();
    }
});


// ============================================================
// 6. JOIN A MICROSOFT TEAMS MEETING
// ============================================================
//
// Input:
//
// {
//   "meetingId": "...",
//   "passcode": "..."
// }
//
// Flow:
//
// Postman
//   -> /api/join
//   -> Entra authentication
//   -> Graph access token
//   -> POST /communications/calls
//   -> Microsoft Teams meeting
//
// ============================================================

app.MapPost(
    "/api/join",
    async (
        JoinRequest request,
        IHttpClientFactory httpClientFactory) =>
{
    try
    {
        // ----------------------------------------------------
        // Validate Meeting ID
        // ----------------------------------------------------

        if (string.IsNullOrWhiteSpace(request.MeetingId))
        {
            return Results.BadRequest(new
            {
                message = "meetingId is required"
            });
        }


        // Teams displays meeting IDs with spaces.
        // Graph expects the actual meeting ID value.
        var meetingId =
            request.MeetingId
                .Replace(" ", "")
                .Trim();

        var passcode =
            string.IsNullOrWhiteSpace(request.Passcode)
                ? null
                : request.Passcode.Trim();


        // ----------------------------------------------------
        // Get Microsoft Graph access token
        // ----------------------------------------------------

        var authResult = await confidentialClient
            .AcquireTokenForClient(
                new[]
                {
                    "https://graph.microsoft.com/.default"
                })
            .ExecuteAsync();


        // ----------------------------------------------------
        // Build Microsoft Graph Create Call request
        // ----------------------------------------------------

        var meetingInfo =
            new Dictionary<string, object?>
            {
                ["@odata.type"] =
                    "#microsoft.graph.joinMeetingIdMeetingInfo",

                ["joinMeetingId"] =
                    meetingId,

                ["passcode"] =
                    passcode
            };


        var mediaConfig =
            new Dictionary<string, object?>
            {
                ["@odata.type"] =
                    "#microsoft.graph.serviceHostedMediaConfig"
            };


        var callPayload =
            new Dictionary<string, object?>
            {
                ["@odata.type"] =
                    "#microsoft.graph.call",

                ["callbackUri"] =
                    callbackUri,

                ["requestedModalities"] =
                    new[] { "audio" },

                ["mediaConfig"] =
                    mediaConfig,

                ["meetingInfo"] =
                    meetingInfo,

                // For our first test the meeting should be
                // created in the same tenant as the bot.
                ["tenantId"] =
                    tenantId
            };


        var json =
            JsonSerializer.Serialize(callPayload);


        Console.WriteLine(
            "Sending Teams meeting join request...");

        Console.WriteLine(
            $"Meeting ID: {meetingId}");

        Console.WriteLine(
            $"Callback: {callbackUri}");


        // ----------------------------------------------------
        // Call Microsoft Graph
        // ----------------------------------------------------

        var httpClient =
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


        var graphResponse =
            await httpClient.SendAsync(graphRequest);


        var responseBody =
            await graphResponse.Content
                .ReadAsStringAsync();


        Console.WriteLine(
            $"Graph Status: {(int)graphResponse.StatusCode}");

        Console.WriteLine(responseBody);


        // ----------------------------------------------------
        // Graph rejected the call
        // ----------------------------------------------------

        if (!graphResponse.IsSuccessStatusCode)
        {
            return Results.Json(
                new
                {
                    message =
                        "Microsoft Graph meeting join failed",

                    graphStatusCode =
                        (int)graphResponse.StatusCode,

                    graphStatus =
                        graphResponse.StatusCode.ToString(),

                    graphResponse =
                        responseBody
                },

                statusCode:
                    (int)graphResponse.StatusCode
            );
        }


        // ----------------------------------------------------
        // Graph accepted the call
        // ----------------------------------------------------

        return Results.Ok(new
        {
            message =
                "Microsoft Graph accepted the Teams meeting join request",

            graphStatusCode =
                (int)graphResponse.StatusCode,

            graphResponse =
                responseBody
        });
    }
    catch (Exception ex)
    {
        Console.WriteLine(ex);

        return Results.BadRequest(new
        {
            message =
                "Join request failed",

            error =
                ex.Message
        });
    }
});


app.Run();


// ============================================================
// REQUEST MODEL
// ============================================================

public record JoinRequest(
    string MeetingId,
    string? Passcode
);
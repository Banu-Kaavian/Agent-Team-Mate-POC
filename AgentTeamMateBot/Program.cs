using System.Security.Cryptography.X509Certificates;
using AgentTeamMateBot.Media;
using AgentTeamMateBot.Services;

var builder = WebApplication.CreateBuilder(args);

Console.WriteLine($"Environment : {builder.Environment.EnvironmentName}");
Console.WriteLine($"ClientId : {builder.Configuration["Bot:ClientId"]}");
Console.WriteLine($"TenantId : {builder.Configuration["Bot:TenantId"]}");

Console.WriteLine();
Console.WriteLine("================================================");
Console.WriteLine("        AGENT TEAM MATE BOT STARTING");
Console.WriteLine("================================================");

// ================================================================
// KESTREL - LOCAL HTTP + PUBLIC HTTPS
// ================================================================

var serviceFqdn = builder.Configuration["Bot:ServiceFqdn"]
    ?? "teammate-bot.westus3.cloudapp.azure.com";

X509Certificate2? httpsCertificate = null;

using (var store = new X509Store(StoreName.My, StoreLocation.LocalMachine))
{
    store.Open(OpenFlags.ReadOnly);

    var certificates = store.Certificates.Find(
        X509FindType.FindBySubjectName,
        serviceFqdn,
        validOnly: true);

    if (certificates.Count > 0)
    {
        httpsCertificate = certificates[0];

        Console.WriteLine($"HTTPS certificate loaded : {httpsCertificate.Subject}");
        Console.WriteLine($"HTTPS thumbprint         : {httpsCertificate.Thumbprint}");
    }
    else
    {
        Console.WriteLine($"HTTPS certificate not found for {serviceFqdn}");
    }
}

builder.WebHost.ConfigureKestrel(options =>
{
    // Local endpoint
    options.ListenLocalhost(5000);

    // Public HTTPS endpoint
    if (httpsCertificate != null)
    {
        options.ListenAnyIP(443, listenOptions =>
        {
            listenOptions.UseHttps(httpsCertificate);
        });
    }
});

// ================================================================
// DEPENDENCY INJECTION
// ================================================================

builder.Services.AddHttpClient();

builder.Services.AddSingleton<GraphAuthService>();
builder.Services.AddSingleton<SpeechRecognitionService>();
builder.Services.AddSingleton<AudioHandler>();
builder.Services.AddSingleton<MediaSessionService>();
builder.Services.AddSingleton<MeetingMediaHandler>();

var callbackUri = builder.Configuration["Bot:CallbackUri"]
    ?? "https://teammate-bot.westus3.cloudapp.azure.com/api/calling";

var app = builder.Build();

// ================================================================
// INITIALIZE GRAPH MEDIA PLATFORM
// ================================================================

app.Services
    .GetRequiredService<MediaSessionService>()
    .Initialize();

// ================================================================
// HEALTH CHECK
// ================================================================

app.MapGet("/", () => Results.Ok(new
{
    Application = "Agent Team Mate",
    Status = "Running",
    Time = DateTime.UtcNow
}));

// ================================================================
// GRAPH AUTH TEST
// ================================================================

app.MapGet("/auth-test", async (GraphAuthService graphAuth) =>
{
    try
    {
        await graphAuth.GetAccessTokenAsync();

        return Results.Ok(new
        {
            Message = "Microsoft Graph authentication successful"
        });
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new
        {
            Message = "Authentication failed",
            Error = ex.Message
        });
    }
});

// ================================================================
// MICROSOFT GRAPH CALLING CALLBACK
// ================================================================

app.MapPost(
    "/api/calling",
    async (
        HttpRequest request,
        MeetingMediaHandler mediaHandler) =>
    {
        try
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine("       TEAMS EVENT RECEIVED");
            Console.WriteLine("================================================");
            Console.WriteLine(DateTime.UtcNow);

            var response =
                await mediaHandler.ProcessNotificationAsync(request);

            var content = response.Content == null
                ? string.Empty
                : await response.Content.ReadAsStringAsync();

            return Results.Content(
                content,
                response.Content?.Headers.ContentType?.MediaType
                    ?? "application/json",
                statusCode: (int)response.StatusCode);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"CALLBACK ERROR : {ex}");
            
            // Graph expects callback acknowledgement.
            return Results.Ok();
        }
    });

// ================================================================
// JOIN TEAMS MEETING
// ================================================================

app.MapPost(
    "/api/join",
    async (
        JoinRequest request,
        MediaSessionService mediaSessionService) =>
    {
        try
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine("       JOIN MEETING REQUEST RECEIVED");
            Console.WriteLine("================================================");

            Console.WriteLine($"Meeting ID : {request.MeetingId}");

            var call = await mediaSessionService.JoinMeetingAsync(
                request.MeetingId,
                request.Passcode);

            Console.WriteLine();
            Console.WriteLine("JOIN REQUEST SENT TO MICROSOFT GRAPH");
            Console.WriteLine($"Call ID : {call.Id}");
            Console.WriteLine($"State   : {call.Resource?.State}");

            return Results.Ok(new
            {
                CallId = call.Id,
                State = call.Resource?.State?.ToString(),
                Media = "AppHostedMediaConfig"
            });
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine("                 JOIN ERROR");
            Console.WriteLine("================================================");

            Console.WriteLine(ex);

            return Results.BadRequest(new
            {
                Message =
                    "Failed to join meeting with app-hosted media",

                Error = ex.Message
            });
        }
    });

// ================================================================
// START APPLICATION
// ================================================================

Console.WriteLine();
Console.WriteLine("================================================");
Console.WriteLine(" Agent Team Mate Bot Ready");
Console.WriteLine($" Callback : {callbackUri}");
Console.WriteLine("================================================");

app.Run();

// ================================================================
// REQUEST MODELS
// ================================================================

public record JoinRequest(
    string MeetingId,
    string? Passcode
);
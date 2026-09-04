using System.Security.Cryptography.X509Certificates;
using AgentTeamMateBot.Media;
using AgentTeamMateBot.Services;

var builder = WebApplication.CreateBuilder(args);

// Machine-local secrets/overrides. Keep real keys in appsettings.Local.json
// on the VM (gitignored). Survives PowerShell restarts.
builder.Configuration.AddJsonFile(
    "appsettings.Local.json",
    optional: true,
    reloadOnChange: true);

// User secrets are Development-only by default. Load them whenever present
// so `dotnet run` works even if ASPNETCORE_ENVIRONMENT is Production.
builder.Configuration.AddUserSecrets(typeof(Program).Assembly, optional: true);
BotLog.Configure(builder.Configuration);

Console.WriteLine($"Environment : {builder.Environment.EnvironmentName}");
Console.WriteLine($"ClientId : {builder.Configuration["Bot:ClientId"] ?? builder.Configuration["ClientId"]}");
Console.WriteLine($"TenantId : {builder.Configuration["Bot:TenantId"] ?? builder.Configuration["TenantId"]}");

Console.WriteLine();
Console.WriteLine("================================================");
Console.WriteLine("        AGENT TEAM MATE BOT STARTING");
Console.WriteLine("================================================");


// ================================================================
// KESTREL - LOCAL HTTP + PUBLIC HTTPS
// ================================================================

var serviceFqdn =
    builder.Configuration["Bot:ServiceFqdn"]
    ?? "teammate-bot.westus3.cloudapp.azure.com";

X509Certificate2? httpsCertificate = null;

using (var store =
       new X509Store(
           StoreName.My,
           StoreLocation.LocalMachine))
{
    store.Open(OpenFlags.ReadOnly);

    var certificates =
        store.Certificates.Find(
            X509FindType.FindBySubjectName,
            serviceFqdn,
            validOnly: true);

    if (certificates.Count > 0)
    {
        httpsCertificate =
            certificates[0];

        Console.WriteLine(
            $"HTTPS certificate loaded : {httpsCertificate.Subject}");

        Console.WriteLine(
            $"HTTPS thumbprint         : {httpsCertificate.Thumbprint}");
    }
    else
    {
        Console.WriteLine(
            $"HTTPS certificate not found for {serviceFqdn}");
    }
}

builder.WebHost.ConfigureKestrel(options =>
{
    // Local endpoint
    options.ListenLocalhost(5000);

    // Public HTTPS endpoint
    if (httpsCertificate != null)
    {
        options.ListenAnyIP(
            443,
            listenOptions =>
            {
                listenOptions.UseHttps(
                    httpsCertificate);
            });
    }
});

// ================================================================
// DEPENDENCY INJECTION
// ================================================================

builder.Services.AddHttpClient();

builder.Services.AddSingleton<SpeechSynthesisService>();
builder.Services.AddSingleton<MeetingContextService>();
builder.Services.AddSingleton<AiResponseService>();
builder.Services.AddSingleton<MeetingExportService>();
builder.Services.AddSingleton<IBotMediaLogger, BotMediaLogger>();
builder.Services.AddSingleton<GraphAuthService>();
builder.Services.AddSingleton<SpeechRecognitionService>();
builder.Services.AddSingleton<AudioHandler>();
builder.Services.AddSingleton<MediaSessionService>();
builder.Services.AddSingleton<AppHostedMediaService>();
builder.Services.AddSingleton<MeetingMediaHandler>();

var callbackUri =
    builder.Configuration["Bot:CallbackUri"]
    ?? "https://teammate-bot.westus3.cloudapp.azure.com/api/calling";

var app = builder.Build();

// ================================================================
// TEMP AUDIO DIRECTORY
// ================================================================

var audioDirectory =
    Path.Combine(
        app.Environment.ContentRootPath,
        "TempAudio");

Directory.CreateDirectory(
    audioDirectory);

Console.WriteLine(
    $"Temp audio directory : {audioDirectory}");

// ================================================================
// INITIALIZE GRAPH COMMUNICATIONS CLIENT
// ================================================================

_ = app.Services
    .GetRequiredService<SpeechSynthesisService>()
    .WarmupAsync();

app.Services
    .GetRequiredService<MediaSessionService>()
    .Initialize();

var appHostedMedia =
    app.Services.GetRequiredService<AppHostedMediaService>();
appHostedMedia.TryInitialize();

// ================================================================
// HEALTH CHECK
// ================================================================

app.MapGet(
    "/",
    (AppHostedMediaService appHosted) =>
        Results.Ok(
            new
            {
                Application = "Agent Team Mate",
                Status = "Running",
                Time = DateTime.UtcNow,
                MediaPlatform = appHosted.IsInitialized ? "ready" : "failed",
                MediaPlatformError = appHosted.InitError
            }));

// ================================================================
// GRAPH AUTH TEST
// ================================================================

app.MapGet(
    "/auth-test",
    async (
        GraphAuthService graphAuth) =>
    {
        try
        {
            await graphAuth
                .GetAccessTokenAsync();

            return Results.Ok(
                new
                {
                    Message =
                        "Microsoft Graph authentication successful"
                });
        }
        catch (Exception ex)
        {
            return Results.BadRequest(
                new
                {
                    Message =
                        "Authentication failed",

                    Error =
                        ex.Message
                });
        }
    });

// ================================================================
// TEMPORARY AUDIO ENDPOINT FOR GRAPH playPrompt
// ================================================================

app.MapGet(
    "/api/audio/{fileName}",
    (
        string fileName) =>
    {
        try
        {
            if (string.IsNullOrWhiteSpace(
                    fileName))
            {
                return Results.BadRequest(
                    "File name is required.");
            }

            var safeFileName =
                Path.GetFileName(
                    fileName);

            if (!safeFileName.EndsWith(
                    ".wav",
                    StringComparison.OrdinalIgnoreCase))
            {
                return Results.BadRequest(
                    "Only WAV files are supported.");
            }

            var filePath =
                Path.Combine(
                    audioDirectory,
                    safeFileName);

            if (!File.Exists(
                    filePath))
            {
                return Results.NotFound();
            }

            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" AUDIO FILE REQUESTED");
            Console.WriteLine("================================================");

            Console.WriteLine(
                $"File : {safeFileName}");

            Console.WriteLine(
                "Serving WAV audio to Microsoft Graph.");

            Console.WriteLine(
                "================================================");

            return Results.File(
                filePath,
                contentType: "audio/wav",
                enableRangeProcessing: false);
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" AUDIO ENDPOINT FAILURE");
            Console.WriteLine("================================================");

            Console.WriteLine(
                ex.Message);

            return Results.StatusCode(
                StatusCodes.Status500InternalServerError);
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

            Console.WriteLine(
                DateTime.UtcNow);

            var response =
                await mediaHandler
                    .ProcessNotificationAsync(
                        request);

            var content =
                response.Content == null
                    ? string.Empty
                    : await response.Content
                        .ReadAsStringAsync();

            return Results.Content(
                content,
                response.Content?
                    .Headers
                    .ContentType?
                    .MediaType
                    ?? "application/json",
                statusCode:
                    (int)response.StatusCode);
        }
        catch (Exception ex)
        {
            Console.WriteLine(
                $"CALLBACK ERROR : {ex}");

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

            Console.WriteLine(
                $"Meeting ID : {request.MeetingId}");

            var call =
                await mediaSessionService
                    .JoinMeetingAsync(
                        request.MeetingId,
                        request.Passcode,
                        request.OrganizerUserId,
                        request.JoinWebUrl);

            Console.WriteLine();
            Console.WriteLine(
                "JOIN REQUEST SENT TO MICROSOFT GRAPH");

            Console.WriteLine(
                $"Call ID : {call.Id}");

            Console.WriteLine(
                $"State   : {call.Resource?.State}");

            return Results.Ok(
                new
                {
                    CallId =
                        call.Id,

                    State =
                        call.Resource?
                            .State?
                            .ToString(),

                    Media =
                        "ServiceHostedMediaConfig"
                });
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine("                 JOIN ERROR");
            Console.WriteLine("================================================");

            Console.WriteLine(
                ex);

            return Results.BadRequest(
                new
                {
                    Message =
                        "Failed to join meeting with service-hosted media",

                    Error =
                        ex.Message
                });
        }
    });

// ================================================================
// PHASE 2: JOIN WITH APPLICATION-HOSTED MEDIA
// ================================================================

app.MapPost(
    "/api/join-apphosted",
    async (
        JoinRequest request,
        AppHostedMediaService appHostedMediaService) =>
    {
        try
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine("  APP-HOSTED JOIN MEETING REQUEST RECEIVED");
            Console.WriteLine("================================================");

            Console.WriteLine(
                $"Meeting ID : {request.MeetingId}");

            var call =
                await appHostedMediaService
                    .JoinMeetingAsync(
                        request.MeetingId,
                        request.Passcode);

            return Results.Ok(
                new
                {
                    CallId =
                        call.Id,

                    State =
                        call.Resource?
                            .State?
                            .ToString(),

                    Media =
                        "AppHostedMediaConfig",

                    Listening =
                        "Continuous AudioSocket.AudioMediaReceived"
                });
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine("           APP-HOSTED JOIN ERROR");
            Console.WriteLine("================================================");

            Console.WriteLine(
                ex);

            return Results.BadRequest(
                new
                {
                    Message =
                        "Failed to join meeting with application-hosted media. Continuous audio is NOT proven.",

                    Error =
                        ex.Message
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
Console.WriteLine(" Media    : SERVICE HOSTED  POST /api/join");
Console.WriteLine(" Phase 2  : APP HOSTED      POST /api/join-apphosted");
Console.WriteLine("================================================");

BotLog.Info("Ready. Waiting for a meeting join.");
if (appHostedMedia.IsInitialized)
{
    BotLog.Info("App-hosted join is available: POST /api/join-apphosted");
}
else
{
    BotLog.Info(
        $"App-hosted join is BLOCKED: {appHostedMedia.InitError ?? "MediaPlatform was not initialized"}");
}

if (!BotLog.Verbose)
{
    BotLog.Info("Debug logs are off. Set Logging:Verbose to true in appsettings.json to turn them on.");
}

app.Run();

// ================================================================
// REQUEST MODELS
// ================================================================

public record JoinRequest(
    string MeetingId,
    string? Passcode,
    string? OrganizerUserId = null,
    string? JoinWebUrl = null
);
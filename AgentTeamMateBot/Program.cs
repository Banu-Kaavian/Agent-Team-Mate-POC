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

builder.Services.AddHttpClient();
builder.Services.AddSingleton<GraphAuthService>();
builder.Services.AddSingleton<SpeechRecognitionService>();
builder.Services.AddSingleton<AudioHandler>();
builder.Services.AddSingleton<MediaSessionService>();
builder.Services.AddSingleton<MeetingMediaHandler>();

var callbackUri = builder.Configuration["Bot:CallbackUri"]
    ?? "https://teammate-bot.westus3.cloudapp.azure.com/api/calling";

var app = builder.Build();

app.Services.GetRequiredService<MediaSessionService>().Initialize();

app.MapGet("/", () => Results.Ok(new
{
    Application = "Agent Team Mate",
    Status = "Running",
    Time = DateTime.UtcNow
}));

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

app.MapPost("/api/calling", async (
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

        var response = await mediaHandler.ProcessNotificationAsync(request);
        var content = response.Content == null
            ? string.Empty
            : await response.Content.ReadAsStringAsync();

        return Results.Content(
            content,
            response.Content?.Headers.ContentType?.MediaType ?? "application/json",
            statusCode: (int)response.StatusCode);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"CALLBACK ERROR : {ex.Message}");
        return Results.Ok();
    }
});

app.MapPost("/api/join", async (
    JoinRequest request,
    MediaSessionService mediaSessionService) =>
{
    try
    {
        var call = await mediaSessionService.JoinMeetingAsync(
            request.MeetingId,
            request.Passcode);

        return Results.Ok(new
        {
            CallId = call.Id,
            State = call.Resource?.State?.ToString(),
            Media = "AppHostedMediaConfig"
        });
    }
    catch (Exception ex)
    {
        Console.WriteLine($"JOIN ERROR : {ex.Message}");

        return Results.BadRequest(new
        {
            Message = "Failed to join meeting with app-hosted media",
            Error = ex.Message
        });
    }
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

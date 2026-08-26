using Microsoft.Identity.Client;

var builder = WebApplication.CreateBuilder(args);

var clientId = builder.Configuration["Bot:ClientId"]
    ?? throw new Exception("Bot:ClientId missing");

var tenantId = builder.Configuration["Bot:TenantId"]
    ?? throw new Exception("Bot:TenantId missing");

var clientSecret = builder.Configuration["Bot:ClientSecret"]
    ?? throw new Exception("Bot:ClientSecret missing");

var confidentialClient = ConfidentialClientApplicationBuilder
    .Create(clientId)
    .WithClientSecret(clientSecret)
    .WithAuthority($"https://login.microsoftonline.com/{tenantId}")
    .Build();

var app = builder.Build();

app.MapGet("/", () => "Agent Team Mate is running");

app.MapGet("/auth-test", async () =>
{
    try
    {
        var result = await confidentialClient
            .AcquireTokenForClient(
                new[] { "https://graph.microsoft.com/.default" })
            .ExecuteAsync();

        return Results.Ok(new
        {
            message = "Microsoft Graph authentication successful",
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

app.MapPost("/api/calling", () =>
{
    return Results.Ok();
});

app.Run();
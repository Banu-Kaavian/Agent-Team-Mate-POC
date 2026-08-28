using Microsoft.Graph.Communications.Client.Authentication;
using Microsoft.Identity.Client;

namespace AgentTeamMateBot.Services;

public class GraphAuthService
{
    private readonly IConfidentialClientApplication _client;
    private readonly string _clientId;
    private readonly string _tenantId;

    public GraphAuthService(IConfiguration configuration)
    {
        _clientId = configuration["Bot:ClientId"]
            ?? throw new Exception("Bot:ClientId missing");

        _tenantId = configuration["Bot:TenantId"]
            ?? throw new Exception("Bot:TenantId missing");

        var clientSecret = configuration["Bot:ClientSecret"]
            ?? throw new Exception("Bot:ClientSecret missing");

        _client = ConfidentialClientApplicationBuilder
            .Create(_clientId)
            .WithClientSecret(clientSecret)
            .WithAuthority($"https://login.microsoftonline.com/{_tenantId}")
            .Build();
    }

    public string ClientId => _clientId;

    public string TenantId => _tenantId;

    public IConfidentialClientApplication ConfidentialClient => _client;

    public ITokenProvider CreateTokenProvider()
    {
        return new MsalTokenProvider(_client);
    }

    public async Task<string> GetAccessTokenAsync()
    {
        try
        {
            var result = await _client
                .AcquireTokenForClient(new[] { "https://graph.microsoft.com/.default" })
                .ExecuteAsync();

            Console.WriteLine("[AUTH] Graph authentication successful");
            return result.AccessToken;
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" GRAPH AUTHENTICATION FAILURE");
            Console.WriteLine("================================================");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex);
            throw;
        }
    }
}

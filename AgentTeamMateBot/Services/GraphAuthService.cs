using Microsoft.Identity.Client;

namespace AgentTeamMateBot.Services;


public class GraphAuthService
{

    private readonly IConfiguration _configuration;


    private readonly IConfidentialClientApplication _client;



    public GraphAuthService(
        IConfiguration configuration)
    {

        _configuration = configuration;


        var clientId =
            _configuration["Bot:ClientId"]
            ?? throw new Exception("ClientId missing");


        var tenantId =
            _configuration["Bot:TenantId"]
            ?? throw new Exception("TenantId missing");


        var clientSecret =
            _configuration["Bot:ClientSecret"]
            ?? throw new Exception("ClientSecret missing");



        _client =
            ConfidentialClientApplicationBuilder
            .Create(clientId)
            .WithClientSecret(clientSecret)
            .WithAuthority(
                $"https://login.microsoftonline.com/{tenantId}")
            .Build();

    }




    public async Task<string> GetAccessTokenAsync()
    {

        var result =
            await _client
            .AcquireTokenForClient(
            new[]
            {
                "https://graph.microsoft.com/.default"
            })
            .ExecuteAsync();



        return result.AccessToken;

    }

}
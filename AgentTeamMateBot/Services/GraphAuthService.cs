using System.IdentityModel.Tokens.Jwt;
using System.Net.Http.Headers;
using Microsoft.Graph.Communications.Client.Authentication;
using Microsoft.Graph.Communications.Common;
using Microsoft.Graph.Communications.Common.Telemetry;
using Microsoft.Identity.Client;

namespace AgentTeamMateBot.Services;

public class GraphAuthService
{
    private readonly IConfidentialClientApplication _client;
    private readonly string _clientId;
    private readonly string _tenantId;
    private readonly string _clientSecret;

    public GraphAuthService(IConfiguration configuration)
    {
        _clientId = configuration["Bot:ClientId"]
            ?? configuration["ClientId"]
            ?? throw new Exception("Bot:ClientId missing");

        _tenantId = configuration["Bot:TenantId"]
            ?? configuration["TenantId"]
            ?? throw new Exception("Bot:TenantId missing");

        _clientSecret = configuration["Bot:ClientSecret"]
            ?? configuration["ClientSecret"]
            ?? throw new Exception("Bot:ClientSecret missing");

        _client = ConfidentialClientApplicationBuilder
            .Create(_clientId)
            .WithClientSecret(_clientSecret)
            .WithAuthority($"https://login.microsoftonline.com/{_tenantId}")
            .Build();
    }

    public string ClientId => _clientId;

    public string TenantId => _tenantId;

    public IConfidentialClientApplication ConfidentialClient => _client;

    public ITokenProvider CreateTokenProvider()
    {
        return new TenantAwareTokenProvider(_clientId, _clientSecret, _tenantId, _client);
    }

    public IRequestAuthenticationProvider CreateAuthenticationProvider(IGraphLogger logger)
    {
        var tokenProvider = CreateTokenProvider();
        var inbound = new DefaultAuthenticationProvider(_clientId, tokenProvider, logger);
        return new TenantAwareAuthenticationProvider(_tenantId, tokenProvider, inbound);
    }

    public async Task<string> GetAccessTokenAsync()
    {
        try
        {
            var result = await _client
                .AcquireTokenForClient(new[] { "https://graph.microsoft.com/.default" })
                .ExecuteAsync();

            LogTokenTenant(result.AccessToken);
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

    internal static string ResolveTenant(string? requestedTenant, string homeTenant)
    {
        if (Guid.TryParse(requestedTenant, out var tenantGuid) && tenantGuid != Guid.Empty)
        {
            return tenantGuid.ToString();
        }

        return homeTenant;
    }

    internal static void LogTokenTenant(string accessToken)
    {
        try
        {
            var jwt = new JwtSecurityTokenHandler().ReadJwtToken(accessToken);
            var tid = jwt.Claims.FirstOrDefault(c => c.Type == "tid")?.Value;
            var appid = jwt.Claims.FirstOrDefault(c => c.Type == "appid")?.Value
                ?? jwt.Claims.FirstOrDefault(c => c.Type == "azp")?.Value;
            var roles = string.Join(",", jwt.Claims.Where(c => c.Type == "roles").Select(c => c.Value));

            Console.WriteLine($"[AUTH] Token tid={tid} appid={appid}");
            if (!string.IsNullOrWhiteSpace(roles))
            {
                Console.WriteLine($"[AUTH] Token roles={roles}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[AUTH] Could not decode token tenant: {ex.Message}");
        }
    }
}

internal sealed class TenantAwareTokenProvider : ITokenProvider
{
    private static readonly string[] Scopes = { "https://graph.microsoft.com/.default" };

    private readonly string _clientId;
    private readonly string _clientSecret;
    private readonly string _homeTenantId;
    private readonly IConfidentialClientApplication _homeClient;

    public TenantAwareTokenProvider(
        string clientId,
        string clientSecret,
        string homeTenantId,
        IConfidentialClientApplication homeClient)
    {
        _clientId = clientId;
        _clientSecret = clientSecret;
        _homeTenantId = homeTenantId;
        _homeClient = homeClient;
    }

    public async Task<string> AcquireTokenAsync(string tenant)
    {
        var tenantToUse = GraphAuthService.ResolveTenant(tenant, _homeTenantId);

        Console.WriteLine($"[AUTH] AcquireToken requested='{tenant}' using='{tenantToUse}'");

        try
        {
            var app = _homeClient;
            if (!string.Equals(tenantToUse, _homeTenantId, StringComparison.OrdinalIgnoreCase))
            {
                app = ConfidentialClientApplicationBuilder
                    .Create(_clientId)
                    .WithClientSecret(_clientSecret)
                    .WithAuthority($"https://login.microsoftonline.com/{tenantToUse}")
                    .Build();
            }

            var result = await app
                .AcquireTokenForClient(Scopes)
                .ExecuteAsync();

            GraphAuthService.LogTokenTenant(result.AccessToken);
            return result.AccessToken;
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" GRAPH AUTHENTICATION FAILURE");
            Console.WriteLine("================================================");
            Console.WriteLine($"Tenant requested : {tenant}");
            Console.WriteLine($"Tenant used      : {tenantToUse}");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex);
            throw;
        }
    }
}

internal sealed class TenantAwareAuthenticationProvider : IRequestAuthenticationProvider
{
    private readonly string _homeTenantId;
    private readonly ITokenProvider _tokenProvider;
    private readonly IRequestAuthenticationProvider _inbound;

    public TenantAwareAuthenticationProvider(
        string homeTenantId,
        ITokenProvider tokenProvider,
        IRequestAuthenticationProvider inbound)
    {
        _homeTenantId = homeTenantId;
        _tokenProvider = tokenProvider;
        _inbound = inbound;
    }

    public async Task AuthenticateOutboundRequestAsync(HttpRequestMessage request, string tenant)
    {
        var tenantToUse = GraphAuthService.ResolveTenant(tenant, _homeTenantId);
        Console.WriteLine($"[AUTH] Outbound {request.Method} {request.RequestUri}");
        Console.WriteLine($"[AUTH] Outbound tenant arg='{tenant}' using='{tenantToUse}'");

        var token = await _tokenProvider.AcquireTokenAsync(tenantToUse).ConfigureAwait(false);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

        if (request.Headers.Contains(HttpConstants.HeaderNames.Tenant))
        {
            request.Headers.Remove(HttpConstants.HeaderNames.Tenant);
        }

        request.Headers.Add(HttpConstants.HeaderNames.Tenant, tenantToUse);
#pragma warning disable CS0618
        request.Properties[HttpConstants.HeaderNames.Tenant] = tenantToUse;
#pragma warning restore CS0618
    }

    public Task<RequestValidationResult> ValidateInboundRequestAsync(HttpRequestMessage request)
    {
        return _inbound.ValidateInboundRequestAsync(request);
    }
}

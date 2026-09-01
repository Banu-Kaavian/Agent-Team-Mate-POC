using System.Text;
using System.Text.Json;

namespace AgentTeamMateBot.Services;

public class AiResponseService
{
    private readonly IConfiguration _configuration;
    private readonly HttpClient _httpClient;

    public AiResponseService(IConfiguration configuration)
    {
        _configuration = configuration;
        _httpClient = new HttpClient();
    }

    public async Task<string?> GetResponseAsync(string userMessage)
    {
        if (string.IsNullOrWhiteSpace(userMessage))
            return null;

        var endpoint = _configuration["AzureOpenAI:Endpoint"];
        var deployment = _configuration["AzureOpenAI:Deployment"];
        var apiKey = _configuration["AzureOpenAI:ApiKey"];

        if (string.IsNullOrWhiteSpace(endpoint))
            throw new Exception("AzureOpenAI:Endpoint missing");

        if (string.IsNullOrWhiteSpace(deployment))
            throw new Exception("AzureOpenAI:Deployment missing");

        if (string.IsNullOrWhiteSpace(apiKey))
            throw new Exception("AzureOpenAI:ApiKey missing");

        var url =
            $"{endpoint.TrimEnd('/')}/openai/v1/responses";

        var payload = new
        {
            model = deployment,

            instructions =
                "You are Agent Team Mate, an AI participant in a Microsoft Teams meeting. " +
                "Respond naturally and briefly because your response will be spoken aloud.",

            input = userMessage
        };

        var json =
            JsonSerializer.Serialize(payload);

        using var request =
            new HttpRequestMessage(
                HttpMethod.Post,
                url);

        request.Headers.Add(
            "api-key",
            apiKey);

        request.Content =
            new StringContent(
                json,
                Encoding.UTF8,
                "application/json");

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" SENDING TO AZURE OPENAI");
        Console.WriteLine("================================================");
        Console.WriteLine(userMessage);
        Console.WriteLine("================================================");

        using var response =
            await _httpClient.SendAsync(request);

        var body =
            await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" AZURE OPENAI FAILURE");
            Console.WriteLine("================================================");
            Console.WriteLine($"Status : {(int)response.StatusCode}");
            Console.WriteLine(body);
            Console.WriteLine("================================================");

            return null;
        }

        using var document =
            JsonDocument.Parse(body);

        if (!document.RootElement.TryGetProperty(
                "output",
                out var output))
        {
            return null;
        }

        foreach (var item in output.EnumerateArray())
        {
            if (!item.TryGetProperty(
                    "content",
                    out var content))
                continue;

            foreach (var part in content.EnumerateArray())
            {
                if (!part.TryGetProperty(
                        "text",
                        out var textElement))
                    continue;

                var answer =
                    textElement.GetString();

                if (string.IsNullOrWhiteSpace(answer))
                    continue;

                Console.WriteLine();
                Console.WriteLine("================================================");
                Console.WriteLine(" AI RESPONSE");
                Console.WriteLine("================================================");
                Console.WriteLine(answer);
                Console.WriteLine("================================================");

                return answer;
            }
        }

        return null;
    }
}
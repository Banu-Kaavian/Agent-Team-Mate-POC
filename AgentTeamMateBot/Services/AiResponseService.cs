using System.Collections.Concurrent;
using System.Text;
using System.Text.Json;

namespace AgentTeamMateBot.Services;

public class AiResponseService
{
    private const int MaxMessagesPerCall = 20;

    private readonly IConfiguration _configuration;
    private readonly HttpClient _httpClient;
    private readonly ConcurrentDictionary<string, CallConversation> _conversations = new();

    public AiResponseService(IConfiguration configuration)
    {
        _configuration = configuration;
        _httpClient = new HttpClient();
    }

    public async Task<string?> GetResponseAsync(
        string callId,
        string userMessage)
    {
        if (string.IsNullOrWhiteSpace(userMessage))
            return null;

        if (string.IsNullOrWhiteSpace(callId))
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

        var conversation =
            _conversations.GetOrAdd(
                callId,
                _ => new CallConversation());

        var messageCount =
            conversation.Add(
                "user",
                userMessage);

        LogConversationMemory(
            callId,
            messageCount);

        var history =
            conversation.Snapshot();

        var input =
            history
                .Select(message => new
                {
                    role = message.Role,
                    content = message.Content
                })
                .ToArray();

        var url =
            $"{endpoint.TrimEnd('/')}/openai/v1/responses";

        var payload = new
        {
            model = deployment,

            instructions =
                "You are Agent Team Mate, an AI participant in a Microsoft Teams meeting. " +
                "Respond naturally and briefly because your response will be spoken aloud. " +
                "Previous messages are the ongoing conversation in this same Teams meeting. " +
                "Use that context when answering.",

            input
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

                conversation.Add(
                    "assistant",
                    answer);

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

    public void ClearConversation(
        string callId)
    {
        if (string.IsNullOrWhiteSpace(callId))
            return;

        if (_conversations.TryRemove(
                callId,
                out _))
        {
            Console.WriteLine();
            Console.WriteLine("================================================");
            Console.WriteLine(" CONVERSATION MEMORY CLEARED");
            Console.WriteLine("================================================");
            Console.WriteLine($"Call ID : {callId}");
            Console.WriteLine("================================================");
        }
    }

    private static void LogConversationMemory(
        string callId,
        int messageCount)
    {
        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" CONVERSATION MEMORY");
        Console.WriteLine("================================================");
        Console.WriteLine($"Call ID : {callId}");
        Console.WriteLine($"Messages in memory : {messageCount}");
        Console.WriteLine("================================================");
    }

    private sealed class CallConversation
    {
        private readonly object _lock = new();
        private readonly List<ConversationMessage> _messages = new();

        public int Add(
            string role,
            string content)
        {
            lock (_lock)
            {
                _messages.Add(
                    new ConversationMessage(
                        role,
                        content));

                while (_messages.Count > MaxMessagesPerCall)
                {
                    _messages.RemoveAt(0);
                }

                return _messages.Count;
            }
        }

        public List<ConversationMessage> Snapshot()
        {
            lock (_lock)
            {
                return _messages.ToList();
            }
        }
    }

    private sealed record ConversationMessage(
        string Role,
        string Content);
}

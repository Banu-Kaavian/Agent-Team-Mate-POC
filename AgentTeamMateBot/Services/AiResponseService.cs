using System.Collections.Concurrent;
using System.Text;
using System.Text.Json;

namespace AgentTeamMateBot.Services;

public class AiResponseService
{
    private const int MaxMessagesPerCall = 20;

    private readonly IConfiguration _configuration;
    private readonly MeetingContextService _meetingContextService;
    private readonly HttpClient _httpClient;
    private readonly HttpClient _documentHttpClient;
    private readonly ConcurrentDictionary<string, CallConversation> _conversations = new();

    public AiResponseService(
        IConfiguration configuration,
        MeetingContextService meetingContextService)
    {
        _configuration = configuration;
        _meetingContextService = meetingContextService;
        _httpClient = new HttpClient();
        _httpClient.Timeout = TimeSpan.FromSeconds(20);
        _documentHttpClient = new HttpClient();
        _documentHttpClient.Timeout = TimeSpan.FromSeconds(90);
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
        var deployment = _configuration["AzureOpenAI:Deployment"]
            ?? _configuration["OPENAI_MODEL"];
        var apiKey = _configuration["AzureOpenAI:ApiKey"]
            ?? _configuration["OPENAI_API_KEY"];

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

        // Skip Graph transcript lookup on the spoken path. It can take minutes
        // and blocks Azure OpenAI. Live transcript from this call is enough.
        var liveTranscript =
            _meetingContextService.GetLiveTranscript(
                callId);

        var apiVersion =
            _configuration["AzureOpenAI:ApiVersion"]
            ?? "2025-01-01-preview";

        var url =
            $"{endpoint.TrimEnd('/')}/openai/deployments/{deployment}/chat/completions?api-version={apiVersion}";

        var messages =
            BuildMessages(
                history,
                meetingContext: null,
                liveTranscript,
                userMessage);

        // gpt-5-mini spends tokens on hidden reasoning. A small
        // max_completion_tokens budget often finishes with empty content.
        var payload = new
        {
            messages,
            max_completion_tokens = 1024,
            reasoning_effort = "minimal"
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

        HttpResponseMessage response;
        try
        {
            response =
                await _httpClient.SendAsync(request);
        }
        catch (Exception ex)
        {
            BotLog.Info($"Error: Azure OpenAI request failed. {ex.Message}");
            throw;
        }

        using (response)
        {
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

            BotLog.Info(
                $"Error: Azure OpenAI {(int)response.StatusCode}. {TrimForLog(body)}");

            return null;
        }

        using var document =
            JsonDocument.Parse(body);

        if (!document.RootElement.TryGetProperty(
                "choices",
                out var choices) ||
            choices.ValueKind != JsonValueKind.Array)
        {
            BotLog.Info("Error: Azure OpenAI response had no choices.");
            return null;
        }

        foreach (var choice in choices.EnumerateArray())
        {
            var answer = ExtractMessageText(choice);

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

        BotLog.Info(
            $"Error: Azure OpenAI returned no text. {DescribeEmptyResponse(document.RootElement)}");
        return null;
        }
    }

    public async Task<string?> GenerateMeetingDocumentAsync(
        string callId,
        string liveTranscript)
    {
        if (string.IsNullOrWhiteSpace(liveTranscript))
            return null;

        var endpoint = _configuration["AzureOpenAI:Endpoint"];
        var deployment = _configuration["AzureOpenAI:Deployment"]
            ?? _configuration["OPENAI_MODEL"];
        var apiKey = _configuration["AzureOpenAI:ApiKey"]
            ?? _configuration["OPENAI_API_KEY"];

        if (string.IsNullOrWhiteSpace(endpoint) ||
            string.IsNullOrWhiteSpace(deployment) ||
            string.IsNullOrWhiteSpace(apiKey))
        {
            return liveTranscript;
        }

        var apiVersion =
            _configuration["AzureOpenAI:ApiVersion"]
            ?? "2025-01-01-preview";

        var url =
            $"{endpoint.TrimEnd('/')}/openai/deployments/{deployment}/chat/completions?api-version={apiVersion}";

        var now = DateTime.Now;
        var system =
            "You write a meeting handoff document for a product workflow. " +
            "Output plain text only. No markdown headings with hashes, no bullet symbols if you can use numbered lines. " +
            "First write a PRD-style summary: problem, goals, requirements, decisions, open questions, and action items. " +
            "Then write a Transcript section as speaker lines exactly like: Name: sentence. " +
            "Use names when the notes include them. Otherwise use Participant. " +
            "Do not invent facts that are not in the notes. " +
            $"Local date and time: {now:dddd, MMMM d, yyyy} at {now:h:mm tt}.";

        var payload = new
        {
            messages = new object[]
            {
                new { role = "system", content = system },
                new
                {
                    role = "user",
                    content =
                        "Create the full meeting summary and transcript from these live notes:\n\n" +
                        liveTranscript
                }
            },
            max_completion_tokens = 4096,
            reasoning_effort = "minimal"
        };

        using var request = new HttpRequestMessage(HttpMethod.Post, url);
        request.Headers.Add("api-key", apiKey);
        request.Content = new StringContent(
            JsonSerializer.Serialize(payload),
            Encoding.UTF8,
            "application/json");

        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" GENERATING MEETING PRD / TRANSCRIPT");
        Console.WriteLine("================================================");
        Console.WriteLine($"Call ID : {callId}");

        HttpResponseMessage response;
        try
        {
            response = await _documentHttpClient.SendAsync(request);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[MEETING DOCUMENT] {ex.Message}");
            return liveTranscript;
        }

        using (response)
        {
            var body = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode)
            {
                Console.WriteLine($"[MEETING DOCUMENT] Azure OpenAI {(int)response.StatusCode}");
                return liveTranscript;
            }

            using var document = JsonDocument.Parse(body);
            if (!document.RootElement.TryGetProperty("choices", out var choices) ||
                choices.ValueKind != JsonValueKind.Array)
            {
                return liveTranscript;
            }

            foreach (var choice in choices.EnumerateArray())
            {
                var text = ExtractMessageText(choice);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    Console.WriteLine($"Generated {text.Length} characters.");
                    Console.WriteLine("================================================");
                    return text;
                }
            }

            Console.WriteLine("[MEETING DOCUMENT] Empty model output. Using live transcript.");
            Console.WriteLine("================================================");
            return liveTranscript;
        }
    }

    private static string? ExtractMessageText(JsonElement choice)
    {
        if (!choice.TryGetProperty("message", out var message))
            return null;

        if (message.TryGetProperty("content", out var content))
        {
            var text = ReadContent(content);
            if (!string.IsNullOrWhiteSpace(text))
                return text;
        }

        if (message.TryGetProperty("refusal", out var refusal))
        {
            var text = refusal.GetString();
            if (!string.IsNullOrWhiteSpace(text))
                return text;
        }

        return null;
    }

    private static string? ReadContent(JsonElement content)
    {
        if (content.ValueKind == JsonValueKind.String)
            return content.GetString();

        if (content.ValueKind != JsonValueKind.Array)
            return null;

        var builder = new StringBuilder();
        foreach (var part in content.EnumerateArray())
        {
            if (part.ValueKind == JsonValueKind.String)
            {
                builder.Append(part.GetString());
                continue;
            }

            if (part.TryGetProperty("text", out var text))
                builder.Append(text.GetString());
        }

        var combined = builder.ToString().Trim();
        return string.IsNullOrWhiteSpace(combined) ? null : combined;
    }

    private static string DescribeEmptyResponse(JsonElement root)
    {
        var finishReason = "unknown";
        if (root.TryGetProperty("choices", out var choices) &&
            choices.ValueKind == JsonValueKind.Array &&
            choices.GetArrayLength() > 0 &&
            choices[0].TryGetProperty("finish_reason", out var reason))
        {
            finishReason = reason.GetString() ?? finishReason;
        }

        var completionTokens = "?";
        var reasoningTokens = "?";
        if (root.TryGetProperty("usage", out var usage))
        {
            if (usage.TryGetProperty("completion_tokens", out var completion))
                completionTokens = completion.ToString();

            if (usage.TryGetProperty("completion_tokens_details", out var details) &&
                details.TryGetProperty("reasoning_tokens", out var reasoning))
            {
                reasoningTokens = reasoning.ToString();
            }
        }

        return $"finish_reason={finishReason}, completion_tokens={completionTokens}, reasoning_tokens={reasoningTokens}.";
    }

    private static string TrimForLog(string body)
    {
        if (string.IsNullOrWhiteSpace(body))
        {
            return "(empty body)";
        }

        var trimmed = body.Replace('\n', ' ').Trim();
        return trimmed.Length <= 240
            ? trimmed
            : trimmed[..240];
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

    private async Task<string?> GetMeetingContextSafelyAsync(
        string callId)
    {
        try
        {
            return await _meetingContextService
                .GetMeetingContextAsync(
                    callId);
        }
        catch (Exception ex)
        {
            Console.WriteLine(
                $"[MEETING CONTEXT] AI lookup failed: {ex.Message}");
            Console.WriteLine(
                "[MEETING CONTEXT] Transcript not available. Continuing without transcript.");
            return null;
        }
    }

    private static object[] BuildMessages(
        IReadOnlyList<ConversationMessage> history,
        string? meetingContext,
        string? liveTranscript,
        string currentQuestion)
    {
        var messages =
            new List<object>
            {
                new
                {
                    role = "system",
                    content =
                        BuildInstructions(
                            meetingContext,
                            liveTranscript,
                            currentQuestion)
                }
            };

        foreach (var message in history)
        {
            messages.Add(
                new
                {
                    role = message.Role,
                    content = message.Content
                });
        }

        return messages.ToArray();
    }

    private static string BuildInstructions(
        string? meetingContext,
        string? liveTranscript,
        string currentQuestion)
    {
        var builder = new StringBuilder();
        var now = DateTime.Now;

        builder.Append(
            "You are Agent Nova in a live Microsoft Teams meeting. ");
        builder.Append(
            "Answer in one or two short spoken sentences. No markdown, lists, URLs, or symbols. ");
        builder.Append(
            "Start with the answer. Do not invent meeting facts. ");
        builder.Append(
            $"The current local date and time is {now:dddd, MMMM d, yyyy} at {now:h:mm tt}.");

        builder.AppendLine();
        builder.AppendLine();
        builder.Append("Meeting transcript: ");
        builder.AppendLine(
            string.IsNullOrWhiteSpace(liveTranscript)
                ? "None yet."
                : liveTranscript);

        if (!string.IsNullOrWhiteSpace(meetingContext))
        {
            builder.AppendLine();
            builder.Append("Graph transcript: ");
            builder.AppendLine(meetingContext);
        }

        builder.AppendLine();
        builder.Append("Current question: ");
        builder.Append(currentQuestion);

        return builder.ToString();
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

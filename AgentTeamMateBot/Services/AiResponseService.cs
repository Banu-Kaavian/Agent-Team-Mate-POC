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
        _httpClient.Timeout = TimeSpan.FromSeconds(45);
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
            max_completion_tokens = 400,
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

        var system =
            "You clean a live speech-to-text meeting log for a workflow API. " +
            "Output only speaker lines, one per line, exactly like: Name: sentence. " +
            "Fix obvious speech-recognition spelling on Participant lines only " +
            "(names, SAP products, tools) without changing meaning. " +
            "Examples: Agent Novak or Page and Nova become Agent Nova; S4 HANA becomes S/4HANA; " +
            "duplicate words like BODS SAP BODS become SAP BODS. " +
            "Use real people names from greetings when they appear; otherwise use Participant. " +
            "Agent Nova lines are already her full spoken answers, including APIs, BAPIs, OData, RFC, " +
            "table names, field mappings, and every technical step. Copy each Agent Nova: line " +
            "verbatim, word for word. Do not summarize, shorten, paraphrase, merge, or drop " +
            "any technical detail from Nova. " +
            "Do not replace Nova's answers with the leave or summarize command. " +
            "Do not invent facts, tools, or decisions that are not in the notes. " +
            "No title, markdown, bullets, or extra sections.";

        var payload = new
        {
            messages = new object[]
            {
                new { role = "system", content = system },
                new
                {
                    role = "user",
                    content =
                        "Write the cleaned transcript as Name: sentence lines. " +
                        "Clean Participant speech. Paste every Agent Nova: line unchanged and in full, " +
                        "including all technical content. This text is the transcript JSON field. Live notes:\n\n" +
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
                    var preserved = PreserveAgentNovaAnswers(liveTranscript, text);
                    Console.WriteLine($"Generated {preserved.Length} characters.");
                    Console.WriteLine("================================================");
                    return preserved;
                }
            }

            Console.WriteLine("[MEETING DOCUMENT] Empty model output. Using live transcript.");
            Console.WriteLine("================================================");
            return liveTranscript;
        }
    }

    public async Task<string> GenerateSpokenRecapAsync(string callId)
    {
        var liveTranscript = _meetingContextService.GetLiveTranscript(callId);
        if (string.IsNullOrWhiteSpace(liveTranscript))
        {
            return "I do not have enough of the meeting yet to recap.";
        }

        var endpoint = _configuration["AzureOpenAI:Endpoint"];
        var deployment = _configuration["AzureOpenAI:Deployment"]
            ?? _configuration["OPENAI_MODEL"];
        var apiKey = _configuration["AzureOpenAI:ApiKey"]
            ?? _configuration["OPENAI_API_KEY"];

        if (string.IsNullOrWhiteSpace(endpoint) ||
            string.IsNullOrWhiteSpace(deployment) ||
            string.IsNullOrWhiteSpace(apiKey))
        {
            return "I cannot recap yet. Azure OpenAI is not configured.";
        }

        var apiVersion =
            _configuration["AzureOpenAI:ApiVersion"]
            ?? "2025-01-01-preview";

        var url =
            $"{endpoint.TrimEnd('/')}/openai/deployments/{deployment}/chat/completions?api-version={apiVersion}";

        var system =
            "You are Agent Nova recapping this Teams meeting out loud. " +
            "Give a simple spoken recap only. Do not offer to send a workflow. " +
            "Focus on the technical points: products, versions, APIs, BAPIs, OData, RFC, tools, " +
            "objects, and decisions that were actually discussed. " +
            "Keep Agent Nova's technical recommendations. " +
            "Use four to eight short sentences. No markdown, bullets, or URLs. " +
            "Do not invent facts.";

        var payload = new
        {
            messages = new object[]
            {
                new { role = "system", content = system },
                new
                {
                    role = "user",
                    content = "Recap this meeting for the people on the call:\n\n" + liveTranscript
                }
            },
            max_completion_tokens = 700,
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
        Console.WriteLine(" GENERATING SPOKEN MEETING RECAP");
        Console.WriteLine("================================================");

        try
        {
            using var response = await _documentHttpClient.SendAsync(request);
            var body = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode)
            {
                Console.WriteLine($"[MEETING RECAP] Azure OpenAI {(int)response.StatusCode}");
                return "I could not build the recap. Please try again.";
            }

            using var document = JsonDocument.Parse(body);
            if (!document.RootElement.TryGetProperty("choices", out var choices) ||
                choices.ValueKind != JsonValueKind.Array)
            {
                return "I could not build the recap. Please try again.";
            }

            foreach (var choice in choices.EnumerateArray())
            {
                var text = ExtractMessageText(choice);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    return text.Trim();
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[MEETING RECAP] {ex.Message}");
        }

        return "I could not build the recap. Please try again.";
    }

    private static string PreserveAgentNovaAnswers(string liveNotes, string cleanedTranscript)
    {
        var originalNova = liveNotes
            .Split(new[] { "\r\n", "\n" }, StringSplitOptions.None)
            .Select(line => line.Trim())
            .Where(line => line.StartsWith("Agent Nova:", StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (originalNova.Count == 0)
        {
            return cleanedTranscript;
        }

        var outputLines = cleanedTranscript
            .Split(new[] { "\r\n", "\n" }, StringSplitOptions.None)
            .ToList();

        var novaIndex = 0;
        for (var i = 0; i < outputLines.Count; i++)
        {
            if (!outputLines[i].TrimStart().StartsWith("Agent Nova:", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (novaIndex < originalNova.Count)
            {
                outputLines[i] = originalNova[novaIndex];
                novaIndex++;
            }
        }

        while (novaIndex < originalNova.Count)
        {
            outputLines.Add(originalNova[novaIndex]);
            novaIndex++;
        }

        return string.Join(Environment.NewLine, outputLines.Where(line => !string.IsNullOrWhiteSpace(line)));
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
            "You are Agent Nova, a teammate in this Microsoft Teams meeting with SAP technical and functional experience. ");
        builder.Append(
            "Talk like a colleague on the call, not like a project manager or a help desk script. ");
        builder.Append(
            "Give a direct answer in two or three short spoken sentences. Never more than four sentences. ");
        builder.Append(
            "Then ask one simple teammate question when it helps, such as what they already tried, which object, " +
            "which system, volume, timeline, or whether they want you to go deeper. " +
            "Ask only one question. Do not interview them. Skip the question if the transcript already answered it. ");
        builder.Append(
            "Do not say next step, next steps, or give a plan unless someone asks what to do next or asks for a plan. ");
        builder.Append(
            "Do not list categories, do not give long requirement catalogs, and do not explain every option. ");
        builder.Append(
            "No markdown, bullets, numbers, URLs, or symbols. Start with the answer. Do not invent meeting facts. ");
        builder.Append(
            "If they ask what was said earlier, what we discussed before, previous points, or who said what, " +
            "answer from the meeting transcript below. Restate those earlier points in plain speech. " +
            "If it is not in the transcript, say you do not have that yet. ");
        builder.Append(
            "If they only asked for a summary, recap only. Do not mention the workflow. ");
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

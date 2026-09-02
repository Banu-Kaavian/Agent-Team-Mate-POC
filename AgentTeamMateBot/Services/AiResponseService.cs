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
    private readonly ConcurrentDictionary<string, CallConversation> _conversations = new();

    public AiResponseService(
        IConfiguration configuration,
        MeetingContextService meetingContextService)
    {
        _configuration = configuration;
        _meetingContextService = meetingContextService;
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

        var meetingContext =
            await GetMeetingContextSafelyAsync(
                callId);

        var liveTranscript =
            _meetingContextService.GetLiveTranscript(
                callId);

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
                BuildInstructions(
                    meetingContext,
                    liveTranscript,
                    userMessage),

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

    private static string BuildInstructions(
        string? meetingContext,
        string? liveTranscript,
        string currentQuestion)
    {
        var builder = new StringBuilder();

        builder.Append(
            "You are Agent Team Mate, a live technical teammate participating in a Microsoft Teams meeting. ");

        builder.Append(
            "Answer the user's current question directly and briefly because the response will be spoken aloud. ");

        builder.Append(
            "Do not repeatedly restate the complete plan from previous turns. ");

        builder.Append(
            "Do not repeatedly ask for confirmation after the user has already confirmed. ");

        builder.Append(
            "If the user says yes, proceed, start, go ahead, just proceed, no more questions, or equivalent, " +
            "stop asking follow-up questions unless a genuinely required piece of information prevents answering. ");

        builder.Append(
            "Do not ask \"Should I proceed?\", \"Want me to start?\", or \"Any last-minute details?\" after the user has already confirmed. ");

        builder.Append(
            "If the request has been answered, end naturally. Do not automatically add a closing question. ");

        builder.Append(
            "Never claim that work will continue after this response. ");

        builder.Append(
            "Do not say you will work on it later, deliver it later, deliver it tomorrow, " +
            "deliver it in two business days, message when ready, or start now and get back to the user. ");

        builder.Append(
            "This application has no background job that can perform future work. ");

        builder.Append(
            "Never invent delivery timelines. ");

        builder.Append(
            "Never claim to have performed an external action unless this application actually has a connected tool that performed it. ");

        builder.Append(
            "Do not claim you validated SAP documentation, retrieved a website, generated a PDF, sent a message, " +
            "or completed any other external action unless such a tool actually ran. ");

        builder.Append(
            "If the user asks for something this application cannot execute, clearly say what you can do in this conversation. ");

        builder.Append(
            "For example: you can draft a wireframe specification in this conversation, " +
            "but the current Agent Team Mate implementation does not have a PDF-generation tool connected. ");

        builder.Append(
            "Distinguish between answering or reasoning using model knowledge, " +
            "information from meeting context, and an actual external action or tool. ");

        builder.Append(
            "Never present model reasoning as an external action. ");

        builder.Append(
            "Use the supplied meeting transcript when answering questions about the discussion. ");

        builder.Append(
            "The live meeting transcript is the current in-meeting speech captured during this call. ");

        builder.Append(
            "If a Graph meeting transcript is also supplied, treat it as a separate source. ");

        builder.Append(
            "Do not claim something was discussed unless it appears in the supplied meeting transcript or Graph meeting transcript. ");

        builder.Append(
            "If the meeting context does not contain the requested information, say that clearly. ");

        builder.Append(
            "Never invent a meeting decision, participant statement, requirement, name, date, or conclusion. ");

        builder.Append(
            "When speaker names are available in the meeting transcript, preserve those names and use them when relevant. ");

        builder.Append(
            "If speaker identity is not available in the transcript, do not guess who said something. ");

        builder.Append(
            "Clearly distinguish between what was discussed in the meeting and what you are recommending from technical knowledge. ");

        builder.Append(
            "For SAP-related questions, you may use model knowledge and prefer SAP standard functionality over custom development, " +
            "but say it is based on model knowledge, not a live documentation lookup. ");

        builder.Append(
            "Recommend custom RAP, custom OData, custom ABAP, or other custom development only when standard functionality is insufficient. ");

        builder.Append(
            "Keep spoken responses concise and natural.");

        builder.AppendLine();
        builder.AppendLine();
        builder.AppendLine("Graph meeting transcript:");

        if (string.IsNullOrWhiteSpace(meetingContext))
        {
            builder.AppendLine(
                "No Graph meeting transcript is available.");
        }
        else
        {
            builder.AppendLine(meetingContext);
        }

        builder.AppendLine();
        builder.AppendLine("Meeting transcript:");

        if (string.IsNullOrWhiteSpace(liveTranscript))
        {
            builder.Append(
                "No live meeting transcript is available yet.");
        }
        else
        {
            builder.AppendLine(liveTranscript);
        }

        builder.AppendLine();
        builder.AppendLine("Current question:");
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

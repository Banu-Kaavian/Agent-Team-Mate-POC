using System.Text.RegularExpressions;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Schema;

namespace AgentTeamMateBot.Services;

public class TeamsChatBot : ActivityHandler
{
    private static readonly Regex MentionMarkupPattern =
        new(@"</?at[^>]*>", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private readonly AppHostedMediaService _appHostedMediaService;
    private readonly MeetingContextService _meetingContextService;
    private readonly AiResponseService _aiResponseService;
    private readonly CloudAdapter _adapter;
    private readonly string _botAppId;

    public TeamsChatBot(
        AppHostedMediaService appHostedMediaService,
        MeetingContextService meetingContextService,
        AiResponseService aiResponseService,
        CloudAdapter adapter,
        IConfiguration configuration)
    {
        _appHostedMediaService = appHostedMediaService;
        _meetingContextService = meetingContextService;
        _aiResponseService = aiResponseService;
        _adapter = adapter;
        _botAppId = configuration["MicrosoftAppId"]
            ?? configuration["Bot:ClientId"]
            ?? throw new InvalidOperationException("MicrosoftAppId / Bot:ClientId missing");
    }

    protected override async Task OnMembersAddedAsync(
        IList<ChannelAccount> membersAdded,
        ITurnContext<IConversationUpdateActivity> turnContext,
        CancellationToken cancellationToken)
    {
        foreach (var member in membersAdded)
        {
            if (member.Id == turnContext.Activity.Recipient.Id)
            {
                continue;
            }

            await turnContext.SendActivityAsync(
                MessageFactory.Text(HelpText()),
                cancellationToken);
        }
    }

    protected override async Task OnMessageActivityAsync(
        ITurnContext<IMessageActivity> turnContext,
        CancellationToken cancellationToken)
    {
        var text = GetMessageText(turnContext.Activity);
        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" TEAMS CHAT MESSAGE");
        Console.WriteLine("================================================");
        Console.WriteLine(text);
        Console.WriteLine($"Conversation tenant: {turnContext.Activity.Conversation?.TenantId}");

        if (MeetingJoinParser.TryParse(text, out var meetingId, out var passcode))
        {
            await JoinFromChatAsync(
                turnContext,
                meetingId,
                passcode,
                cancellationToken);
            return;
        }

        await AnswerFromChatAsync(turnContext, text, cancellationToken);
    }

    private async Task JoinFromChatAsync(
        ITurnContext<IMessageActivity> turnContext,
        string meetingId,
        string? passcode,
        CancellationToken cancellationToken)
    {
        Console.WriteLine($"Parsed meetingId={meetingId} passcode={passcode}");

        if (!_appHostedMediaService.IsInitialized)
        {
            await turnContext.SendActivityAsync(
                MessageFactory.Text(
                    "I cannot join yet. App-hosted media is not ready on the VM."),
                cancellationToken);
            return;
        }

        await turnContext.SendActivityAsync(
            MessageFactory.Text($"Joining meeting {meetingId}..."),
            cancellationToken);

        var conversation = turnContext.Activity.GetConversationReference();

        _ = Task.Run(async () =>
        {
            try
            {
                var call = await _appHostedMediaService.JoinMeetingAsync(
                    meetingId,
                    passcode).ConfigureAwait(false);

                _meetingContextService.MergeChatIntoCall(
                    conversation.Conversation?.Id ?? string.Empty,
                    call.Id);

                await _adapter.ContinueConversationAsync(
                    _botAppId,
                    conversation,
                    async (ctx, ct) =>
                    {
                        await ctx.SendActivityAsync(
                            MessageFactory.Text(
                                $"I joined. Call ID {call.Id}. Ask me in chat or say Agent Nova in the meeting."),
                            ct);
                    },
                    CancellationToken.None);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[TEAMS CHAT JOIN] {ex}");
                try
                {
                    await _adapter.ContinueConversationAsync(
                        _botAppId,
                        conversation,
                        async (ctx, ct) =>
                        {
                            await ctx.SendActivityAsync(
                                MessageFactory.Text($"I could not join. {ex.Message}"),
                                ct);
                        },
                        CancellationToken.None);
                }
                catch (Exception notifyEx)
                {
                    Console.WriteLine($"[TEAMS CHAT JOIN] notify failed: {notifyEx}");
                }
            }
        });
    }

    private async Task AnswerFromChatAsync(
        ITurnContext<IMessageActivity> turnContext,
        string rawText,
        CancellationToken cancellationToken)
    {
        var speaker = string.IsNullOrWhiteSpace(turnContext.Activity.From?.Name)
            ? "Someone"
            : turnContext.Activity.From.Name.Trim();
        var spoken = StripMentions(rawText);
        var conversationId = turnContext.Activity.Conversation?.Id;
        var callId = _appHostedMediaService.ActiveCallId;

        var chatLine = $"Chat ({speaker}): {spoken}";
        foreach (var name in GetAttachmentNames(turnContext.Activity))
        {
            chatLine += $"{Environment.NewLine}Chat attachment: {name}";
        }

        _meetingContextService.AppendChatMessage(conversationId ?? "chat", callId, chatLine);
        Console.WriteLine($"[CHAT CONTEXT] {chatLine}");

        var contextId = callId ?? conversationId ?? "chat";

        if (!IsDirectedAtNova(turnContext))
        {
            return;
        }

        var question = spoken;
        if (WakeWordDetector.IsAgentInvocation(question))
        {
            question = WakeWordDetector.RemoveActivationPhrase(question);
        }

        if (string.IsNullOrWhiteSpace(question) ||
            !WakeWordDetector.IsActionableRequest(
                "Agent Nova " + question))
        {
            await turnContext.SendActivityAsync(
                MessageFactory.Text("I'm here. Ask about the meeting or the chat."),
                cancellationToken);
            return;
        }

        Console.WriteLine("[CHAT] Asking Nova...");
        try
        {
            var answer = await _aiResponseService.GetResponseAsync(
                contextId,
                question);
            if (string.IsNullOrWhiteSpace(answer))
            {
                answer = "I could not answer that just now. Please try again.";
            }

            await turnContext.SendActivityAsync(
                MessageFactory.Text(answer),
                cancellationToken);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[CHAT] Answer failed: {ex.Message}");
            await turnContext.SendActivityAsync(
                MessageFactory.Text("I hit an error answering that. Please try again."),
                cancellationToken);
        }
    }

    private static bool IsDirectedAtNova(ITurnContext<IMessageActivity> turnContext)
    {
        var conversation = turnContext.Activity.Conversation;
        if (conversation != null &&
            (conversation.IsGroup != true ||
             string.Equals(
                 conversation.ConversationType,
                 "personal",
                 StringComparison.OrdinalIgnoreCase)))
        {
            return true;
        }

        if (turnContext.Activity.Entities != null &&
            turnContext.Activity.Entities.Any(entity =>
                string.Equals(entity.Type, "mention", StringComparison.OrdinalIgnoreCase)))
        {
            return true;
        }

        return WakeWordDetector.IsAgentInvocation(turnContext.Activity.Text);
    }

    private static string StripMentions(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        var stripped = MentionMarkupPattern.Replace(text, " ");
        stripped = Regex.Replace(stripped, @"<[^>]+>", " ");
        stripped = Regex.Replace(stripped, @"\s+", " ").Trim();
        return stripped;
    }

    private static IEnumerable<string> GetAttachmentNames(IMessageActivity activity)
    {
        if (activity.Attachments == null)
        {
            yield break;
        }

        foreach (var attachment in activity.Attachments)
        {
            if (!string.IsNullOrWhiteSpace(attachment.Name))
            {
                yield return attachment.Name.Trim();
            }
        }
    }

    private static string GetMessageText(IMessageActivity activity)
    {
        var parts = new List<string>();

        if (!string.IsNullOrWhiteSpace(activity.Text))
        {
            parts.Add(activity.Text);
        }

        if (!string.IsNullOrWhiteSpace(activity.Summary))
        {
            parts.Add(activity.Summary);
        }

        if (activity.Attachments != null)
        {
            foreach (var attachment in activity.Attachments)
            {
                if (!string.IsNullOrWhiteSpace(attachment.ContentUrl))
                {
                    parts.Add(attachment.ContentUrl);
                }

                if (!string.IsNullOrWhiteSpace(attachment.Name))
                {
                    parts.Add(attachment.Name);
                }

                if (attachment.Content != null)
                {
                    parts.Add(attachment.Content.ToString() ?? string.Empty);
                }
            }
        }

        return string.Join(Environment.NewLine, parts.Where(p => !string.IsNullOrWhiteSpace(p)));
    }

    private static string HelpText()
    {
        return
            "Paste a Teams join invite to add me to the meeting.\n" +
            "After I join, ask in this chat or say Agent Nova in the call.\n" +
            "Example invite:\n" +
            "Join: https://teams.microsoft.com/meet/251659872407654?p=pD0ef1v2FypexQJN3D\n" +
            "Meeting ID: 251 659 872 407 654\n" +
            "Passcode: TA6QM9KL";
    }
}

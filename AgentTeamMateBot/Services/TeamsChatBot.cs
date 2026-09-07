using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Schema;

namespace AgentTeamMateBot.Services;

public class TeamsChatBot : ActivityHandler
{
    private readonly AppHostedMediaService _appHostedMediaService;
    private readonly CloudAdapter _adapter;
    private readonly string _botAppId;

    public TeamsChatBot(
        AppHostedMediaService appHostedMediaService,
        CloudAdapter adapter,
        IConfiguration configuration)
    {
        _appHostedMediaService = appHostedMediaService;
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
        var text = turnContext.Activity.Text;
        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" TEAMS CHAT MESSAGE");
        Console.WriteLine("================================================");
        Console.WriteLine(text);
        Console.WriteLine($"Conversation tenant: {turnContext.Activity.Conversation?.TenantId}");

        if (!MeetingJoinParser.TryParse(text, out var meetingId, out var passcode))
        {
            await turnContext.SendActivityAsync(
                MessageFactory.Text(HelpText()),
                cancellationToken);
            return;
        }

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

        // Join off the Bot Framework request so Graph Communications does not pick up
        // the Teams chat tenant from this HTTP turn (tenant mismatch vs Call.TenantId).
        _ = Task.Run(async () =>
        {
            try
            {
                var call = await _appHostedMediaService.JoinMeetingAsync(
                    meetingId,
                    passcode).ConfigureAwait(false);

                await _adapter.ContinueConversationAsync(
                    _botAppId,
                    conversation,
                    async (ctx, ct) =>
                    {
                        await ctx.SendActivityAsync(
                            MessageFactory.Text(
                                $"I joined. Call ID {call.Id}. Say Agent Nova when you need me."),
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

    private static string HelpText()
    {
        return
            "Send the meeting ID and passcode, for example:\n" +
            "join 251 659 872 407 654 TA6QM9KL\n" +
            "or\n" +
            "Meeting ID: 251 659 872 407 654 Passcode: TA6QM9KL";
    }
}

using Microsoft.Bot.Builder;
using Microsoft.Bot.Schema;

namespace AgentTeamMateBot.Services;

public class TeamsChatBot : ActivityHandler
{
    private readonly AppHostedMediaService _appHostedMediaService;

    public TeamsChatBot(AppHostedMediaService appHostedMediaService)
    {
        _appHostedMediaService = appHostedMediaService;
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

        try
        {
            var call = await _appHostedMediaService.JoinMeetingAsync(
                meetingId,
                passcode);

            await turnContext.SendActivityAsync(
                MessageFactory.Text(
                    $"I joined. Call ID {call.Id}. Say Agent Nova when you need me."),
                cancellationToken);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[TEAMS CHAT JOIN] {ex}");
            await turnContext.SendActivityAsync(
                MessageFactory.Text($"I could not join. {ex.Message}"),
                cancellationToken);
        }
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

using MediaLogLevel = Microsoft.Skype.Bots.Media.LogLevel;

namespace AgentTeamMateBot.Services;

public class BotMediaLogger : IBotMediaLogger
{
    private readonly ILogger<BotMediaLogger> _logger;

    public BotMediaLogger(
        ILogger<BotMediaLogger> logger)
    {
        _logger = logger;
    }

    public void WriteLog(
        MediaLogLevel level,
        string logStatement)
    {
        var logLevel =
            level switch
            {
                MediaLogLevel.Error =>
                    LogLevel.Error,

                MediaLogLevel.Warning =>
                    LogLevel.Warning,

                MediaLogLevel.Information =>
                    LogLevel.Information,

                MediaLogLevel.Verbose =>
                    LogLevel.Trace,

                _ =>
                    LogLevel.Trace
            };

        _logger.Log(
            logLevel,
            "[MEDIA SDK] {Message}",
            logStatement);
    }
}
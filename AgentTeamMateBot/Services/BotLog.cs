using System.Text;

namespace AgentTeamMateBot.Services;

public static class BotLog
{
    private static readonly TextWriter RealOut = Console.Out;

    public static bool Verbose { get; private set; }

    public static void Configure(IConfiguration configuration)
    {
        Verbose = configuration.GetValue("Logging:Verbose", false);
        Console.SetOut(new GatedWriter());
    }

    public static void Info(string message)
    {
        RealOut.WriteLine(message);
    }

    private sealed class GatedWriter : TextWriter
    {
        public override Encoding Encoding => RealOut.Encoding;

        public override void Write(char value)
        {
            if (Verbose)
            {
                RealOut.Write(value);
            }
        }

        public override void Write(string? value)
        {
            if (Verbose)
            {
                RealOut.Write(value);
            }
        }

        public override void Write(char[] buffer, int index, int count)
        {
            if (Verbose)
            {
                RealOut.Write(buffer, index, count);
            }
        }

        public override void WriteLine(string? value)
        {
            if (Verbose)
            {
                RealOut.WriteLine(value);
            }
        }

        public override void WriteLine()
        {
            if (Verbose)
            {
                RealOut.WriteLine();
            }
        }

        public override void Flush()
        {
            RealOut.Flush();
        }
    }
}

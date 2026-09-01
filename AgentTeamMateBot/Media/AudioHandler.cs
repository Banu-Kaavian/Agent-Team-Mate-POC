using AgentTeamMateBot.Services;
using Microsoft.Skype.Bots.Media;

namespace AgentTeamMateBot.Media;

public class AudioHandler
{
    private readonly SpeechRecognitionService _speechService;
    private long _packetCount;
    private long _livePacketCount;
    private bool _formatLogged;

    public AudioHandler(SpeechRecognitionService speechService)
    {
        _speechService = speechService;
    }

    public void ProcessAudio(
        byte[] data,
        AudioFormat audioFormat,
        bool isSilence,
        string callId,
        long timestamp = 0)
    {
        if (data == null || data.Length == 0)
        {
            Console.WriteLine("[AUDIO] Empty packet received");
            return;
        }

        LogFormatOnce(audioFormat, data.Length);

        var pcm16kMono = ConvertToSpeechPcm(data, audioFormat);
        Interlocked.Increment(ref _packetCount);

        if (!isSilence)
        {
            var livePacket = Interlocked.Increment(ref _livePacketCount);
            if (livePacket == 1 || livePacket % 50 == 0)
            {
                Console.WriteLine();
                Console.WriteLine("================================================");
                Console.WriteLine(" LIVE TEAMS AUDIO RECEIVED");
                Console.WriteLine("================================================");
                Console.WriteLine($"Call ID    : {callId}");
                Console.WriteLine($"Bytes      : {data.Length}");
                Console.WriteLine($"Timestamp  : {timestamp}");
                Console.WriteLine($"Format     : {audioFormat}");
                Console.WriteLine("================================================");
            }
        }

        _speechService.ProcessAudio(pcm16kMono);
    }

    private void LogFormatOnce(AudioFormat audioFormat, int byteLength)
    {
        if (_formatLogged)
        {
            return;
        }

        _formatLogged = true;

        var (sampleRate, bits, channels) = DescribeFormat(audioFormat);
        Console.WriteLine();
        Console.WriteLine("================================================");
        Console.WriteLine(" TEAMS AUDIO FORMAT");
        Console.WriteLine("================================================");
        Console.WriteLine($"SDK format   : {audioFormat}");
        Console.WriteLine($"Sample rate  : {sampleRate} Hz");
        Console.WriteLine($"Bits         : {bits}");
        Console.WriteLine($"Channels     : {channels}");
        Console.WriteLine($"Frame bytes  : {byteLength}");
        Console.WriteLine("Speech input  : 16000 Hz, 16-bit, mono PCM");
        Console.WriteLine("================================================");
    }

    private static (int SampleRate, int Bits, int Channels) DescribeFormat(AudioFormat format)
    {
        return format switch
        {
            AudioFormat.Pcm16K => (16000, 16, 1),
            AudioFormat.Pcm44KStereo => (44100, 16, 2),
            _ => (16000, 16, 1)
        };
    }

    private static byte[] ConvertToSpeechPcm(byte[] source, AudioFormat format)
    {
        return format switch
        {
            AudioFormat.Pcm16K => source,
            AudioFormat.Pcm44KStereo => Convert44KStereoTo16KMono(source),
            _ => source
        };
    }

    private static byte[] Convert44KStereoTo16KMono(byte[] input)
    {
        const int inRate = 44100;
        const int outRate = SpeechRecognitionService.SpeechSampleRate;

        var stereoFrames = input.Length / 4;
        if (stereoFrames <= 0)
        {
            return Array.Empty<byte>();
        }

        var mono = new short[stereoFrames];
        for (var i = 0; i < stereoFrames; i++)
        {
            var left = BitConverter.ToInt16(input, i * 4);
            var right = BitConverter.ToInt16(input, i * 4 + 2);
            mono[i] = (short)((left + right) / 2);
        }

        var outSamples = (int)((long)stereoFrames * outRate / inRate);
        var output = new byte[outSamples * 2];

        for (var i = 0; i < outSamples; i++)
        {
            var srcIndex = (double)i * inRate / outRate;
            var idx = (int)srcIndex;
            var frac = srcIndex - idx;
            var s0 = mono[Math.Min(idx, stereoFrames - 1)];
            var s1 = mono[Math.Min(idx + 1, stereoFrames - 1)];
            var sample = (short)(s0 + ((s1 - s0) * frac));
            BitConverter.TryWriteBytes(output.AsSpan(i * 2, 2), sample);
        }

        return output;
    }
}

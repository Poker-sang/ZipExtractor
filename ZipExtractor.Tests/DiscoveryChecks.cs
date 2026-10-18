using Spectre.Console;
using ZipExtractor.Extraction;
using ZipExtractor.Logging;

internal static class DiscoveryChecks
{
    public static void Verify(string directory)
    {
        File.WriteAllText(Path.Combine(directory, "[ordinary].txt"), "ordinary file");
        var child = Directory.CreateDirectory(Path.Combine(directory, "child"));
        File.WriteAllBytes(Path.Combine(child.FullName, "renamed.data"), [0x37, 0x7a, 0xbc, 0xaf, 0x27, 0x1c]);
        var updates = new List<DiscoveryProgress>();
        var sets = ArchiveDiscovery.Scan(directory, progress: updates.Add);
        if (sets.Count is not 1 || updates[^1] is not { Stage: "扫描完成", FilesScanned: 2, ArchivesFound: 1 })
            throw new InvalidOperationException("Recursive discovery counts are incorrect.");
        if (updates.FindIndex(p => p.Stage == "读取文件头") > updates.FindIndex(p => p.Path == child.FullName && p.Stage == "枚举目录"))
            throw new InvalidOperationException("Discovery still buffers the entire tree before inspecting files.");

        using var cancellation = new CancellationTokenSource();
        try
        {
            _ = ArchiveDiscovery.Scan(directory, progress: p =>
            {
                if (p.Stage == "枚举目录")
                    cancellation.Cancel();
            }, token: cancellation.Token);
            throw new InvalidOperationException("Directory enumeration ignored cancellation.");
        }
        catch (OperationCanceledException)
        {
        }

        using var writer = new StringWriter();
        var console = AnsiConsole.Create(new() { Out = new AnsiConsoleOutput(writer), Ansi = AnsiSupport.No, Interactive = InteractionSupport.No });
        console.Profile.Width = 500;
        var log = new ConsoleLog(console);
        var probe = new WaitingProbe(() => writer.ToString());
        _ = log.ScanArchives(directory, probe);
        if (!probe.SawHeartbeat || !writer.ToString().Contains("扫描完成", StringComparison.Ordinal) || writer.ToString().Contains('\u001b'))
            throw new InvalidOperationException("Scan heartbeat or plain output is missing.");
    }

    private sealed class WaitingProbe(Func<string> output) : IArchiveProbe
    {
        public bool SawHeartbeat { get; private set; }

        public ArchiveFormat Probe(string path, CancellationToken token = default)
        {
            // Wait inside the probe: progress must arrive without another discovery callback.
            Thread.Sleep(2500);
            SawHeartbeat = output().Contains("深度探测", StringComparison.Ordinal) && output().Contains("[ordinary].txt", StringComparison.Ordinal);
            return ArchiveFormat.Unknown;
        }
    }
}

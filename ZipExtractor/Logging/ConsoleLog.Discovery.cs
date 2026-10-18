using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using ZipExtractor.Extraction;

namespace ZipExtractor.Logging;

public sealed partial class ConsoleLog
{
    public IReadOnlyList<ArchiveSet> ScanArchives(string directory, IArchiveProbe? probe = null, CancellationToken token = default)
    {
        var elapsed = Stopwatch.StartNew();
        var current = new DiscoveryProgress("枚举目录", directory, 0, 0);
        var gate = new Lock();
        var finished = false;
        void WriteProgress()
        {
            var snapshot = Volatile.Read(ref current);
            Info($"{snapshot.Stage} · 已扫描 {snapshot.FilesScanned} 个文件 · 已发现 {snapshot.ArchivesFound} 个任务 · 耗时 {elapsed.Elapsed:hh\\:mm\\:ss} · {snapshot.Path}");
        }

        Info("扫描压缩包与分卷…（每 2 秒报告状态，Ctrl+C 取消）");
        // A timer keeps the current path visible even while a native probe is still running.
        using var timer = new Timer(_ =>
        {
            lock (gate)
                if (!finished)
                    WriteProgress();
        }, null, TimeSpan.FromSeconds(2), TimeSpan.FromSeconds(2));
        try
        {
            return ArchiveDiscovery.Scan(directory, skipGameFolders: true, probe: probe, progress: value => Volatile.Write(ref current, value), token: token);
        }
        finally
        {
            lock (gate)
            {
                finished = true;
                WriteProgress();
            }
        }
    }
}

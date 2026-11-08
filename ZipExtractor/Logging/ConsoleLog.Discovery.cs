using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using Spectre.Console;
using ZipExtractor.Extraction;

namespace ZipExtractor.Logging;

public sealed partial class ConsoleLog
{
    public IReadOnlyList<ArchiveSet> ScanArchives(string directory, IArchiveProbe? probe = null, CancellationToken token = default)
    {
        Info("扫描压缩包与分卷…（Ctrl+C 取消）");
        if (!console.Profile.Capabilities.Interactive || !console.Profile.Capabilities.Ansi)
            return ScanArchivesCore(directory, probe, null, token);
        var progress = console.Progress().AutoClear(true);
        _ = progress.Columns(new SpinnerColumn(), new TaskDescriptionColumn { Wrap = false, Alignment = Justify.Left });
        return progress.Start(context =>
        {
            var task = context.AddTask("准备扫描");
            task.IsIndeterminate = true;
            try
            {
                return ScanArchivesCore(directory, probe, message => task.Description = Escape(message), token);
            }
            finally
            {
                task.StopTask();
            }
        });
    }

    private IReadOnlyList<ArchiveSet> ScanArchivesCore(string directory, IArchiveProbe? probe, Action<string>? update, CancellationToken token)
    {
        var elapsed = Stopwatch.StartNew();
        var current = new DiscoveryProgress("枚举目录", directory, 0, 0);
        var gate = new Lock();
        var finished = false;
        void WriteProgress()
        {
            var snapshot = Volatile.Read(ref current);
            var message = $"{snapshot.Stage} · 已扫描 {snapshot.FilesScanned} 个文件 · 已发现 {snapshot.ArchivesFound} 个任务 · 耗时 {elapsed.Elapsed:hh\\:mm\\:ss} · {snapshot.Path}";
            if (update is null)
                Info(message);
            else
                update(message);
        }

        // A timer keeps the current path visible even while a native probe is still running.
        using var timer = new Timer(_ =>
        {
            lock (gate)
                if (!finished)
                    WriteProgress();
        }, null, TimeSpan.Zero, update is null ? TimeSpan.FromSeconds(2) : TimeSpan.FromMilliseconds(200));
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

using System;
using System.Diagnostics;
using System.Threading;
using Spectre.Console;
using Spectre.Console.Rendering;
using ZipExtractor.Extraction;

namespace ZipExtractor.Logging;

internal sealed class TransferProgressColumn : ProgressColumn
{
    private readonly Stopwatch _timer = Stopwatch.StartNew();
    private readonly Lock _gate = new();
    private ExtractionProgress? _progress;

    public void Reset()
    {
        lock (_gate)
        {
            _progress = null;
            _timer.Restart();
        }
    }

    public void Update(ExtractionProgress progress)
    {
        lock (_gate)
            _progress = progress;
    }

    public override IRenderable Render(RenderOptions options, ProgressTask task, TimeSpan deltaTime)
    {
        lock (_gate)
        {
            if (_progress is not { } progress)
                return new Text("等待引擎…", new(Color.Grey));
            var amount = progress.Total is > 0 ? $"{Math.Clamp(100.0 * progress.Completed / progress.Total.Value, 0, 100):F1}%" : "处理中";
            var rate = ConsoleLog.Bytes(progress.Completed / Math.Max(_timer.Elapsed.TotalSeconds, 0.1));
            return new Text($"{amount} · {ConsoleLog.Bytes(progress.Completed)} · {rate}/s · {_timer.Elapsed:mm\\:ss}", new(Color.Cyan));
        }
    }
}

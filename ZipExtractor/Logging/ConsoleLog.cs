using System;
using System.IO;
using System.Threading;
using Spectre.Console;
using ZipExtractor.Extraction;

namespace ZipExtractor.Logging;

/// <summary>One console renderer for status messages and native progress callbacks.</summary>
public sealed partial class ConsoleLog(IAnsiConsole console)
{
    private readonly Lock _gate = new();
    private readonly IAnsiConsole _console = console;

    public static ConsoleLog Default { get; } = new(AnsiConsole.Create(new() { Ansi = Console.IsOutputRedirected || Environment.GetEnvironmentVariable("NO_COLOR") is not null ? AnsiSupport.No : AnsiSupport.Detect, Interactive = Console.IsOutputRedirected ? InteractionSupport.No : InteractionSupport.Detect }));

    public void Header(PipelineOptions options)
    {
        _console.Write(new Rule("[bold cyan]ZipExtractor[/]  [grey]7-Zip + UnRAR[/]").LeftJustified());
        Detail("输入", Path.GetFullPath(options.InputDirectory));
        Detail("输出", Path.GetFullPath(options.CompleteDirectory));
        Detail("源包", options.KeepSources ? "保留" : "全部校验成功后送入回收站");
        _console.WriteLine();
    }

    public void Archive(int index, int count, ArchiveSet archive)
    {
        _console.WriteLine();
        _console.MarkupLine($"[bold cyan]▶ {index}/{count}  {Escape(Path.GetFileName(archive.EntryPath))}[/] [grey]· {archive.Format} · {archive.Parts.Count} 卷[/]");
        Detail("来源", archive.EntryPath);
    }

    public ExtractionResult RunAttempts(Func<AttemptSession, ExtractionResult> action)
    {
        if (!_console.Profile.Capabilities.Interactive || !_console.Profile.Capabilities.Ansi)
            return action(new(this));
        var transfer = new TransferProgressColumn();
        var progress = _console.Progress().AutoClear(true).HideCompleted(false);
        _ = progress.Columns(new SpinnerColumn(), new TaskDescriptionColumn { Wrap = false }, new ProgressBarColumn { Width = Math.Clamp(_console.Profile.Width / 5, 10, 24) }, transfer);
        return progress.Start(context =>
        {
            var task = context.AddTask("[grey]准备解压[/]");
            task.IsIndeterminate = true;
            try
            {
                return action(new(this, context, task, transfer));
            }
            finally
            {
                task.StopTask();
            }
        });
    }

    public void Info(string message, int level = 0) => Write("信息", message, "cyan", level);

    public void Success(string message, int level = 0) => Write("完成", message, "green", level);

    public void Warning(string message, int level = 0) => Write("注意", message, "yellow", level);

    public void Error(string message, int level = 0) => Write("失败", message, "red", level);

    public void Detail(string label, string message) => Write(label, message, "grey");

    public void GameSkipped(string path) => Write("保护", $"游戏目录「{DisplayName(path)}」：保留内容，跳过后续处理", "yellow");

    public void Summary(int count, int failures, TimeSpan elapsed)
    {
        _console.WriteLine();
        var color = failures is 0 ? "green" : "yellow";
        _console.Write(new Rule($"[{color}]成功 {count - failures} · 失败 {failures} · 共 {count} 个任务 · 用时 {elapsed:hh\\:mm\\:ss}[/]").LeftJustified());
    }

    public static string StatusText(ExtractionStatus status) => status switch
    {
        ExtractionStatus.Success => "校验通过",
        ExtractionStatus.WrongPassword => "密码未通过",
        ExtractionStatus.PasswordOrCorruptData => "密码或加密数据校验未通过",
        ExtractionStatus.MissingVolume => "缺少分卷",
        ExtractionStatus.InvalidArchive => "压缩包损坏或结构无效",
        ExtractionStatus.Unsupported => "不支持的内容或操作",
        ExtractionStatus.IoError => "文件读写错误",
        ExtractionStatus.Cancelled => "已取消",
        _ => throw new ArgumentOutOfRangeException(nameof(status))
    };

    public static string Bytes(double bytes) => bytes switch
    {
        >= 1073741824 => $"{bytes / 1073741824:F1} GiB",
        >= 1048576 => $"{bytes / 1048576:F1} MiB",
        >= 1024 => $"{bytes / 1024:F1} KiB",
        _ => $"{bytes:F0} B"
    };

    private void Write(string label, string message, string color, int level = 0)
    {
        lock (_gate)
            _console.MarkupLine($"[grey]{DateTime.Now:HH:mm:ss}[/] {new(' ', Math.Min(level, 8) * 2)}[{color}][[{Escape(label)}]][/] {Escape(message)}");
    }

    private static string Escape(string value) => Markup.Escape(value.Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t").Replace("\u001b", "\\e"));

    private static string DisplayName(string path)
    {
        var name = Path.GetFileName(path.TrimEnd(Path.DirectorySeparatorChar));
        return name.StartsWith("attempt-", StringComparison.Ordinal) ? "当前解压结果" : name;
    }

}

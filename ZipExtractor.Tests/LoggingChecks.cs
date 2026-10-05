using System;
using System.IO;
using System.Text.RegularExpressions;
using Spectre.Console;
using ZipExtractor.Extraction;
using ZipExtractor.Logging;

internal static partial class LoggingChecks
{
    public static void Verify(IArchiveExtractor engine, ArchiveSet archive, string password, string root, bool interactive)
    {
        using var writer = new StringWriter();
        var console = AnsiConsole.Create(new() { Out = new AnsiConsoleOutput(writer), Ansi = interactive ? AnsiSupport.Yes : AnsiSupport.No, Interactive = interactive ? InteractionSupport.Yes : InteractionSupport.No, ColorSystem = ColorSystemSupport.TrueColor });
        console.Profile.Width = 120;
        var log = new ConsoleLog(console);
        const string visiblePassword = "中文 [red]literal[/] 密码";
        var renderedPassword = false;
        var renderedWarning = false;
        var success = log.RunAttempts(session =>
        {
            session.Attempt(engine.Name, visiblePassword, 1, 2, 1);
            renderedPassword = AnsiSequences().Replace(writer.ToString(), "").Contains(visiblePassword, StringComparison.Ordinal);
            var bad = Path.Combine(root, Guid.NewGuid().ToString("N"));
            _ = Directory.CreateDirectory(bad);
            var failure = session.Extract(engine, new(archive, bad, "incorrect"));
            if (failure.Succeeded)
                throw new InvalidOperationException("Logging changed the failed extraction result.");
            var warningStart = writer.GetStringBuilder().Length;
            session.Warning(ConsoleLog.StatusText(failure.Status));
            var warning = AnsiSequences().Replace(writer.ToString()[warningStart..], "");
            renderedWarning = warning.Contains("注意", StringComparison.Ordinal) && !warning.Contains("尝试", StringComparison.Ordinal);
            session.Attempt(engine.Name, password, 2, 2, 1);
            var good = Path.Combine(root, Guid.NewGuid().ToString("N"));
            _ = Directory.CreateDirectory(good);
            var result = session.Extract(engine, new(archive, good, password));
            return !result.Succeeded || !File.Exists(Path.Combine(good, "内容.bin"))
                ? throw new InvalidOperationException("Logging prevented extraction or retained a failed live session.")
                : result;
        });
        if (!success.Succeeded)
            throw new InvalidOperationException("Attempt session changed the extraction result.");
        if (interactive && (!renderedPassword || !renderedWarning))
            throw new InvalidOperationException("Transient password/warning row did not update correctly.");
        log.Success("[literal] 解压完成");
        log.Summary(2, 1, TimeSpan.FromSeconds(3));
        var output = writer.ToString();
        var plain = AnsiSequences().Replace(output, "");
        if (!plain.Contains("[literal] 解压完成", StringComparison.Ordinal) || (!interactive && (plain.Contains(visiblePassword, StringComparison.Ordinal) || plain.Contains("尝试", StringComparison.Ordinal) || plain.Contains("注意", StringComparison.Ordinal))))
            throw new InvalidOperationException("Password or filename was parsed as markup.");
        if (interactive != output.Contains('\u001b'))
            throw new InvalidOperationException("ANSI mode did not match the output destination.");
    }

    [GeneratedRegex("\\x1b\\[[0-?]*[ -/]*[@-~]")]
    private static partial Regex AnsiSequences();
}

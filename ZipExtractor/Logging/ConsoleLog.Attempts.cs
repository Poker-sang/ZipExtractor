using System;
using Spectre.Console;
using ZipExtractor.Extraction;

namespace ZipExtractor.Logging;

public sealed partial class ConsoleLog
{
    /// <summary>A single transient row shared by all password attempts for one archive.</summary>
    public sealed class AttemptSession
    {
        private readonly ConsoleLog _owner;
        private readonly ProgressContext? _context;
        private readonly ProgressTask? _task;
        private readonly TransferProgressColumn? _transfer;
        private string _candidate = "";

        internal AttemptSession(ConsoleLog owner)
        {
            _owner = owner;
        }

        internal AttemptSession(ConsoleLog owner, ProgressContext context, ProgressTask task, TransferProgressColumn transfer) : this(owner)
        {
            _context = context;
            _task = task;
            _transfer = transfer;
        }

        public void Attempt(string engine, string? password, int index, int count, int level)
        {
            var label = password switch
            {
                null => "无密码",
                "" => "密码「」（空字符串）",
                _ => $"密码「{password}」"
            };
            _candidate = $"{new(' ', Math.Min(level, 8) * 2)}{engine} · {index}/{count} · {label}";
            if (_task is null)
                return;
            _transfer!.Reset();
            _task.Value = 0;
            _task.IsIndeterminate = true;
            Update("尝试", "grey");
        }

        public void Warning(string message) => Update("注意", "yellow", message);

        public ExtractionResult Extract(IArchiveExtractor engine, ExtractionRequest request)
        {
            // Redirected output cannot erase old lines. Keep only permanent task results there.
            if (_task is null)
                return engine.Extract(request);
            var reporter = new CallbackProgress(value =>
            {
                request.Progress?.Report(value);
                _transfer!.Update(value);
                _task.IsIndeterminate = value.Total is null or 0;
                if (value.Total is > 0)
                {
                    _task.MaxValue = value.Total.Value;
                    _task.Value = Math.Min(value.Completed, value.Total.Value);
                }
            });
            return engine.Extract(request with { Progress = reporter });
        }

        private void Update(string status, string color, string? message = null)
        {
            if (_task is null)
                return;
            lock (_owner._gate)
            {
                _task.Description = $"[{color}][[{Escape(status)}]] {Escape(_candidate)}{(message is null ? "" : " · " + Escape(message))}[/]";
                _context!.Refresh();
            }
        }
    }
}

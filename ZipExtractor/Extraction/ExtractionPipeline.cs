using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using ZipExtractor.Logging;

namespace ZipExtractor.Extraction;

public sealed record PipelineOptions(string InputDirectory, string CompleteDirectory, bool KeepSources = false, int MaxDepth = 20, bool UpdateDirectoryLastWriteTime = false);
/// <summary>Owns recursion and transactions; backends only extract one archive with one password.</summary>
public sealed class ExtractionPipeline(IReadOnlyList<IArchiveExtractor> engines, IArchiveProbe? probe = null)
{
    private readonly IArchiveProbe? _probe = probe ?? engines.OfType<IArchiveProbe>().FirstOrDefault();

    public int Run(PipelineOptions options, IReadOnlyCollection<string?> passwords, IProgress<ExtractionProgress>? progress = null, CancellationToken cancellationToken = default)
    {
        var input = Path.GetFullPath(options.InputDirectory);
        var complete = Path.GetFullPath(options.CompleteDirectory);
        if (Contains(input, complete) || Contains(complete, input))
            throw new ArgumentException("输入和完成目录必须互不包含。");
        if (options.MaxDepth < 1)
            throw new ArgumentOutOfRangeException(nameof(options.MaxDepth));
        _ = Directory.CreateDirectory(complete);
        var log = ConsoleLog.Default;
        var timer = Stopwatch.StartNew();
        log.Info("扫描压缩包与分卷…");
        // Snapshot before extracting or moving anything. Never enumerate a changing directory tree.
        var sets = ArchiveDiscovery.Scan(input, skipGameFolders: true, probe: _probe, token: cancellationToken);
        var failures = 0;
        var index = 0;
        log.Info($"发现 {sets.Count} 个解压任务");
        foreach (var set in sets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var job = Path.Combine(complete, ".extract-" + Guid.NewGuid().ToString("N"));
            _ = Directory.CreateDirectory(job);
            try
            {
                log.Archive(++index, sets.Count, set);
                var result = ExtractTree(set, job, passwords, options.MaxDepth, progress, out var extracted, 0, options.UpdateDirectoryLastWriteTime, cancellationToken);
                if (!result.Succeeded)
                {
                    failures++;
                    log.Error($"{ConsoleLog.StatusText(result.Status)}：{result.Message}；当前任务源文件已保留");
                    if (result.Status == ExtractionStatus.Cancelled)
                        cancellationToken.ThrowIfCancellationRequested();
                    continue;
                }

                var relative = Path.GetRelativePath(input, Path.GetDirectoryName(set.EntryPath)!);
                var target = FileSystemHelper.GetUniquePath(Path.Combine(complete, relative, SafeName(set.DisplayName)), options.UpdateDirectoryLastWriteTime);
                _ = Directory.CreateDirectory(Path.GetDirectoryName(target)!);
                new DirectoryInfo(extracted!).MoveTo(target, options.UpdateDirectoryLastWriteTime);
                log.Success($"{set.DisplayName} · 全部校验通过");
                log.Detail("保存", target);
                if (!options.KeepSources)
                    foreach (var part in (result.VerifiedParts ?? [set.EntryPath]).Intersect(set.Parts, StringComparer.OrdinalIgnoreCase))
                        _ = new FileInfo(part).RemoveIfExists(moveToRecycleBin: true, updateDirectoryLastWriteTime: options.UpdateDirectoryLastWriteTime);
            }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException)
            {
                failures++;
                log.Error($"文件读写错误：{e.Message}；当前任务源文件已保留");
            }
            finally
            {
                DeleteWorkDirectory(job, options.UpdateDirectoryLastWriteTime);
            }
        }

        log.Summary(sets.Count, failures, timer.Elapsed);
        return failures;
    }

    private ExtractionResult ExtractTree(ArchiveSet set, string work, IReadOnlyCollection<string?> passwords, int depth, IProgress<ExtractionProgress>? progress, out string? extracted, int level, bool updateDirectoryLastWriteTime, CancellationToken token)
    {
        extracted = null;
        if (depth <= 0)
            return new(ExtractionStatus.Unsupported, "达到嵌套解压深度上限。");
        if (set.DiscoveryError is { } error)
            return new(ExtractionStatus.MissingVolume, error);
        var engine = engines.FirstOrDefault(e => e.Supports(set.Format));
        if (engine is null)
            return new(ExtractionStatus.Unsupported, $"没有 {set.Format} 解压引擎。");
        var candidates = new string?[]
        {
            null
        }.Concat(passwords).Distinct().ToArray();
        string? output = null;
        var result = ConsoleLog.Default.RunAttempts(session =>
        {
            var lastResult = new ExtractionResult(ExtractionStatus.WrongPassword, "没有可用密码。");
            ExtractionResult? ambiguous = null;
            for (var i = 0; i < candidates.Length; i++)
            {
                token.ThrowIfCancellationRequested();
                var attempt = Path.Combine(work, "attempt-" + Guid.NewGuid().ToString("N"));
                _ = Directory.CreateDirectory(attempt);
                session.Attempt(engine.Name, candidates[i], i + 1, candidates.Length, level);
                lastResult = session.Extract(engine, new(set, attempt, candidates[i], progress, token));
                if (lastResult.Succeeded)
                {
                    output = attempt;
                    return lastResult;
                }

                DeleteWorkDirectory(attempt, updateDirectoryLastWriteTime);
                if (lastResult.Status == ExtractionStatus.PasswordOrCorruptData)
                    ambiguous = lastResult;
                if (lastResult.Status is not (ExtractionStatus.WrongPassword or ExtractionStatus.PasswordOrCorruptData))
                    return lastResult;
                session.Warning(ConsoleLog.StatusText(lastResult.Status));
            }

            return ambiguous ?? lastResult;
        });
        extracted = output;
        if (extracted is null)
            return result;
        // Mixed content is preserved, but game folders and their descendants are never expanded.
        foreach (var nested in ArchiveDiscovery.Scan(extracted, skipGameFolders: true, probe: _probe, token: token))
        {
            token.ThrowIfCancellationRequested();
            var nestedWork = Path.Combine(work, "nested-" + Guid.NewGuid().ToString("N"));
            _ = Directory.CreateDirectory(nestedWork);
            ConsoleLog.Default.Info($"第 {level + 1} 层嵌套 · {Path.GetFileName(nested.EntryPath)}", level + 1);
            var child = ExtractTree(nested, nestedWork, passwords, depth - 1, progress, out var childOutput, level + 1, updateDirectoryLastWriteTime, token);
            if (!child.Succeeded)
                return new(child.Status, $"嵌套压缩包 {Path.GetFileName(nested.EntryPath)}：{child.Message}");
            var childTarget = FileSystemHelper.GetUniquePath(Path.Combine(Path.GetDirectoryName(nested.EntryPath)!, SafeName(nested.DisplayName)), updateDirectoryLastWriteTime);
            new DirectoryInfo(childOutput!).MoveTo(childTarget, updateDirectoryLastWriteTime);
            foreach (var part in (child.VerifiedParts ?? [nested.EntryPath]).Intersect(nested.Parts, StringComparer.OrdinalIgnoreCase))
            {
                File.SetAttributes(part, File.GetAttributes(part) & ~FileAttributes.ReadOnly);
                FileSystemHelper.DeleteFile(part, updateDirectoryLastWriteTime); // Only intermediates inside this transaction.
            }

            DeleteWorkDirectory(nestedWork, updateDirectoryLastWriteTime);
            ConsoleLog.Default.Success($"嵌套解压 · {nested.DisplayName}", level + 1);
        }

        return result;
    }

    private static void DeleteWorkDirectory(string path, bool updateDirectoryLastWriteTime)
    {
        if (!Directory.Exists(path))
            return;
        foreach (var file in Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories))
            File.SetAttributes(file, File.GetAttributes(file) & ~FileAttributes.ReadOnly);
        FileSystemHelper.DeleteDirectory(path, true, updateDirectoryLastWriteTime);
    }

    private static bool Contains(string parent, string child) => string.Equals(parent, child, StringComparison.OrdinalIgnoreCase) || child.StartsWith(parent.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);

    private static string SafeName(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        name = name.TrimEnd('.', ' ');
        return string.IsNullOrEmpty(name) ? "archive" : name;
    }
}

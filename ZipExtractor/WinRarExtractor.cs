using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace ZipExtractor;

public abstract class WinRarExtractor : ExtractorBase
{
    private const string WinRarPath = @"C:\App\WinRAR\WinRAR.exe";

    private static readonly HashSet<string> _SingleArchiveOrFirstVolumeExtensions =
    [
        ".rar", ".zip", ".7z", ".r00", ".001"
    ];

    /// <summary>
    /// 递归解压压缩包，支持密码尝试
    /// </summary>
    /// <param name="archivePath">压缩包路径</param>
    /// <param name="rootPath">根路径</param>
    /// <param name="passwords">可能的密码列表</param>
    public static void ExtractRecursively(string archivePath, string rootPath, string?[] passwords)
    {
        if (!File.Exists(WinRarPath))
        {
            Console.WriteLine($"WinRAR未找到：{WinRarPath}");
            return;
        }

        if (string.IsNullOrWhiteSpace(archivePath))
        {
            Console.WriteLine("压缩包路径不能为空");
            return;
        }

        if (string.IsNullOrWhiteSpace(rootPath))
        {
            Console.WriteLine("根路径不能为空");
            return;
        }

        var archiveFile = new FileInfo(archivePath);

        if (!archiveFile.Exists)
        {
            Console.WriteLine($"文件不存在：{archiveFile.FullName}");
            return;
        }

        var finalResults = new List<FileSystemInfo>();
        var archiveQueue = new Queue<FileInfo>();
        archiveQueue.Enqueue(archiveFile);

        while (archiveQueue.Count > 0)
        {
            var currentArchive = archiveQueue.Dequeue();

            // 首先将传入文件移动到 TempExt 并改名为 .zip
            var fileType = CheckFileCompressionType(currentArchive);
            switch (fileType)
            {
                case FileCompressionType.Text:
                    Console.WriteLine($"跳过文本文件：{archivePath}");
                    return;
                case FileCompressionType.Volume:
                    Console.WriteLine($"跳过非第一的分卷：{archivePath}");
                    return;
            }

            // 记录所有中间产生的压缩包（用于事后清理）
            if (PrepareIntoTemp(currentArchive, fileType, out var commonPrefix) is not { } intermediateArchives)
                return;

            Console.WriteLine();

            Console.WriteLine("正在解压：");

            foreach (var intermediateArchive in intermediateArchives)
                Console.WriteLine(intermediateArchive);

            // 解压到当前文件所在目录
            var extractDir = currentArchive.Directory ?? TempDir;

            // 避免冲突
            var outputFolder = FileSystemHelper.GetUniquePath(Path.Combine(extractDir.FullName, currentArchive.NameWithoutExtension));

            var outputDir = Directory.CreateDirectory(outputFolder);

            // 尝试解压：先尝试无密码，再尝试提供的密码
            var status = ExtractStatus.WrongPassword;

            foreach (var password in passwords)
            {
                Console.WriteLine($"尝试密码：{password ?? "（无密码）"}");
                status = ExtractWithVerify(currentArchive, outputDir, password);
                if (status is ExtractStatus.Success or ExtractStatus.Failed)
                    break;
            }

            if (status is not ExtractStatus.Success)
            {
                Console.WriteLine("解压失败，所有密码均不正确或文件损坏");
                var cp = intermediateArchives is [_] ? "" : commonPrefix!;
                MoveArchivesToError(intermediateArchives, cp);
                _ = outputDir.RemoveIfExists();
                return;
            }

            Console.WriteLine($"密码正确，解压成功：{outputDir.FullName}");

            var extractedTotalSize = FileSystemHelper.GetDirectorySize(outputDir);
            PromptCleanupIntermediates(intermediateArchives, extractedTotalSize);

            var extractedFiles = outputDir.GetFiles("*", SearchOption.AllDirectories);

            // 检查所有文件是否都符合"可继续解压"的条件
            string? anyOther = null;
            var archiveCandidates = new List<FileInfo>();
            foreach (var file in extractedFiles)
            {
                var type = CheckFileCompressionType(file);

                if (type is FileCompressionType.First 
                    or FileCompressionType.Rar
                    or FileCompressionType.Zip
                    or FileCompressionType._7Z)
                    archiveCandidates.Add(file);

                if (type is FileCompressionType.Other)
                {
                    anyOther = file.Name;
                    break;
                }
            }

            if (anyOther is not null)
            {
                Console.WriteLine($"检测到非压缩文件：{anyOther}，停止解压。");
                finalResults.Add(extractedFiles is [{ } onlyOne]
                    ? onlyOne
                    : outputDir);
                continue;
            }

            // 都符合，处理并加入队列
            foreach (var file in archiveCandidates)
                archiveQueue.Enqueue(file);
        }

        Console.WriteLine();

        if (finalResults.Count is 0)
        {
            Console.WriteLine("解压过程异常结束");
            return;
        }

        Console.WriteLine("所有解压完成！");

        // 移动到 CompleteExt 之前：合并前三层内相邻同名且父仅含该子目录的冗余嵌套
        foreach (var finalResult in finalResults.OfType<DirectoryInfo>())
            FileSystemHelper.NormalizeRedundantNestedFolders(finalResult);

        // 只有一个文件结果时，使用压缩包所在目录名作为目标文件夹名
        var destDirName = finalResults switch
        {
            [FileInfo] when archiveFile.Directory is { } directory
                            && !directory.FullName.EqualsFileName(TempPath)
                            && !directory.FullName.EqualsFileName(rootPath) => directory.Name,
            [DirectoryInfo] => "",
            _ => Path.GetFileNameWithoutExtension(archiveFile.Name)
        };

        MoveResultsToComplete(finalResults, destDirName);
    }

    private enum ExtractStatus
    {
        Success,
        WrongPassword,
        Failed
    }

    private static ExtractStatus ExtractWithVerify(FileInfo archivePath, DirectoryInfo outputFolder, string? password)
    {
        _ = ExtractCore(archivePath.FullName, outputFolder.FullName, password);
        outputFolder.Refresh();

        // 每次解压后进行大小比例校验
        var archiveSize = archivePath.Length;
        var extractedSize = FileSystemHelper.GetDirectorySize(outputFolder);
        var ratio = archiveSize > 0 ? (double)extractedSize / archiveSize : 0.0;

        switch (ratio)
        {
            case < 0.05:
            {
                if (ratio is not 0)
                    Console.WriteLine($"解压结果过小（{ratio:P2}<5%），可能密码错误，继续尝试其他密码...");
                _ = FileSystemHelper.CleanEmptyDirectories(outputFolder, true);
                return ExtractStatus.WrongPassword;
            }
            case < 0.5:
            {
                Console.WriteLine($"解压结果过小（{ratio:P2}<50%），判定为解压失败，已停止。");
                return ExtractStatus.Failed;
            }
            default:
                return ExtractStatus.Success;
        }
    }

    /// <summary>
    /// 执行解压操作
    /// </summary>
    private static bool ExtractCore(string archivePath, string outputPath, string? password)
    {
        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = WinRarPath,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            // 构建命令行参数
            // x: 解压带完整路径
            // -ibck: 后台模式
            // -or: 自动重命名已存在文件
            // -y: 对所有询问回答yes
            // -p: 解压密码
            var args = new StringBuilder("x -ibck -or -y -p")
                .Append(string.IsNullOrEmpty(password)
                    // 无密码
                    ? "-"
                    : $"\"{password}\"")
                .Append($" \"{archivePath}\" \"{outputPath}\\\"")
                .ToString();

            startInfo.Arguments = args;

            using var process = Process.Start(startInfo);
            if (process is null)
            {
                Console.WriteLine("无法启动WinRAR进程");
                return false;
            }

            process.WaitForExit();

            // WinRAR 返回 0 表示成功
            return process.ExitCode is 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"解压出错：{ex.Message}");
            return false;
        }
    }

    private static void MoveArchivesToError(IReadOnlyCollection<FileInfo> archiveFiles, string commonPrefix)
    {
        var destDir = ErrorPath;
        if (!string.IsNullOrWhiteSpace(commonPrefix))
            destDir = FileSystemHelper.GetUniquePath(Path.Combine(destDir, commonPrefix));
        foreach (var archiveFile in archiveFiles)
        {
            var dest = Path.Combine(destDir, archiveFile.Name);
            if (archiveFile.TryMoveTo(dest))
                Console.WriteLine("已移动失败压缩包到：" + dest);
            else
                Console.WriteLine($"移动失败压缩包到 {nameof(ErrorPath)} 失败：{archiveFile.FullName} -> {dest}");
        }
    }

    private static void MoveResultsToComplete(IReadOnlyCollection<FileSystemInfo> resultDirs, string destDirName)
    {
        var destDir = CompletePath;
        if (!string.IsNullOrWhiteSpace(destDirName))
            destDir = FileSystemHelper.GetUniquePath(Path.Combine(destDir, destDirName));
        foreach (var resultDir in resultDirs)
        {
            var dest = Path.Combine(destDir, resultDir.Name);
            if (resultDir.TryMoveTo(dest))
                Console.WriteLine("已移动解压结果到：" + dest);
            else
                Console.WriteLine($"移动解压结果到 {nameof(CompletePath)} 失败：{resultDir.FullName} -> {dest}");
        }
    }

    private static void PromptCleanupIntermediates(IReadOnlyCollection<FileInfo> intermediates, long extractedTotalSize)
    {
        // 自动清理判定：若解压后的总大小 >= 源文件大小的 50%，则直接清理
        var originalArchivesSize = FileSystemHelper.GetFilesSize(intermediates);
        var autoClean = originalArchivesSize > 0 && extractedTotalSize >= originalArchivesSize / 2;
        var clean = autoClean;

        Console.WriteLine();

        if (!autoClean)
        {
            // 等待用户输入
            Console.WriteLine("解压后大小不到源文件的50%，请检查是否有文件未解压：");
            foreach (var intermediate in intermediates)
                Console.WriteLine($" - {intermediate.FullName} ({intermediate.Length} 字节)");
            Console.WriteLine("按回车键删除所有中间压缩包，按其他键保留所有文件...");
            var key = Console.ReadKey(true);
            if (key.Key is ConsoleKey.Enter)
                clean = true;
        }

        if (clean)
        {
            Console.WriteLine("解压后大小超过源文件的50%，自动清理中间压缩包...");
            Console.WriteLine("正在清理中间文件...");

            foreach (var archive in intermediates)
                if (archive.FullName.StartWithFileName(TempPath) && archive.RemoveIfExists())
                    Console.WriteLine($"已删除：{archive.FullName}");
            _ = FileSystemHelper.CleanEmptyDirectories(TempPath, false);
            Console.WriteLine("清理完成！");
        }
        else
        {
            Console.WriteLine("已保留所有文件。");
        }
    }

    private static IReadOnlyList<FileInfo>? PrepareIntoTemp(FileInfo source, FileCompressionType fileType, out string? commonPrefix)
    {
        commonPrefix = null;
        try
        {
            if (!source.Exists)
                return null;

            if (fileType is FileCompressionType.Zip
                or FileCompressionType._7Z
                or FileCompressionType.Rar
                or FileCompressionType.Other)
            {
                // 按规则重命名
                var newExt = fileType switch
                {
                    FileCompressionType._7Z => ".7z",
                    FileCompressionType.Rar => ".rar",
                    _ => ".zip"
                };
                var destFile = Path.ChangeExtension(source.FullName, newExt);
                var newSourcePath = FileSystemHelper.GetUniquePath(destFile);
                _ = source.TryMoveTo(newSourcePath);
            }

            var list = FindSiblingSplitVolumes(source, out commonPrefix);

            if (list.Count is 0)
                throw new ArgumentException(source.Name);

            if (list is [var onlyVolume])
            {
                var destPathSingle = FileSystemHelper.GetUniquePath(Path.Combine(TempPath, onlyVolume.Name));
                _ = source.TryMoveTo(destPathSingle);
                return [source];
            }

            var tempDestDir = FileSystemHelper.GetUniquePath(Path.Combine(TempPath, commonPrefix!));
            var result = new List<FileInfo>();

            foreach (var volume in list)
            {
                var destPath = Path.Combine(tempDestDir, volume.Name);
                _ = volume.TryMoveTo(destPath);
                result.Add(volume);
            }

            return result;

        }
        catch (Exception ex)
        {
            Console.WriteLine($"准备文件失败：{ex.Message}");
            return null;
        }
    }

    private enum FileCompressionType
    {
        /// <summary>
        /// 第一分卷或是单个压缩文件
        /// </summary>
        First,

        /// <summary>
        /// 非第一分卷
        /// </summary>
        Volume,

        /// <summary>
        /// z开头或p结尾
        /// </summary>
        Zip,

        /// <summary>
        /// 7开头或z结尾
        /// </summary>
        _7Z,

        /// <summary>
        /// r开头或r结尾
        /// </summary>
        Rar,

        /// <summary>
        /// 其他文件
        /// </summary>
        Other,

        /// <summary>
        /// 文本文件
        /// </summary>
        Text
    }

    private static FileCompressionType CheckFileCompressionType(FileInfo file)
    {
        var (ext1, ext2) = FileSystemHelper.GetLastTwoExtensions(file.Name.ToLower());

        // 文本文件且小于1MB
        if (ext2 is "txt" or "md" && file.Length < 1 << 20)
            return FileCompressionType.Text;

        // part1.rar/part2.rar
        if (ext2 is "rar" && ext1 is ['p', 'a', 'r', 't', .. var idx] && int.TryParse(idx, out var result1))
            return result1 > 1 ? FileCompressionType.Volume : FileCompressionType.First;

        // 检查是否为第一卷
        if (_SingleArchiveOrFirstVolumeExtensions.Contains('.' + ext2))
            return FileCompressionType.First;

        switch (ext1)
        {
            // 7z.001
            case "7z" when int.TryParse(ext2, out var result4):
                return result4 > 1 ? FileCompressionType.Volume : FileCompressionType.First;
            // zip.001/zipx.001
            case "zip" or "zipx" when int.TryParse(ext2, out _):
                return FileCompressionType.Volume;
        }

        return ext2 switch
        {
            // z01
            ['z', .. { Length: >= 2 } remains]
                when int.TryParse(remains, out _) => FileCompressionType.Volume,
            // r00/r01
            ['r', .. { Length: >= 2 } remains]
                when int.TryParse(remains, out var result2) =>
                result2 > 0
                    ? FileCompressionType.Volume
                    : FileCompressionType.First,
            _ => ext2 switch
            {
                ['z', ..] or [.., 'p'] => FileCompressionType.Zip,
                ['7', ..] or [.., 'z'] => FileCompressionType._7Z,
                ['r', ..] or [.., 'r'] => FileCompressionType.Rar,
                _ => FileCompressionType.Other
            }
        };
    }

    /// <summary>
    /// 给定第一卷压缩包完整路径，查找同级目录下的所有其他分卷。
    /// 支持以下常见分卷命名：
    /// <code>
    /// - name.part1.rar → name.partN.rar
    /// - name.r00 → name.rNN
    /// - name.7z.001 / name.zip.001 / name.zipx.001 → name.ext.00N
    /// - name.zip / name.zipx → name.zNN
    /// </code>
    /// 返回按分卷序号升序排序的完整路径列表（包含第一卷自身）。
    /// </summary>
    private static IList<FileInfo> FindSiblingSplitVolumes(FileInfo firstVolume, out string? commonPrefix)
    {
        commonPrefix = null;

        if (firstVolume.Directory is not { Exists: true } dir)
            return [];

        // Name 分为 nameWithoutExt.ext1.ext2
        var (nameWithoutExt2, nameExt2) = FileSystemHelper.GetExtension(firstVolume.Name);
        nameExt2 = nameExt2?.ToLower();
        var (nameWithoutExt12, nameExt1) = FileSystemHelper.GetExtension(nameWithoutExt2);
        nameExt1 = nameExt1?.ToLower();

        var allFiles = dir.GetFiles("*", SearchOption.TopDirectoryOnly);

        var results = new SortedList<int, FileInfo>();

        foreach (var file in allFiles)
        {
            var (withoutExt2, ext2) = FileSystemHelper.GetExtension(file.Name);
            ext2 = ext2?.ToLower();
            switch (nameExt1, nameExt2)
            {
                // 1) name.part1.rar → name.partN.rar
                case (['p', 'a', 'r', 't', ..], "rar"):
                {
                    commonPrefix = nameWithoutExt12;
                    var (withoutExt12, ext1) = FileSystemHelper.GetExtension(withoutExt2);
                    ext1 = ext1?.ToLower();
                    if (withoutExt12.EqualsFileName(commonPrefix)
                        && ext1 is ['p', 'a', 'r', 't', .. var idx]
                        && ext2 is "rar"
                        && int.TryParse(idx, out var i))
                        results[i] = file;
                    break;
                }
                // 2) name.r00 → name.rNN
                case (_, ['r', .. { Length: >= 2 } remains])
                    when int.TryParse(remains, out _):
                {
                    commonPrefix = nameWithoutExt2;
                    if (withoutExt2.EqualsFileName(commonPrefix)
                        && ext2 is ['r', .. { Length: >= 2 } rm]
                        && int.TryParse(rm, out var i))
                        results[i] = file;
                    break;
                }
                // 3) name.zip.001 / name.zipx.001 / name.7z.001 → name.ext.00N
                case ("7z" or "zip" or "zipx", _) when int.TryParse(nameExt2, out _):
                {
                    commonPrefix = nameWithoutExt2;
                    if (withoutExt2.EqualsFileName(commonPrefix)
                        && int.TryParse(ext2, out var i))
                        results[i] = file;
                    break;
                }
                // 4) name.zip / name.zipx → name.zNN
                case (_, "zip" or "zipx"):
                {
                    commonPrefix = nameWithoutExt2;
                    results[0] = firstVolume; // 第一卷不会遍历到，手动添加
                    if (withoutExt2.EqualsFileName(commonPrefix)
                        && ext2 is ['z', .. { Length: >= 2 } remains]
                        && int.TryParse(remains, out var i))
                        results[i] = file;
                    break;
                }
                // 5) 单个压缩包
                case (null, "rar" or "7z"):
                    return [file];
            }
        }

        return results.Values;
    }
}

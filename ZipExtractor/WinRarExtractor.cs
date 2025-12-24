using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace ZipExtractor;

public static class WinRarExtractor
{
    private const string WinRarPath = @"C:\App\WinRAR\WinRAR.exe";
    private const string BasePath = @"D:\";
    private static readonly string _TempPath = Path.Combine(BasePath, "TempExt");
    private static readonly string _CompletePath = Path.Combine(BasePath, "CompleteExt");
    private static readonly string _ErrorPath = Path.Combine(BasePath, "ErrorExt");

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

        if (!File.Exists(WinRarPath))
        {
            Console.WriteLine($"WinRAR未找到：{WinRarPath}");
            return;
        }

        var tempDir = Directory.CreateDirectory(_TempPath);
        _ = Directory.CreateDirectory(_CompletePath);
        _ = Directory.CreateDirectory(_ErrorPath);

        var finalResults = new List<DirectoryInfo>();
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

            Console.WriteLine("\n正在解压：");

            foreach (var intermediateArchive in intermediateArchives)
                Console.WriteLine(intermediateArchive);

            // 解压到当前文件所在目录
            var extractPath = currentArchive.Directory ?? tempDir;

            // 避免冲突
            var outputFolder = GetUniquePath(Path.Combine(extractPath.FullName, Path.GetFileNameWithoutExtension(currentArchive.Name)));

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
                MoveArchivesToError(intermediateArchives, commonPrefix);
                _ = RemoveIfExists(outputDir);
                return;
            }

            Console.WriteLine($"密码正确，解压成功：{outputDir.FullName}");

            var extractedTotalSize = GetDirectorySize(outputDir);
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
                finalResults.Add(outputDir);
                continue;
            }

            // 都符合，处理并加入队列
            foreach (var file in archiveCandidates)
                archiveQueue.Enqueue(file);
        }

        if (finalResults.Count is 0)
        {
            Console.WriteLine("\n解压过程异常结束");
            return;
        }

        Console.WriteLine("\n所有解压完成！");

        // 移动到 CompleteExt 之前：合并前三层内相邻同名且父仅含该子目录的冗余嵌套
        foreach (var finalResult in finalResults)
            NormalizeRedundantNestedFolders(finalResult);

        // 只有一个结果时，使用压缩包所在目录名作为目标文件夹名
        var destPath = Path.Combine(_CompletePath, finalResults.Count is 1
            ? archiveFile.Directory?.Name ?? ""
            : Path.GetFileNameWithoutExtension(archiveFile.Name));

        MoveResultsToComplete(finalResults, destPath);
    }

    private static ExtractStatus ExtractWithVerify(FileInfo archivePath, DirectoryInfo outputFolder, string? password)
    {
        _ = ExtractCore(archivePath.FullName, outputFolder.FullName, password);

        // 每次解压后进行大小比例校验
        var archiveSize = archivePath.Length;
        var extractedSize = GetDirectorySize(outputFolder);
        var ratio = archiveSize > 0 ? (double)extractedSize / archiveSize : 0.0;
        Console.WriteLine($"大小校验：解压后 {extractedSize} / 解压前 {archiveSize} = {ratio:P2}");

        switch (ratio)
        {
            case < 0.05:
            {
                if (ratio is not 0)
                    Console.WriteLine("解压结果过小（<5%），可能密码错误，继续尝试其他密码...");
                _ = CleanEmptyDirectories(outputFolder, true);
                return ExtractStatus.WrongPassword;
            }
            case < 0.5:
            {
                Console.WriteLine("解压结果过小（<50%），判定为解压失败，已停止。");
                return ExtractStatus.Failed;
            }
            default:
                return ExtractStatus.Success;
        }
    }

    private enum ExtractStatus
    {
        Success,
        WrongPassword,
        Failed
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

    private static long GetDirectorySize(DirectoryInfo dir)
    {
        return dir.Exists ? GetFilesSize(dir.GetFiles("*", SearchOption.AllDirectories)) : 0;
    }

    private static long GetFilesSize(IEnumerable<FileInfo> files)
    {
        long sum = 0;
        foreach (var file in files)
            try
            {
                sum += file.Length;
            }
            catch
            {
            }

        return sum;
    }

    private static void MoveArchivesToError(IReadOnlyList<FileInfo> archiveFiles, string? commonPrefix)
    {
        if (archiveFiles is [var onlyOne])
        {
            var dest = GetUniquePath(Path.Combine(_ErrorPath, onlyOne.Name));
            if (TryMoveEntry(onlyOne, dest))
                Console.WriteLine("已移动失败压缩包到：" + dest);
            else
                Console.WriteLine($"移动失败压缩包到 {nameof(_ErrorPath)} 失败：{onlyOne.FullName} -> {dest}");
            return;
        }

        var destDir = GetUniquePath(Path.Combine(_ErrorPath, commonPrefix!));
        foreach (var archiveFile in archiveFiles)
        {
            var dest = Path.Combine(destDir, archiveFile.Name);
            if (TryMoveEntry(archiveFile, dest))
                Console.WriteLine("已移动失败压缩包到：" + dest);
            else
                Console.WriteLine($"移动失败压缩包到 {nameof(_ErrorPath)} 失败：{archiveFile.FullName} -> {dest}");
        }
    }

    private static void MoveResultsToComplete(IReadOnlyCollection<DirectoryInfo> resultDirs, string destPath)
    {
        foreach (var resultDir in resultDirs)
        {
            var dest = GetUniquePath(Path.Combine(destPath, resultDir.Name));
            if (TryMoveEntry(resultDir, dest))
                Console.WriteLine("已移动解压结果到：" + dest);
            else
                Console.WriteLine($"移动解压结果到 {nameof(_CompletePath)} 失败：{resultDir.FullName} -> {dest}");
        }
    }

    private static void PromptCleanupIntermediates(IReadOnlyCollection<FileInfo> intermediates, long extractedTotalSize)
    {
        // 自动清理判定：若解压后的总大小 >= 源文件大小的 50%，则直接清理
        var originalArchivesSize = GetFilesSize(intermediates);
        var autoClean = originalArchivesSize > 0 && extractedTotalSize >= originalArchivesSize / 2;
        var clean = autoClean;

        if (!autoClean)
        {
            // 等待用户输入
            Console.WriteLine("\n解压后大小不到源文件的50%，请检查是否有文件未解压：");
            foreach (var intermediate in intermediates)
                Console.WriteLine($" - {intermediate.FullName} ({intermediate.Length} 字节)");
            Console.WriteLine("\n按回车键删除所有中间压缩包，按其他键保留所有文件...");
            var key = Console.ReadKey(true);
            if (key.Key is ConsoleKey.Enter)
                clean = true;
        }

        if (clean)
        {
            Console.WriteLine("\n解压后大小超过源文件的50%，自动清理中间压缩包...");
            Console.WriteLine("正在清理中间文件...");

            foreach (var archive in intermediates)
                if (archive.FullName.StartsWith(_TempPath) && RemoveIfExists(archive))
                    Console.WriteLine($"已删除：{archive.FullName}");
            _ = CleanEmptyDirectories(_TempPath, true);
            Console.WriteLine("清理完成！");
        }
        else
        {
            Console.WriteLine("\n已保留所有文件。");
        }
    }

    public static int CleanEmptyDirectories(string root, bool includeRoot)
    {
        return CleanEmptyDirectories(new DirectoryInfo(root), includeRoot);
    }

    public static int CleanEmptyDirectories(DirectoryInfo rootDir, bool includeRoot)
    {
        var count = CleanEmptyDirectoriesInternal(rootDir, includeRoot);
        Console.WriteLine($"已删除 {rootDir.FullName} 内空文件夹 {count} 个。");
        return count;
    }

    public static int CleanEmptyDirectoriesInternal(DirectoryInfo root, bool includeRoot)
    {
        var count = 0;
        try
        {
            count += root.GetDirectories().Sum(dir => CleanEmptyDirectoriesInternal(dir, true));
            if (includeRoot)
                if (root.GetFileSystemInfos().Length is 0)
                {
                    root.Delete();
                    count++;
                }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"清理 {root} 空文件夹时出错：{ex.Message}");
        }

        return count;
    }

    /// <summary>
    /// 在目录内最多 <paramref name="maxDepth"/> 次，将“父仅有唯一子目录且两者同名”的相邻两层合并。
    /// </summary>
    private static void NormalizeRedundantNestedFolders(DirectoryInfo rootDir, int maxDepth = 3)
    {
        if (!rootDir.Exists)
            return;

        for (var i = 0; i < maxDepth - 1; i++)
            try
            {
                var entries = rootDir.GetFileSystemInfos("*", SearchOption.TopDirectoryOnly);


                switch (entries)
                {
                    case { Length: > 1 }:
                    {
                        foreach (var entry in entries)
                            if (entry is DirectoryInfo info)
                                NormalizeRedundantNestedFolders(info, maxDepth - 1);
                        break;
                    }
                    case []:
                    {
                        rootDir.Delete(false);
                        break;
                    }
                    case [var onlyChild]:
                    {
                        if (rootDir.Parent is { } parent && onlyChild.Name.Contains(rootDir.Name))
                        {
                            MoveEntryToDirectoryAndMerge(onlyChild.FullName, parent.FullName);

                            if (onlyChild.Name != rootDir.Name)
                                rootDir.Delete();
                        }

                        break;
                    }
                }
            }
            catch
            {
            }
    }

    /// <summary>
    /// 将<paramref name="sourcePath"/>移动到<paramref name="destDir"/>目录下，若存在则空目录则合并
    /// </summary>
    /// <remarks>
    /// <paramref name="sourcePath"/>可以是<paramref name="destDir"/>的孙项目（a/b/c.txt -> a = a/c.txt）
    /// </remarks>
    /// <param name="sourcePath"></param>
    /// <param name="destDir">目标目录</param>
    private static void MoveEntryToDirectoryAndMerge(string sourcePath, string destDir)
    {
        var sourceDir = Path.GetDirectoryName(sourcePath);
        if (string.IsNullOrEmpty(sourceDir))
            return;
        var destPath = Path.Combine(destDir, Path.GetFileName(sourcePath));

        try
        {
            if (Directory.Exists(destPath))
            {
                var entries = Directory.GetFileSystemEntries(destPath);
                // 一个子项且子项是sourcePath，或空目录，允许合并
                if (!((entries is [var onlyChild] && onlyChild == sourcePath)
                      || entries.Length is 0))
                {
                    Console.WriteLine($"目标目录已存在且非空，无法合并：{destPath}");
                    return;
                }

                // 移出子项，删除空目录，再移动
                var newSourcePath = GetUniquePath(destPath);
                Directory.Move(sourcePath, newSourcePath);
                sourcePath = newSourcePath;
                Directory.Delete(destPath);
            }

            Directory.Move(sourcePath, destPath);
        }
        catch (Exception e)
        {
            Console.WriteLine($"移动目录内容失败（{sourcePath} -> {destPath}）：{e.Message}");
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
                var destFile = Path.Combine(_TempPath, Path.ChangeExtension(source.Name, newExt));
                var newSourcePath = GetUniquePath(destFile);
                _ = TryMoveEntry(source, newSourcePath);
                source = new(newSourcePath);
            }

            var list = FindSiblingSplitVolumes(source, out commonPrefix);

            if (list.Count is 0)
                throw new ArgumentException(source.Name);

            if (list is [var onlyVolume])
            {
                var destPathSingle = GetUniquePath(Path.Combine(_TempPath, onlyVolume.Name));
                TryMoveEntry(source, destPathSingle);
                return [new(destPathSingle)];
            }

            var tempDestDir = GetUniquePath(Path.Combine(_TempPath, commonPrefix!));
            var result = new List<FileInfo>();

            foreach (var volume in list)
            {
                var destPath = Path.Combine(tempDestDir, volume.Name);
                TryMoveEntry(volume, destPath);
                result.Add(new(destPath));
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
        var fileName = file.Name.ToLower();
        var (ext1, ext2) = GetLastTwoExtensions(fileName);

        // 文本文件且小于1MB
        if (ext2 is "txt" or "md" && file.Length < 1 << 20)
            return FileCompressionType.Text;

        // part1.rar/part2.rar
        if (ext2 is "rar" && ext1 is ['p', 'a', 'r', 't', .. var idx] && int.TryParse(idx, out var result1))
            return result1 > 1 ? FileCompressionType.Volume : FileCompressionType.First;

        // 检查是否为已知格式或第一卷
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
            ['z', .. { Length: >= 2 }] => FileCompressionType.Volume,
            // r00/r01
            ['r', .. { Length: >= 2 } remains] when int.TryParse(remains, out var result2) =>
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
    /// 
    /// </summary>
    /// <param name="name"></param>
    /// <returns>没有'.'</returns>
    private static (string? Ext1, string? Ext2) GetLastTwoExtensions(string name)
    {
        var (nameWithoutExt, ext) = GetExtension(name);
        return (Path.GetExtension(nameWithoutExt), ext);
    }

    private static (string Name, string? Ext) GetExtension(string name)
    {
        var lastIndex = name.LastIndexOf('.');
        return lastIndex < 0
            ? (name, null)
            : (name[..lastIndex], name[(lastIndex + 1)..]);
    }

    private static string GetUniquePath(string path)
    {
        if (!File.Exists(path) && !Directory.Exists(path))
            return path;
        var dir = Path.GetDirectoryName(path)!;
        var name = Path.GetFileNameWithoutExtension(path);
        var ext = Path.GetExtension(path);
        var i = 1;
        string candidate;
        do
        {
            candidate = Path.Combine(dir, $"{name} ({i}){ext}");
            i++;
        } while (File.Exists(candidate) || Directory.Exists(candidate));
        return candidate;
    }

    private static bool RemoveIfExists(FileSystemInfo info)
    {
        try
        {
            if (info.Exists)
            {
                info.Delete();
                return true;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"删除失败 {info.FullName}：{ex.Message}");
        }

        return false;
    }

    /// <summary>
    /// 将移动<paramref name="source"/>为<paramref name="destPath"/>
    /// </summary>
    /// <param name="source"></param>
    /// <param name="destPath">目标位置</param>
    private static bool TryMoveEntry(FileSystemInfo source, string destPath)
    {
        try
        {
            if (Path.GetDirectoryName(destPath) is { } directory)
                Directory.CreateDirectory(directory);
            switch (source)
            {
                case DirectoryInfo directoryInfo:
                    directoryInfo.MoveTo(destPath);
                    break;
                case FileInfo fileInfo:
                    fileInfo.MoveTo(destPath);
                    break;
                default:
                    throw new ArgumentException("source must be DirectoryInfo or FileInfo");
            }

            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"移动失败（{source.FullName} -> {destPath}）：{ex.Message}");
            return false;
        }
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
        var name = firstVolume.Name.ToLower();

        // Name 分为 nameWithoutExt.ext1.ext2
        var (nameWithoutExt2, nameExt2) = GetExtension(name);
        var (nameWithoutExt12, nameExt1) = GetExtension(nameWithoutExt2);

        var allFiles = dir.GetFiles("*", SearchOption.TopDirectoryOnly);

        var results = new SortedList<int, FileInfo>();

        foreach (var file in allFiles)
        {
            var (withoutExt2, ext2) = GetExtension(file.Name.ToLower());
            switch (nameExt1, nameExt2)
            {
                // 1) name.part1.rar → name.partN.rar
                case (['p', 'a', 'r', 't', ..], "rar"):
                {
                    commonPrefix = nameWithoutExt12;
                    var (withoutExt12, ext1) = GetExtension(withoutExt2);
                    if (withoutExt12 == commonPrefix
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
                    if (withoutExt2 == commonPrefix
                        && ext2 is ['r', .. { Length: >= 2 } rm]
                        && int.TryParse(rm, out var i))
                        results[i] = file;
                    break;
                }
                // 3) name.zip.001 / name.zipx.001 / name.7z.001 → name.ext.00N
                case ("7z" or "zip" or "zipx", _) when int.TryParse(nameExt2, out _):
                {
                    commonPrefix = nameWithoutExt2;
                    // name.7z == name.7z
                    if (withoutExt2 == commonPrefix && int.TryParse(ext2, out var i))
                        results[i] = file;
                    break;
                }
                // 4) name.zip / name.zipx → name.zNN
                case (_, "zip" or "zipx"):
                {
                    commonPrefix = nameWithoutExt2;
                    results[0] = firstVolume; // 第一卷不会遍历到，手动添加
                    if (withoutExt2 == nameWithoutExt2
                        && ext2 is ['z', .. { Length: >= 2 } remains]
                        && int.TryParse(remains, out var i))
                        results[i] = file;
                    break;
                }
            }
        }

        return results.Values;
    }
}

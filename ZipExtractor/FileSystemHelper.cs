using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using Microsoft.VisualBasic.FileIO;
using ZipExtractor.Logging;
using SearchOption = System.IO.SearchOption;

namespace ZipExtractor;

public static class FileSystemHelper
{
    private static void Log(string message) => ConsoleLog.Default.Detail("整理", message);

    private static void Log(Exception ex, string message) => ConsoleLog.Default.Warning($"{message}：{ex.Message}");

    /// <summary>
    /// 获取文件名的最后两个扩展名
    /// </summary>
    /// <param name = "name"></param>
    /// <returns>没有'.'</returns>
    public static (string? Ext1, string? Ext2) GetLastTwoExtensions(string name)
    {
        var (nameWithoutExt, ext) = GetExtension(name);
        return (GetExtension(nameWithoutExt).Ext, ext);
    }

    /// <summary>
    /// 获取文件名和扩展名
    /// </summary>
    /// <param name = "name"></param>
    /// <returns>没有'.'</returns>
    public static (string Name, string? Ext) GetExtension(string name)
    {
        var lastIndex = name.LastIndexOf('.');
        return lastIndex < 0 ? (name, null) : (name[..lastIndex], name[(lastIndex + 1)..]);
    }

    /// <summary>
    /// 获取在指定<paramref name = "path"/>不重名的文件或目录路径
    /// </summary>
    /// <param name = "path"></param>
    /// <param name = "updateDirectoryLastWriteTime"></param>
    /// <returns></returns>
    public static string GetUniquePath(string path, bool updateDirectoryLastWriteTime = false)
    {
        if (!File.Exists(path))
        {
            if (!Directory.Exists(path))
                return path;
            // 尝试清理空目录
            _ = CleanEmptyDirectories(path, true, updateDirectoryLastWriteTime);
            if (!Directory.Exists(path))
                return path;
        }

        if (!Exists(path))
            return path;
        var dir = Path.GetDirectoryName(path)!;
        var name = Path.GetFileNameWithoutExtension(path);
        var ext = Path.GetExtension(path);
        var i = 1;
        string candidate;
        // 如果name已经是“xxx(n)”格式，name = xxx
        if (name.EndsWith(')') && name.LastIndexOf(" (", StringComparison.Ordinal) is var leftParenIndex and >= 0)
        {
            var numberPart = name[(leftParenIndex + 2)..^1];
            if (int.TryParse(numberPart, out _))
                name = name[..leftParenIndex];
        }

        do
        {
            candidate = Path.Combine(dir, $"{name} ({i}){ext}");
            i++;
        }
        while (Exists(candidate));
        return candidate;
    }

    /// <summary>Deletes a directory, optionally allowing its parent's modification date to change.</summary>
    public static void DeleteDirectory(string path, bool recursive = false, bool updateDirectoryLastWriteTime = false)
    {
        using var dates = new DirectoryTimestampScope(updateDirectoryLastWriteTime, path, Path.GetDirectoryName(Path.GetFullPath(path)));
        Directory.Delete(path, recursive);
    }

    /// <summary>Deletes a file, optionally allowing its directory's modification date to change.</summary>
    public static void DeleteFile(string path, bool updateDirectoryLastWriteTime = false)
    {
        using var dates = new DirectoryTimestampScope(updateDirectoryLastWriteTime, Path.GetDirectoryName(Path.GetFullPath(path)));
        File.Delete(path);
    }

    public static long GetEntrySize(FileSystemInfo entry)
    {
        return entry.Exists ? entry is FileInfo fileInfo ? fileInfo.Length : entry is DirectoryInfo directoryInfo ? GetDirectorySize(directoryInfo) : 0 : 0;
    }

    /// <summary>
    /// 获取目录总大小
    /// </summary>
    /// <param name = "dir"></param>
    /// <returns></returns>
    public static long GetDirectorySize(DirectoryInfo dir)
    {
        return dir.Exists ? GetFilesSize(dir.GetFiles("*", SearchOption.AllDirectories)) : 0;
    }

    /// <summary>
    /// 获取文件集合大小
    /// </summary>
    /// <param name = "files"></param>
    /// <returns></returns>
    public static long GetFilesSize(IEnumerable<FileInfo> files)
    {
        long sum = 0;
        foreach (var file in files)
            try
            {
                sum += file.Length;
            }
            catch
            {
                // ignored
            }

        return sum;
    }

    /// <inheritdoc cref = "CleanEmptyDirectoriesInternal"/>
    public static int CleanEmptyDirectories(string root, bool includeRoot, bool updateDirectoryLastWriteTime = false)
    {
        return CleanEmptyDirectories(new DirectoryInfo(root), includeRoot, updateDirectoryLastWriteTime);
    }

    /// <inheritdoc cref = "CleanEmptyDirectoriesInternal"/>
    public static int CleanEmptyDirectories(DirectoryInfo root, bool includeRoot, bool updateDirectoryLastWriteTime = false)
    {
        var count = CleanEmptyDirectoriesInternal(root, includeRoot, updateDirectoryLastWriteTime);
        if (count is not 0)
            Log($"已删除 {root.FullName} 内空文件夹 {count} 个。");
        return count;
    }

    /// <summary>
    /// 递归清理<paramref name = "root"/>的空文件夹
    /// </summary>
    /// <param name = "root"></param>
    /// <param name = "includeRoot"><paramref name = "root"/>也为空时是否删除</param>
    /// <returns>清理的空文件夹数量</returns>
    private static int CleanEmptyDirectoriesInternal(DirectoryInfo root, bool includeRoot, bool updateDirectoryLastWriteTime)
    {
        if (!root.Exists)
            return 0;
        if (IsGameFolder(root))
            return 0;
        var count = 0;
        try
        {
            count += root.GetDirectories().Sum(dir => CleanEmptyDirectoriesInternal(dir, true, updateDirectoryLastWriteTime));
            if (includeRoot)
                if (root.GetFileSystemInfos().Length is 0)
                {
                    DeleteDirectory(root.FullName, updateDirectoryLastWriteTime: updateDirectoryLastWriteTime);
                    count++;
                }
        }
        catch (Exception ex)
        {
            Log(ex, $"清理 {root} 空文件夹时出错");
        }

        return count;
    }

    /// <summary>
    /// 判断是否重复的阈值
    /// </summary>
    public enum RedundantThreshold
    {
        /// <summary>
        /// 只有父子名称相同时合并
        /// </summary>
        Equal,
        /// <summary>
        /// 总是合并
        /// </summary>
        Always,
        /// <summary>
        /// 子名称包含父名称时合并
        /// </summary>
        ContainsParent,
        /// <summary>
        /// 父名称包含子名称时合并
        /// </summary>
        ContainsChild
    }

    /// <summary>
    /// 判断当前文件夹是否需要和它的唯一子项合并。
    /// </summary>
    /// <param name = "current">当前文件夹</param>
    /// <param name = "onlyChild">当前文件夹的唯一子项</param>
    /// <param name = "mergedNameWithoutExtension">合并后的目标基础名；当唯一子项为文件时会自动保留原扩展名</param>
    /// <returns>是否合并</returns>
    public delegate bool ShouldMergeNestedFolder(DirectoryInfo current, FileSystemInfo onlyChild, out string mergedNameWithoutExtension);

    /// <summary>
    /// 在目录内最多 <paramref name = "maxDepth"/> 次，将“父仅有唯一子项，且当前文件夹/唯一子项符合 <paramref name = "shouldMerge"/>”的相邻两层合并。空文件夹直接删除
    /// </summary>
    /// <param name = "root">根目录不会被合并/删除</param>
    /// <param name = "shouldMerge">判断当前文件夹是否需要和它的唯一子项合并，参数为当前文件夹和它的唯一子项</param>
    /// <param name = "maxDepth">
    /// 最小有效值是1，表示仅处理<paramref name = "root"/>与其唯一子项的合并。
    /// 每增加1，表示允许递归继续向下处理一层子目录。
    /// </param>
    /// <param name = "useUniqueName">若子项移动后重名，是否重命名，否则放弃移动</param>
    public static void NormalizeMatchedNestedFolders(DirectoryInfo root, ShouldMergeNestedFolder shouldMerge, int maxDepth = 3, bool useUniqueName = false, bool updateDirectoryLastWriteTime = false)
    {
        ArgumentNullException.ThrowIfNull(shouldMerge);
        if (IsGameFolder(root))
            return;
        _ = CleanEmptyDirectories(root, false, updateDirectoryLastWriteTime);
        foreach (var directoryInfo in root.GetDirectories())
            NormalizeMatchedNestedFoldersCore(directoryInfo, maxDepth, shouldMerge, useUniqueName, updateDirectoryLastWriteTime, out _);
    }

    /// <summary>
    /// 在目录内最多 <paramref name = "maxDepth"/> 次，将“名称包含 <paramref name = "namePart"/>，且符合合并标准”的文件夹与相邻层合并。合并后使用另一侧名称
    /// </summary>
    /// <param name = "root">根目录不会被合并/删除</param>
    /// <param name = "namePart">当前文件夹名称需要包含的字符串，如“新建文件夹”</param>
    /// <param name = "maxDepth">
    /// 最小有效值是1，表示仅处理<paramref name = "root"/>与其唯一子项的合并。
    /// 每增加1，表示允许递归继续向下处理一层子目录。
    /// </param>
    /// <param name = "useUniqueName">若子项移动后重名，是否重命名，否则放弃移动</param>
    public static void NormalizeNestedFoldersContainingName(DirectoryInfo root, string namePart, int maxDepth = 3, bool useUniqueName = false, bool updateDirectoryLastWriteTime = false)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(namePart);
        NormalizeMatchedNestedFolders(root, (current, onlyChild, out mergedNameWithoutExtension) =>
        {
            var (childName, _) = GetMergedChildName(onlyChild);
            var currentMatches = current.Name.Contains(namePart, StringComparison.OrdinalIgnoreCase);
            if (currentMatches)
            {
                mergedNameWithoutExtension = childName;
                return true;
            }

            mergedNameWithoutExtension = current.Name;
            var childMatches = onlyChild is DirectoryInfo childDirectory && childDirectory.Name.Contains(namePart, StringComparison.OrdinalIgnoreCase);
            return childMatches;
        }, maxDepth, useUniqueName, updateDirectoryLastWriteTime);
    }

    /// <summary>
    /// 在目录内最多 <paramref name = "maxDepth"/> 次，将“父仅有唯一子项，且两者名称重复”的相邻两层合并。空文件夹直接删除
    /// </summary>
    /// <param name = "root">根目录不会被合并/删除</param>
    /// <param name = "maxDepth">
    /// 最小有效值是1，表示仅处理<paramref name = "root"/>与其唯一子项的合并。
    /// 每增加1，表示允许递归继续向下处理一层子目录。
    /// </param>
    /// <param name = "strict"></param>
    /// <param name = "useParentName">合并后原子项是否采用父文件夹的名字</param>
    /// <param name = "useUniqueName">若子项移动后重名，是否重命名，否则放弃移动</param>
    public static void NormalizeRedundantNestedFolders(DirectoryInfo root, int maxDepth = 3, RedundantThreshold strict = RedundantThreshold.ContainsParent, bool useParentName = false, bool useUniqueName = false, bool updateDirectoryLastWriteTime = false)
    {
        NormalizeMatchedNestedFolders(root, (current, onlyChild, out mergedNameWithoutExtension) =>
        {
            var (childName, _) = GetMergedChildName(onlyChild);
            mergedNameWithoutExtension = useParentName ? current.Name : childName;
            return strict switch
            {
                RedundantThreshold.Equal => childName.EqualsFileName(current.Name),
                RedundantThreshold.Always => true,
                RedundantThreshold.ContainsParent => childName.ContainsFileName(current.Name),
                RedundantThreshold.ContainsChild => current.Name.ContainsFileName(childName),
                _ => false
            };
        }, maxDepth, useUniqueName, updateDirectoryLastWriteTime);
    }

    /// <inheritdoc cref = "NormalizeMatchedNestedFolders(DirectoryInfo, ShouldMergeNestedFolder, int, bool, bool)"/>
    /// <param name = "current">当前目录</param>
    /// <param name = "newCurrent">若当前目录被合并，则返回新的代替项；若被删除则返回<see langword="null"/></param>
    [SuppressMessage("ReSharper", "InvalidXmlDocComment")]
    private static void NormalizeMatchedNestedFoldersCore(DirectoryInfo current, int maxDepth, ShouldMergeNestedFolder shouldMerge, bool useUniqueName, bool updateDirectoryLastWriteTime, out FileSystemInfo? newCurrent)
    {
        newCurrent = current;
        if (maxDepth <= 0 || !current.Exists)
            return;
        try
        {
            var entries = current.GetFileSystemInfos("*", SearchOption.TopDirectoryOnly);
            // Protect the folder itself as well as its entire subtree, even with only one file.
            if (IsGameFolder(entries))
            {
                ConsoleLog.Default.GameSkipped(current.FullName);
                return;
            }

            switch (entries)
            {
                // current -> a, b, c, ...
                case { Length: > 1 }:
                {
                    foreach (var directory in entries.OfType<DirectoryInfo>())
                        NormalizeMatchedNestedFoldersCore(directory, maxDepth - 1, shouldMerge, useUniqueName, updateDirectoryLastWriteTime, out _);
                    break;
                }

                // current -> 空
                case []:
                {
                    DeleteDirectory(current.FullName, updateDirectoryLastWriteTime: updateDirectoryLastWriteTime);
                    newCurrent = null;
                    break;
                }

                // current -> onlyChild
                case [var onlyChild]:
                {
                    if (onlyChild is DirectoryInfo directoryInfo)
                    {
                        NormalizeMatchedNestedFoldersCore(directoryInfo, maxDepth - 1, shouldMerge, useUniqueName, updateDirectoryLastWriteTime, out onlyChild);
                        // A duplicate-name wrapper can be removed by moving the game folder as a whole.
                        // Keep its contents and internal layout protected from normalization.
                        if (onlyChild is DirectoryInfo resultDirectory && IsGameFolder(resultDirectory) && !current.Name.EqualsFileName(resultDirectory.Name))
                            return;
                    }

                    if (onlyChild is null)
                    {
                        if (current.Exists && current.GetFileSystemInfos("*", SearchOption.TopDirectoryOnly).Length is 0)
                        {
                            DeleteDirectory(current.FullName, updateDirectoryLastWriteTime: updateDirectoryLastWriteTime);
                            newCurrent = null;
                        }

                        break;
                    }

                    // parent -> current -> onlyChild
                    if (current.Parent is { } parent)
                    {
                        // childName:
                        // File -> Name without extension
                        // Directory -> Name
                        var (_, ext) = GetMergedChildName(onlyChild);
                        var merge = shouldMerge(current, onlyChild, out var mergedNameWithoutExtension);
                        if (merge)
                        {
                            var originalPath = onlyChild.FullName;
                            var suggestedDestPath = parent.Combine(mergedNameWithoutExtension).GetComparablePath();
                            suggestedDestPath += ext is null ? null : "." + ext;
                            if (onlyChild.MoveToDirectoryAndMerge(suggestedDestPath, useUniqueName, updateDirectoryLastWriteTime))
                            {
                                Log($"合并重复文件夹：({originalPath}, {current.FullName}) -> {onlyChild.FullName}");
                                // 若成功移动，则此时onlyChild已和current在同一目录下
                                // 移出OnlyChild后，root变为空目录，删除它
                                // 判断名称是否相等，以防删除onlyChild
                                if (!onlyChild.Name.EqualsFileName(current.Name))
                                    DeleteDirectory(current.FullName, updateDirectoryLastWriteTime: updateDirectoryLastWriteTime);
                                newCurrent = onlyChild;
                            }
                        }
                    }

                    break;
                }
            }
        }
        catch
        {
            // ignored
        }
    }

    /// <summary>Direct EXE/DLL/BIN children identify a game folder; protect its entire subtree.</summary>
    public static bool IsGameFolder(DirectoryInfo directory) => directory.Exists && IsGameFolder(directory.EnumerateFileSystemInfos());

    public static bool IsGameFolder(IEnumerable<FileSystemInfo> entries)
    {
        return entries.OfType<FileInfo>().Any(t => t.Extension.Equals(".dll", StringComparison.OrdinalIgnoreCase) || t.Extension.Equals(".exe", StringComparison.OrdinalIgnoreCase) || t.Extension.Equals(".bin", StringComparison.OrdinalIgnoreCase));
    }

    private static (string Name, string? Extension) GetMergedChildName(FileSystemInfo entry)
    {
        return entry is FileInfo file ? GetExtension(file.Name) : (entry.Name, null);
    }

    /// <param name = "info"></param>
    extension(FileInfo info)
    {
        /// <summary>
        /// 删除<paramref name = "info"/>，如果不存在则不操作
        /// </summary>
        /// <returns>是否成功删除</returns>
        public bool RemoveIfExists(bool moveToRecycleBin = false, bool updateDirectoryLastWriteTime = false)
        {
            using var dates = new DirectoryTimestampScope(updateDirectoryLastWriteTime, info.Parent?.FullName);
            try
            {
                if (info.Exists)
                {
                    if (moveToRecycleBin)
                        FileSystem.DeleteFile(info.FullName, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
                    else
                        info.Delete();
                    return true;
                }
            }
            catch (Exception ex)
            {
                Log(ex, $"删除失败 {info.FullName}");
            }

            return false;
        }
    }

    /// <param name = "info"></param>
    extension(DirectoryInfo info)
    {
        public string Combine(FileInfo file) => Path.Combine(info.FullName, file.Name);
        public string Combine(string fileName) => Path.Combine(info.FullName, fileName);
        /// <summary>
        /// 删除<paramref name = "info"/>，如果不存在则不操作
        /// </summary>
        /// <returns>是否成功删除</returns>
        public bool RemoveIfExists(bool recursive, bool moveToRecycleBin = false, bool updateDirectoryLastWriteTime = false)
        {
            using var dates = new DirectoryTimestampScope(updateDirectoryLastWriteTime, info.Parent?.FullName);
            try
            {
                if (info.Exists)
                {
                    if (moveToRecycleBin)
                        FileSystem.DeleteDirectory(info.FullName, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
                    else
                        info.Delete(recursive);
                    return true;
                }
            }
            catch (Exception ex)
            {
                Log(ex, $"删除失败 {info.FullName}");
            }

            return false;
        }
    }

    /// <param name = "info"></param>
    extension(FileSystemInfo info)
    {
        public string NameWithoutExtension => Path.GetFileNameWithoutExtension(info.Name);
        /// <inheritdoc cref = "DirectoryInfo.Parent"/>
        public DirectoryInfo Parent => info switch
        {
            DirectoryInfo dirInfo => dirInfo.Parent,
            FileInfo fileInfo => fileInfo.Directory,
            _ => null
        } ?? throw new InvalidOperationException("FileSystemInfo has no parent");

        /// <inheritdoc cref = "DirectoryInfo.MoveTo"/>
        public void MoveTo(string destPath, bool updateDirectoryLastWriteTime = false)
        {
            destPath = Path.GetFullPath(destPath);
            if (destPath.EqualsFileName(info.FullName))
                return;
            using var dates = new DirectoryTimestampScope(updateDirectoryLastWriteTime, info.Parent?.FullName, Path.GetDirectoryName(destPath));
            switch (info)
            {
                case DirectoryInfo dirInfo:
                    var modified = Directory.GetLastWriteTimeUtc(dirInfo.FullName);
                    dirInfo.MoveTo(destPath);
                    Directory.SetLastWriteTimeUtc(destPath, updateDirectoryLastWriteTime ? DateTime.UtcNow : modified);
                    break;
                case FileInfo fileInfo:
                    fileInfo.MoveTo(destPath);
                    break;
                default:
                    throw new ArgumentException($"source must be {nameof(DirectoryInfo)} or {nameof(FileInfo)}");
            }
        }

        /// <inheritdoc cref = "Directory.Exists"/>
        public static bool Exists(string path) => File.Exists(path) || Directory.Exists(path);
        /// <summary>
        /// 将移动<paramref name = "info"/>为<paramref name = "destPath"/>
        /// </summary>
        /// <param name = "destPath">目标位置</param>
        public bool TryMoveTo(string destPath, bool updateDirectoryLastWriteTime = false)
        {
            try
            {
                if (Path.GetDirectoryName(destPath) is { } directory)
                    _ = Directory.CreateDirectory(directory);
                info.MoveTo(destPath, updateDirectoryLastWriteTime);
                return true;
            }
            catch (Exception ex)
            {
                Log(ex, $"移动失败（{info.FullName} -> {destPath}）");
                return false;
            }
        }

        /// <summary>
        /// 将<paramref name = "info"/>移动到<paramref name = "suggestedDestPath"/>，若存在则空目录则合并
        /// </summary>
        /// <remarks>
        /// <paramref name = "info"/>可以是<paramref name = "suggestedDestPath"/>的子项目（a/b/c.txt -> a = a/c.txt）
        /// </remarks>
        /// <param name = "suggestedDestPath">目标位置</param>
        /// <param name = "renameWhenDuplicated">当无法删除目标同名项目的时候，重命名<paramref name = "suggestedDestPath"/></param>
        private bool MoveToDirectoryAndMerge(string suggestedDestPath, bool renameWhenDuplicated, bool updateDirectoryLastWriteTime)
        {
            // 异常目录
            if (info.Parent is not { Exists: true })
                return false;
            // 本身已在目标位置
            if (suggestedDestPath.EqualsFileName(info.FullName))
                return true;
            try
            {
                if (File.Exists(suggestedDestPath))
                    throw new InvalidOperationException("目标位置有同名文件");
                if (Directory.Exists(suggestedDestPath))
                {
                    var entries = Directory.GetFileSystemEntries(suggestedDestPath);
                    // 一个子项且子项是source
                    if (entries is [var onlyChild] && onlyChild.EqualsFileName(info.FullName))
                    {
                        // 移出子项，变成空目录
                        info.MoveTo(GetUniquePath(suggestedDestPath, updateDirectoryLastWriteTime), updateDirectoryLastWriteTime);
                        DeleteDirectory(suggestedDestPath, updateDirectoryLastWriteTime: updateDirectoryLastWriteTime);
                    }
                    else if (entries.Length is not 0)
                        throw new InvalidOperationException("目标位置已存在非空目录");
                }

                info.MoveTo(suggestedDestPath, updateDirectoryLastWriteTime);
                return true;
            }
            catch (Exception e)
            {
                if (renameWhenDuplicated)
                {
                    info.MoveTo(GetUniquePath(suggestedDestPath, updateDirectoryLastWriteTime), updateDirectoryLastWriteTime);
                    return true;
                }

                Log($"移动目录内容失败（{info.FullName} -> {suggestedDestPath}）：{e.Message}");
                return false;
            }
        }

        public bool MoveRelativePathTo(string toPath, string fromPath = ExtractorBase.TempPath, bool moveToUniquePath = true, bool updateDirectoryLastWriteTime = false)
        {
            var relativePath = Path.GetRelativePath(fromPath, info.FullName);
            var dest = Path.Combine(toPath, relativePath);
            if (moveToUniquePath)
                dest = GetUniquePath(dest, updateDirectoryLastWriteTime);
            return info.TryMoveTo(dest, updateDirectoryLastWriteTime);
        }
    }

    extension(string str)
    {
        /// <inheritdoc cref = "string.Equals(string, StringComparison)"/>
        public bool EqualsFileName(string value) => str.GetComparablePath().Equals(value.GetComparablePath(), StringComparison.OrdinalIgnoreCase);
        /// <inheritdoc cref = "string.Contains(string, StringComparison)"/>
        public bool ContainsFileName(string value) => str.GetComparablePath().Contains(value.GetComparablePath(), StringComparison.OrdinalIgnoreCase);
        /// <inheritdoc cref = "string.StartsWith(string, StringComparison)"/>
        public bool StartWithFileName(string value) => str.GetComparablePath().StartsWith(value.GetComparablePath(), StringComparison.OrdinalIgnoreCase);
        private string GetComparablePath() => Path.GetFullPath(str).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }
}

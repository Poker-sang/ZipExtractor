using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ZipExtractor;

public static class FileSystemHelper
{
    /// <summary>
    /// 获取文件名的最后两个扩展名
    /// </summary>
    /// <param name="name"></param>
    /// <returns>没有'.'</returns>
    public static (string? Ext1, string? Ext2) GetLastTwoExtensions(string name)
    {
        var (nameWithoutExt, ext) = GetExtension(name);
        return (GetExtension(nameWithoutExt).Ext, ext);
    }

    /// <summary>
    /// 获取文件名和扩展名
    /// </summary>
    /// <param name="name"></param>
    /// <returns>没有'.'</returns>
    public static (string Name, string? Ext) GetExtension(string name)
    {
        var lastIndex = name.LastIndexOf('.');
        return lastIndex < 0
            ? (name, null)
            : (name[..lastIndex], name[(lastIndex + 1)..]);
    }

    /// <summary>
    /// 获取在指定<paramref name="path"/>不重名的文件或目录路径
    /// </summary>
    /// <param name="path"></param>
    /// <returns></returns>
    public static string GetUniquePath(string path)
    {
        if (!FileSystemInfo.Exists(path))
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
        } while (FileSystemInfo.Exists(candidate));
        return candidate;
    }

    /// <summary>
    /// 删除<paramref name="info"/>，如果不存在则不操作
    /// </summary>
    /// <param name="info"></param>
    /// <returns>是否成功删除</returns>
    public static bool RemoveIfExists(FileSystemInfo info)
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
    /// 获取目录总大小
    /// </summary>
    /// <param name="dir"></param>
    /// <returns></returns>
    public static long GetDirectorySize(DirectoryInfo dir)
    {
        return dir.Exists ? GetFilesSize(dir.GetFiles("*", SearchOption.AllDirectories)) : 0;
    }

    /// <summary>
    /// 获取文件集合大小
    /// </summary>
    /// <param name="files"></param>
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

    /// <summary>
    /// 将移动<paramref name="source"/>为<paramref name="destPath"/>
    /// </summary>
    /// <param name="source"></param>
    /// <param name="destPath">目标位置</param>
    public static bool TryMoveEntry(FileSystemInfo source, string destPath)
    {
        try
        {
            if (Path.GetDirectoryName(destPath) is { } directory)
                _ = Directory.CreateDirectory(directory);
            source.MoveTo(destPath);
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"移动失败（{source.FullName} -> {destPath}）：{ex.Message}");
            return false;
        }
    }

    /// <inheritdoc cref="CleanEmptyDirectoriesInternal" />
    public static int CleanEmptyDirectories(string root, bool includeRoot)
    {
        return CleanEmptyDirectories(new DirectoryInfo(root), includeRoot);
    }

    /// <inheritdoc cref="CleanEmptyDirectoriesInternal" />
    public static int CleanEmptyDirectories(DirectoryInfo root, bool includeRoot)
    {
        var count = CleanEmptyDirectoriesInternal(root, includeRoot);
        if (count is not 0)
            Console.WriteLine($"已删除 {root.FullName} 内空文件夹 {count} 个。");
        return count;
    }

    /// <summary>
    /// 递归清理<paramref name="root"/>的空文件夹
    /// </summary>
    /// <param name="root"></param>
    /// <param name="includeRoot"><paramref name="root"/>也为空时是否删除</param>
    /// <returns>清理的空文件夹数量</returns>
    private static int CleanEmptyDirectoriesInternal(DirectoryInfo root, bool includeRoot)
    {
        if (!root.Exists)
            return 0;

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

    public enum RedundantThreshold
    {
        /// <summary>
        /// 时只有父子名称相同时合并
        /// </summary>
        Equal,

        /// <summary>
        /// 子名称包含父名称时合并
        /// </summary>
        Contains,

        /// <summary>
        /// 总是合并
        /// </summary>
        Always,

        /// <summary>
        /// 总是合并，且重命名目标文件以避免名称冲突
        /// </summary>
        AlwaysWithRename
    }

    /// <summary>
    /// 在目录内最多 <paramref name="maxDepth"/> 次，将“父仅有唯一子项，且两者名称重复”的相邻两层合并。
    /// </summary>
    /// <param name="root"></param>
    /// <param name="maxDepth"></param>
    /// <param name="strict"></param>
    public static void NormalizeRedundantNestedFolders(DirectoryInfo root, int maxDepth = 3, RedundantThreshold strict = RedundantThreshold.Contains)
    {
        if (!root.Exists)
            return;

        for (var i = 0; i < maxDepth - 1; i++)
            try
            {
                var entries = root.GetFileSystemInfos("*", SearchOption.TopDirectoryOnly);

                switch (entries)
                {
                    case { Length: > 1 }:
                    {
                        foreach (var entry in entries)
                            if (entry is DirectoryInfo info)
                                NormalizeRedundantNestedFolders(info, maxDepth - 1, strict);
                        break;
                    }
                    case []:
                    {
                        root.Delete(false);
                        break;
                    }
                    case [var onlyChild]:
                    {
                        if (onlyChild is DirectoryInfo directoryInfo)
                            NormalizeRedundantNestedFolders(directoryInfo, maxDepth - 1, strict);

                        if (root.Parent is { } parent)
                        {
                            var merge = strict switch
                            {
                                RedundantThreshold.Equal => onlyChild.Name.EqualsFileName(root.Name),
                                RedundantThreshold.Contains => onlyChild.Name.ContainsFileName(root.Name),
                                RedundantThreshold.Always or RedundantThreshold.AlwaysWithRename => true,
                                _ => false
                            };

                            if (merge)
                            {
                                Console.WriteLine($"合并重复文件夹：{onlyChild.FullName}, {root.FullName}");

                                MoveEntryToDirectoryAndMerge(onlyChild, parent.FullName, strict is RedundantThreshold.AlwaysWithRename);

                                if (!onlyChild.Name.EqualsFileName(root.Name))
                                    root.Delete();
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

    /// <summary>
    /// 将<paramref name="source"/>移动到<paramref name="destDir"/>目录下，若存在则空目录则合并
    /// </summary>
    /// <remarks>
    /// <paramref name="source"/>可以是<paramref name="destDir"/>的孙项目（a/b/c.txt -> a = a/c.txt）
    /// </remarks>
    /// <param name="source"></param>
    /// <param name="destDir">目标目录</param>
    /// <param name="renameWhenDuplicated">当无法删除目标同名项目的时候，重命名<see cref="destDir"/></param>
    public static void MoveEntryToDirectoryAndMerge(FileSystemInfo source, string destDir, bool renameWhenDuplicated = false)
    {
        var sourceDir = source.Parent;
        if (sourceDir is not { Exists: true })
            return;
        var destPath = Path.Combine(destDir, source.Name);

        try
        {
            while (true)
            {
                if (destPath.EqualsFileName(source.FullName))
                    return;

                if (File.Exists(destPath))
                {
                    if (renameWhenDuplicated)
                    {
                        destPath = GetUniquePath(destPath);
                        continue;
                    }

                    Console.WriteLine($"移动目录内容失败（{source.FullName} -> {destPath}）：目标位置有同名文件");
                    return;
                }

                var dest = Directory.CreateDirectory(destPath);
                if (dest.Exists)
                {
                    var entries = dest.GetFileSystemInfos();
                    // 一个子项且子项是sourcePath，或空目录，允许合并
                    if (!((entries is [var onlyChild] && onlyChild.FullName.EqualsFileName(source.FullName))
                          || entries.Length is 0))
                    {
                        if (renameWhenDuplicated)
                        {
                            destPath = GetUniquePath(destPath);
                            continue;
                        }

                        Console.WriteLine($"目标目录已存在且非空，无法合并：{destPath}");
                        return;
                    }

                    // 移出子项，删除空目录，再移动
                    var newSourcePath = GetUniquePath(destPath);
                    source.MoveTo(newSourcePath);
                    Directory.Delete(destPath);
                }

                source.MoveTo(destPath);
            }
        }
        catch (Exception e)
        {
            Console.WriteLine($"移动目录内容失败（{source.FullName} -> {destPath}）：{e.Message}");
        }
    }

    extension(FileSystemInfo info)
    {
        /// <inheritdoc cref="DirectoryInfo.Parent" />
        public DirectoryInfo? Parent =>
            info switch
            {
                DirectoryInfo dirInfo => dirInfo.Parent,
                FileInfo fileInfo => fileInfo.Directory,
                _ => null
            };

        /// <inheritdoc cref="DirectoryInfo.MoveTo" />
        public void MoveTo(string destPath)
        {
            switch (info)
            {
                case DirectoryInfo dirInfo:
                    dirInfo.MoveTo(destPath);
                    break;
                case FileInfo fileInfo:
                    fileInfo.MoveTo(destPath);
                    break;
                default:
                    throw new ArgumentException($"source must be {nameof(DirectoryInfo)} or {nameof(FileInfo)}");
            }
        }

        /// <inheritdoc cref="Directory.Exists" />
        public static bool Exists(string path) => File.Exists(path) || Directory.Exists(path);
    }

    extension(string str)
    {
        /// <inheritdoc cref="string.Equals(string, StringComparison)" />
        public bool EqualsFileName(string value) => str.GetComparablePath().Equals(value.GetComparablePath(), StringComparison.OrdinalIgnoreCase);

        /// <inheritdoc cref="string.Contains(string, StringComparison)" />
        public bool ContainsFileName(string value) => str.GetComparablePath().Contains(value.GetComparablePath(), StringComparison.OrdinalIgnoreCase);

        /// <inheritdoc cref="string.StartsWith(string, StringComparison)" />
        public bool StartWithFileName(string value) => str.GetComparablePath().StartsWith(value.GetComparablePath(), StringComparison.OrdinalIgnoreCase);

        private string GetComparablePath() => Path.GetFullPath(str).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }
}

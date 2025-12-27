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
        if (!File.Exists(path))
        {
            if (!Directory.Exists(path))
                return path;
            // 尝试清理空目录
            _ = CleanEmptyDirectories(path, true);
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
        } while (Exists(candidate));
        return candidate;
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
    /// 在目录内最多 <paramref name="maxDepth"/> 次，将“父仅有唯一子项，且两者名称重复”的相邻两层合并。空文件夹直接删除
    /// </summary>
    /// <param name="root"></param>
    /// <param name="maxDepth">
    /// 最小有效值是1，表示仅处理<paramref name="root"/>与其唯一子项的合并。
    /// 每增加1，表示允许递归继续向下处理一层子目录。
    /// </param>
    /// <param name="strict"></param>
    /// <param name="useParentName">合并后原子项是否采用父文件夹的名字</param>
    /// <param name="useUniqueName">若子项移动后重名，是否重命名，否则放弃移动</param>
    public static void NormalizeRedundantNestedFolders(
        DirectoryInfo root,
        int maxDepth = 3,
        RedundantThreshold strict = RedundantThreshold.ContainsParent,
        bool useParentName = false,
        bool useUniqueName = false)
    {
        CleanEmptyDirectories(root, false);
        NormalizeRedundantNestedFoldersCore(root, maxDepth, strict, useParentName, useUniqueName);
    }

    private static void NormalizeRedundantNestedFoldersCore(DirectoryInfo current, int maxDepth, RedundantThreshold strict, bool useParentName, bool useUniqueName)
    {
        if (maxDepth <= 0 || !current.Exists)
            return;
        try
        {
            var entries = current.GetFileSystemInfos("*", SearchOption.TopDirectoryOnly);

            switch (entries)
            {
                // current -> a, b, c, ...
                case { Length: > 1 }:
                {
                    foreach (var directory in entries.OfType<DirectoryInfo>())
                        NormalizeRedundantNestedFoldersCore(directory, maxDepth - 1, strict, useParentName, useUniqueName);
                    break;
                }
                // current -> 空
                case []:
                {
                    current.Delete(false);
                    break;
                }
                // current -> onlyChild
                case [var onlyChild]:
                {
                    if (onlyChild is DirectoryInfo directoryInfo)
                        NormalizeRedundantNestedFoldersCore(directoryInfo, maxDepth - 1, strict, useParentName, useUniqueName);

                    // parent -> current -> onlyChild
                    if (current.Parent is { } parent)
                    {
                        if (onlyChild is not FileInfo { Extension: var ext, NameWithoutExtension: var childName })
                        {
                            childName = onlyChild.Name;
                            ext = "";
                        }

                        var merge = strict switch
                        {
                            RedundantThreshold.Equal => onlyChild.Name.EqualsFileName(current.Name),
                            RedundantThreshold.Always => true,
                            RedundantThreshold.ContainsParent => current.Name.ContainsFileName(childName),
                            RedundantThreshold.ContainsChild => childName.ContainsFileName(current.Name),
                            _ => false
                        };

                        if (merge)
                        {
                            var originalPath = onlyChild.FullName;

                            var suggestedDestPath = parent.Combine(useParentName ? current.Name : childName).GetComparablePath();
                            suggestedDestPath += ext;

                            if (onlyChild.MoveToDirectoryAndMerge(suggestedDestPath, useUniqueName))
                            {
                                Console.WriteLine($"合并重复文件夹：({originalPath}, {current.FullName}) -> {onlyChild.FullName}");

                                // 若成功移动，则此时onlyChild已和current在同一目录下
                                // 移出OnlyChild后，root变为空目录，删除它
                                // 判断名称是否相等，以防删除onlyChild
                                if (!onlyChild.Name.EqualsFileName(current.Name))
                                    current.Delete();
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

    /// <param name="info"></param>
    extension(FileInfo info)
    {
        public string NameWithoutExtension => Path.GetFileNameWithoutExtension(info.Name);
    }

    /// <param name="info"></param>
    extension(DirectoryInfo info)
    {
        public string Combine(FileInfo file) => Path.Combine(info.FullName, file.Name);

        public string Combine(string fileName) => Path.Combine(info.FullName, fileName);

        /// <summary>
        /// 删除<paramref name="info"/>，如果不存在则不操作
        /// </summary>
        /// <returns>是否成功删除</returns>
        public bool RemoveIfExists(bool recursive)
        {
            try
            {
                if (info.Exists)
                {
                    info.Delete(recursive);
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"删除失败 {info.FullName}：{ex.Message}");
            }

            return false;
        }
    }

    /// <param name="info"></param>
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

        /// <summary>
        /// 将移动<paramref name="info"/>为<paramref name="destPath"/>
        /// </summary>
        /// <param name="destPath">目标位置</param>
        public bool TryMoveTo(string destPath)
        {
            try
            {
                if (Path.GetDirectoryName(destPath) is { } directory)
                    _ = Directory.CreateDirectory(directory);
                info.MoveTo(destPath);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"移动失败（{info.FullName} -> {destPath}）：{ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 将<paramref name="info"/>移动到<paramref name="suggestedDestPath"/>，若存在则空目录则合并
        /// </summary>
        /// <remarks>
        /// <paramref name="info"/>可以是<paramref name="suggestedDestPath"/>的子项目（a/b/c.txt -> a = a/c.txt）
        /// </remarks>
        /// <param name="suggestedDestPath">目标位置</param>
        /// <param name="renameWhenDuplicated">当无法删除目标同名项目的时候，重命名<paramref name="suggestedDestPath"/></param>
        public bool MoveToDirectoryAndMerge(string suggestedDestPath, bool renameWhenDuplicated)
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
                        info.MoveTo(GetUniquePath(suggestedDestPath));
                        Directory.Delete(suggestedDestPath);
                    }
                    else if (entries.Length is not 0)
                        throw new InvalidOperationException("目标位置已存在非空目录");
                }

                info.MoveTo(suggestedDestPath);

                return true;
            }
            catch (Exception e)
            {
                if (renameWhenDuplicated)
                {
                    info.MoveTo(GetUniquePath(suggestedDestPath));
                    return true;
                }

                Console.WriteLine($"移动目录内容失败（{info.FullName} -> {suggestedDestPath}）：{e.Message}");

                return false;
            }
        }

        /// <summary>
        /// 删除<paramref name="info"/>，如果不存在则不操作
        /// </summary>
        /// <returns>是否成功删除</returns>
        public bool RemoveIfExists()
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

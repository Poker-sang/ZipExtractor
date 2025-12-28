using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Text;

namespace ZipExtractor;

public abstract class SevenZipExtractor : ExtractorBase
{
    private const string SevenZipPath = @"C:\App\7-Zip\7z.exe";

    private static (int ExitCode, IReadOnlyList<string> Outputs, IReadOnlyList<string> Errors)? SevenZipProcess(string arguments)
    {
        var processStartInfo = new ProcessStartInfo
        {
            FileName = SevenZipPath,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
            Arguments = arguments
        };

        using var process = Process.Start(processStartInfo);
        if (process is null)
        {
            Console.WriteLine($"{ColorRed}无法启动7-Zip进程{ColorReset}");
            return null;
        }

        var so = new List<string>();
        process.OutputDataReceived += (sender, args) =>
        {
            var line = args.Data;
            if (!string.IsNullOrWhiteSpace(line))
                so.Add(line);
        };
        var se = new List<string>();
        process.ErrorDataReceived += (sender, args) =>
        {
            var line = args.Data;
            if (!string.IsNullOrWhiteSpace(line))
                se.Add(line);
        };

        process.BeginOutputReadLine();
        process.BeginErrorReadLine();
        process.WaitForExit();
        return (process.ExitCode, so, se);
    }

    /// <summary>
    /// 递归解压压缩包，支持密码尝试
    /// </summary>
    /// <param name="passwords">可能的密码列表</param>
    public static void ExtractAll(IReadOnlyCollection<string?> passwords)
    {
        if (!File.Exists(SevenZipPath))
        {
            Console.WriteLine($"7-Zip未找到：{SevenZipPath}");
            return;
        }

        Console.WriteLine($"{ColorCyan}=== 开始解压任务 ==={ColorReset}");
        Console.WriteLine($"{ColorGray}密码列表数量：{passwords.Count}{ColorReset}");
        var rootDir = TempDir;
        Console.WriteLine($"{ColorGray}工作目录：{rootDir.FullName}{ColorReset}");
        var uncompletedVolumes = new Dictionary<string, VolumesInfo>();
        ExtractAll(rootDir, passwords, null, uncompletedVolumes);
        // 处理剩余不完整的分卷压缩包
        if (uncompletedVolumes.Count > 0)
        {
            Console.WriteLine($"\n{ColorYellow}处理 {uncompletedVolumes.Count} 个不完整的分卷压缩包...{ColorReset}");
            foreach (var uncompletedVolume in uncompletedVolumes.Values)
            {
                Console.WriteLine($"{ColorGray}  移动不完整分卷：{uncompletedVolume.CommonPrefix} ({uncompletedVolume.FoundVolumes.Count}/{uncompletedVolume.TotalVolumes} 卷){ColorReset}");
                MoveVolumeTo(uncompletedVolume, ErrorPath);
            }
        }
        Console.WriteLine($"\n{ColorCyan}=== 解压任务完成 ==={ColorReset}");
    }

    private static void ExtractAll(DirectoryInfo currDir, IReadOnlyCollection<string?> passwords,
        ArchiveInfo? originalArchive, Dictionary<string, VolumesInfo> uncompletedVolumes)
    {
        if (!currDir.Exists)
        {
            Console.WriteLine($"{ColorRed}根目录不存在：{currDir.FullName}{ColorReset}");
            return;
        }

        Console.WriteLine($"\n{ColorBlue}扫描目录：{ColorGray}{currDir.FullName}{ColorReset}");
        var files = currDir.GetFiles("*", SearchOption.AllDirectories);
        Console.WriteLine($"{ColorGray}  找到 {files.Length} 个文件{ColorReset}");
        var infos = new List<ArchiveInfo>();

        foreach (var archiveFile in files)
        {
            // 检查内部文件是否都是压缩包
            if (ArrangeFile(archiveFile, passwords) is { } info)
            {
                infos.Add(info);
                continue;
            }

            // 文本文件且小于1MB，跳过
            if (archiveFile.Extension.ToLower() is ".md" or ".txt" or "" && archiveFile.Length < 1 << 20)
                continue;

            // 有其他非压缩文件，认为解压完成，移动到完成目录
            if (originalArchive is { File: var originalFile })
            {
                Console.WriteLine($"{ColorGray}  发现非压缩文件，解压完成{ColorReset}");
                _ = currDir.MoveRelativePathTo(CompletePath);
                originalFile.Delete();
                Console.WriteLine($"{ColorGreen}✓ 已成功解压：{originalFile.Name} -> {currDir.Name}{ColorReset}");
                return;
            }
        }

        if (infos.Count > 0)
            Console.WriteLine($"{ColorGray}  识别到 {infos.Count} 个压缩文件{ColorReset}");

        foreach (var archiveInfo in infos)
        {
            // 单卷压缩包，直接解压
            if (!archiveInfo.IsMultiVolume)
            {
                Console.WriteLine($"\n{ColorBlue}解压单卷压缩包：{ColorReset}{archiveInfo.File.Name}");
                var outputDir = Extract(archiveInfo.File, archiveInfo.File.NameWithoutExtension, passwords);
                // 解压成功
                if (outputDir is not null)
                {
                    Console.WriteLine($"{ColorGray}  删除压缩包{ColorReset}");
                    // 删除压缩包
                    archiveInfo.File.Delete();
                    // 继续尝试解压解压后的文件
                    ExtractAll(outputDir, passwords, archiveInfo, uncompletedVolumes);
                }
                else
                {
                    Console.WriteLine($"{ColorRed}  解压失败，移动到错误目录{ColorReset}");
                    // 解压失败，移动到错误目录
                    _ = archiveInfo.File.MoveRelativePathTo(ErrorPath);
                }

                continue;
            }

            // 多卷压缩包，收集分卷信息
            if (uncompletedVolumes.TryGetValue(archiveInfo.CommonPrefix, out var volumesInfo))
            {
                Console.WriteLine($"\n{ColorBlue}收集分卷：{ColorReset}{archiveInfo.File.Name} {ColorGray}[{archiveInfo.VolumeIndex + 1}/{archiveInfo.Volumes}]{ColorReset}");
                if (volumesInfo.AddNewVolume(archiveInfo))
                {
                    Console.WriteLine($"{ColorGreen}  分卷已齐全 ({volumesInfo.TotalVolumes} 卷)，开始解压：{volumesInfo.CommonPrefix}{ColorReset}");
                    // 分卷齐全，开始解压
                    var outputDir = Extract(volumesInfo.FirstVolume.File, volumesInfo.CommonPrefix, passwords);
                    // 解压成功
                    if (outputDir is not null)
                    {
                        Console.WriteLine($"{ColorGray}  删除所有分卷{ColorReset}");
                        // 删除所有分卷压缩包
                        foreach (var foundVolume in volumesInfo.FoundVolumes)
                            foundVolume.File.Delete();
                        // 解压成功，继续尝试解压解压后的文件
                        ExtractAll(outputDir, passwords, archiveInfo, uncompletedVolumes);
                    }
                    else
                    {
                        Console.WriteLine($"{ColorRed}  解压失败，移动到错误目录{ColorReset}");
                        MoveVolumeTo(volumesInfo, ErrorPath);
                    }

                    // 移除已处理的分卷信息
                    _ = uncompletedVolumes.Remove(archiveInfo.CommonPrefix);
                }
                else
                {
                    Console.WriteLine($"{ColorYellow}  等待其他分卷 ({volumesInfo.FoundVolumes.Count}/{volumesInfo.TotalVolumes}){ColorReset}");
                }
            }
            else
            {
                Console.WriteLine($"\n{ColorBlue}发现新的分卷压缩包：{ColorReset}{archiveInfo.CommonPrefix} {ColorGray}[{archiveInfo.VolumeIndex + 1}/{archiveInfo.Volumes}]{ColorReset}");
                volumesInfo = new()
                {
                    FirstVolume = archiveInfo,
                    FoundVolumes = new(archiveInfo.Volumes) { archiveInfo },
                    TotalVolumes = archiveInfo.Volumes,
                    CommonPrefix = archiveInfo.CommonPrefix
                };
                uncompletedVolumes[archiveInfo.CommonPrefix] = volumesInfo;
                Console.WriteLine($"{ColorYellow}  等待其他分卷 (1/{archiveInfo.Volumes}){ColorReset}");
            }
        }
    }

    private static ArchiveInfo? ArrangeFile(FileInfo archiveFile, IReadOnlyCollection<string?> passwords)
    {
        if (archiveFile is not { Exists: true, Directory: not null })
        {
            Console.WriteLine($"{ColorRed}文件异常：{archiveFile.FullName}{ColorReset}");
            return null;
        }

        Console.WriteLine($"{ColorGray}  检查文件：{archiveFile.Name}{ColorReset}");
        if (CheckFileFormat(archiveFile, passwords) is not { } info)
        {
            Console.WriteLine($"{ColorGray}    不是压缩文件{ColorReset}");
            _ = archiveFile.MoveRelativePathTo(ErrorPath);
            return null;
        }

        if (info.Type is ArchiveInfo.Format.Other)
        {
            Console.WriteLine($"{ColorYellow}    不支持的压缩格式{ColorReset}");
            _ = archiveFile.MoveRelativePathTo(ErrorPath);
            return null;
        }

        var volumeInfo = info.IsMultiVolume ? $" {ColorGray}(分卷 {info.VolumeIndex + 1}/{info.Volumes}){ColorReset}" : "";
        Console.WriteLine($"{ColorGreen}    识别为 {info.Type} 格式{volumeInfo}{ColorReset}");
        return info;
    }

    private struct VolumesInfo
    {
        public required ArchiveInfo FirstVolume;

        public required List<ArchiveInfo> FoundVolumes;

        public required int TotalVolumes;

        public required string CommonPrefix;

        public bool ReadyForExtract => TotalVolumes == FoundVolumes.Count;

        public bool AddNewVolume(ArchiveInfo newArchive)
        {
            if (newArchive.IsFirstVolume)
            {
                FoundVolumes.Insert(0, newArchive);
                FirstVolume = newArchive;
            }

            return ReadyForExtract;
        }
    }

    private enum ExtractStatus
    {
        Success,
        WrongPassword,
        Failed
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="archivePath"></param>
    /// <param name="outputFolderName"></param>
    /// <param name="passwords"></param>
    /// <returns>解压出的文件</returns>
    private static DirectoryInfo? Extract(FileInfo archivePath, string outputFolderName, IReadOnlyCollection<string?> passwords)
    {
        var outPath = FileSystemHelper.GetUniquePath(Path.Combine(archivePath.DirectoryName!, outputFolderName));
        var outputDir = new DirectoryInfo(outPath);
        Console.WriteLine($"{ColorGray}  目标目录：{outputDir.Name}{ColorReset}");

        var passwordIndex = 0;
        foreach (var password in passwords)
        {
            passwordIndex++;
            var passwordHint = string.IsNullOrEmpty(password) ? "无密码" : $"密码 {passwordIndex}/{passwords.Count}";
            Console.WriteLine($"{ColorGray}  尝试解压 ({passwordHint})...{ColorReset}");

            var result = ExtractCore(archivePath.FullName, outputDir.FullName, password);
            switch (result)
            {
                case ExtractStatus.Success:
                    Console.WriteLine($"{ColorGreen}  ✓ 解压成功{ColorReset}");
                    return outputDir;
                case ExtractStatus.WrongPassword:
                    Console.WriteLine($"{ColorYellow}  × 密码错误{ColorReset}");
                    // 解压失败的空文件夹
                    _ = outputDir.RemoveIfExists(true);
                    continue;
                case ExtractStatus.Failed:
                default:
                    Console.WriteLine($"{ColorRed}  × 解压失败{ColorReset}");
                    _ = outputDir.RemoveIfExists(true);
                    return null;
            }
        }

        Console.WriteLine($"{ColorRed}  × 解压失败：没有正确的密码{ColorReset}");
        return null;
    }

    /// <summary>
    /// 执行解压操作
    /// </summary>
    private static ExtractStatus ExtractCore(string archivePath, string outputPath, string? password)
    {
        try
        {
            // 构建命令行参数
            // x: 完全提取（保留目录结构）
            // -o: 输出目录
            // -aou: 不覆盖而是重命名解压结果
            // -p: 密码（用 -p- 表示无密码）
            // -y: 对所有询问回答 yes
            var args = new StringBuilder("x -y -aou -p")
                .Append(string.IsNullOrEmpty(password) ? "-" : $"\"{password}\"")
                .Append($" \"{archivePath}\"")
                .Append($" -o\"{outputPath}\"")
                .ToString();

            if (SevenZipProcess(args) is not (ExitCode: var exitCode, _, Errors: var errors))
                return ExtractStatus.Failed;

            if (exitCode is 0)
                return ExtractStatus.Success;

            var sb = new StringBuilder();

            foreach (var error in errors)
            {
                sb.AppendLine(error);
                if (error.Contains("Wrong password", StringComparison.OrdinalIgnoreCase))
                    return ExtractStatus.WrongPassword;
            }

            Console.WriteLine($"{ColorRed}解压失败：{sb}{ColorReset}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"{ColorRed}解压出错：{ex.Message}{ColorReset}");
        }

        return ExtractStatus.Failed;
    }

    private static void MoveVolumeTo(VolumesInfo volumesInfo, string toPath, string fromPath = TempPath)
    {
        // 解压失败，移动到同一个错误目录下
        // 由于之前不一定在同一目录，移动到一起可能重名
        var relativePath = Path.GetRelativePath(fromPath, volumesInfo.FirstVolume.File.FullName);
        var dest = Path.Combine(toPath, relativePath);
        var destPath = FileSystemHelper.GetUniquePath(dest);
        foreach (var foundVolume in volumesInfo.FoundVolumes)
            _ = foundVolume.File.TryMoveTo(FileSystemHelper.GetUniquePath(Path.Combine(destPath, foundVolume.File.Name)));
    }

    private static ArchiveInfo? CheckFileFormat(FileInfo file, IReadOnlyCollection<string?> passwords)
    {
        string? oldName = null;

        try
        {
            var result = CheckFileFormatCore(file, null, out var info);
            if (result is StatusCode.NeedRename)
            {
                Console.WriteLine($"{ColorGray}    重命名文件后重试...{ColorReset}");
                oldName = file.FullName;
                file.MoveTo(Path.ChangeExtension(oldName, null));
                result = CheckFileFormatCore(file, null, out info);
            }

            foreach (var password in passwords)
            {
                if (result is not StatusCode.WrongPassword)
                    return info;
                if (password is not null)
                {
                    Console.WriteLine($"{ColorGray}    尝试使用密码验证...{ColorReset}");
                    result = CheckFileFormatCore(file, password, out info);
                }
            }

            return info;
        }
        finally
        {
            if (oldName is not null)
                file.MoveTo(oldName);
        }
    }

    private static StatusCode CheckFileFormatCore(FileInfo file, string? password, out ArchiveInfo? compressionInfo)
    {
        try
        {
            // l 列出压缩包内容
            // -y 假设"是"对所有提示
            // -p 设置密码
            var args = new StringBuilder("l -y -p")
                .Append(password is null ? "-" : $"\"{password}\"")
                .Append($" \"{file.FullName}\"")
                .ToString();

            compressionInfo = null;

            if (SevenZipProcess(args) is not (_, Outputs: var outputs, Errors: var errors))
                return StatusCode.Failed;

            var info = new ArchiveInfo { File = file };
            var isArchive = false;
            foreach (var output in outputs)
            {
                if (output.Split(" = ") is not [var key, var value])
                    continue;
                switch (key)
                {
                    case "Type":
                        isArchive = true;
                        if (output.Contains("rar", StringComparison.OrdinalIgnoreCase))
                            info.Type = ArchiveInfo.Format.Rar;
                        else if (output.Contains("7z", StringComparison.OrdinalIgnoreCase))
                            info.Type = ArchiveInfo.Format._7Z;
                        else if (output.Contains("zip", StringComparison.OrdinalIgnoreCase))
                            info.Type = ArchiveInfo.Format.Zip;
                        else
                            info.Type = ArchiveInfo.Format.Other;
                        break;
                    case "Volume Index":
                        info.VolumeIndex = int.Parse(value);
                        break;
                    case "Volumes":
                        info.Volumes = int.Parse(value);
                        break;
                }
            }

            if (info.IsMultiVolume)
            {
                info.CommonPrefix = info.Type switch
                {
                    // .partN.rar
                    ArchiveInfo.Format.Rar => Path.GetFileNameWithoutExtension(file.NameWithoutExtension),
                    // .zip / .zNN
                    // .7z.NNN
                    ArchiveInfo.Format.Zip or ArchiveInfo.Format._7Z => file.NameWithoutExtension,
                    _ => null
                };
            }

            if (isArchive)
            {
                compressionInfo = info;
                return info.Type is not ArchiveInfo.Format.Other ? StatusCode.Success : StatusCode.Other;
            }

            foreach (var error in errors)
            {
                Console.WriteLine($"{ColorRed}{error}{ColorReset}");
                if (error.Contains("Wrong password", StringComparison.OrdinalIgnoreCase))
                    return StatusCode.WrongPassword;
                if (error.Contains("Cannot open", StringComparison.OrdinalIgnoreCase))
                    return error.Contains("Cannot open the file as [", StringComparison.OrdinalIgnoreCase)
                        ? StatusCode.NeedRename
                        : StatusCode.Other;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"{ColorRed}{ex.Message}{ColorReset}");
        }

        compressionInfo = null;

        return StatusCode.Failed;
    }

    private struct ArchiveInfo
    {
        public required FileInfo File;

        public Format Type;

        public int VolumeIndex;

        public int Volumes;

        [MemberNotNullWhen(true, nameof(CommonPrefix))]
        public bool IsMultiVolume => Type is not Format.Other && Volumes > 1;

        public bool IsFirstVolume => VolumeIndex is 0;

        public string? CommonPrefix;

        public enum Format
        {
            /// <summary>
            /// 其他压缩文件（暂不处理）
            /// </summary>
            Other,
            Zip,
            _7Z,
            Rar
        }
    }

    private enum StatusCode
    {
        /// <summary>
        /// 是压缩文件
        /// </summary>
        Success,

        /// <summary>
        /// 7-Zip只根据后缀名尝试了一种格式，删除后缀名后重试，一般在第一次尝试无密码时就会出现
        /// </summary>
        NeedRename,

        /// <summary>
        /// 非压缩文件
        /// </summary>
        Other,

        /// <summary>
        /// 错误密码（加密文件名7z/rar）
        /// </summary>
        WrongPassword,

        /// <summary>
        /// 错误
        /// </summary>
        Failed
    }
}

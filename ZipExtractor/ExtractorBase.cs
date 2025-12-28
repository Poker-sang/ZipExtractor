using System.IO;

namespace ZipExtractor;

public abstract class ExtractorBase
{
    private const string BasePath = @"D:\";

    /// <summary>
    /// 解压文件的目录
    /// </summary>
    public const string TempPath = BasePath + "TempExt";

    /// <summary>
    /// 解压完成的目录
    /// </summary>
    public static readonly string CompletePath = BasePath + "CompleteExt";

    /// <summary>
    /// 解压失败的目录
    /// </summary>
    public static readonly string ErrorPath = BasePath + "ErrorExt";

    /// <inheritdoc cref="TempPath" />
    public static DirectoryInfo TempDir { get; } = Directory.CreateDirectory(TempPath);

    /// <inheritdoc cref="CompletePath" />
    public static DirectoryInfo CompleteDir { get; } = Directory.CreateDirectory(CompletePath);

    /// <inheritdoc cref="ErrorPath" />
    public static DirectoryInfo ErrorDir { get; } = Directory.CreateDirectory(ErrorPath);

    // ANSI 颜色代码
    public const string ColorReset = "\e[0m";
    public const string ColorGreen = "\e[32m";
    public const string ColorRed = "\e[31m";
    public const string ColorYellow = "\e[33m";
    public const string ColorBlue = "\e[34m";
    public const string ColorGray = "\e[90m";
    public const string ColorCyan = "\e[36m";
}

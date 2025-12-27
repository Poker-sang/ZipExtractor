using System.IO;

namespace ZipExtractor;

public abstract class ExtractorBase
{
    private const string BasePath = @"D:\";
    public const string TempPath = BasePath + "TempExt";
    public static readonly string CompletePath = BasePath + "CompleteExt";
    public static readonly string ErrorPath = BasePath + "ErrorExt";
    public static readonly DirectoryInfo TempDir;
    public static readonly DirectoryInfo CompleteDir;
    public static readonly DirectoryInfo ErrorDir;

    static ExtractorBase()
    {
        TempDir = Directory.CreateDirectory(TempPath);
        CompleteDir = Directory.CreateDirectory(CompletePath);
        ErrorDir = Directory.CreateDirectory(ErrorPath);
    }
}

using System;
using System.IO;
using System.Threading;
using ZipExtractor;
using ZipExtractor.Extraction;
using ZipExtractor.Logging;

// 密码候选（先试空密码）
string?[] passwords = [
    null,
    "FLYYZ",
    "yejiang",
    "yecgaa",
    "Drgon Slayer",
    "病名为祈",
    "2A46M-5",
    "shiki",
    "Geass",
    "hmoe.top",
    "izaya",
    "fengliyds",
    "背影",
    "hihihiha",
    "LSFS",
    "dx",
    "GS_mel",
    "xxld",
    "⑨",
    "nameless",
    "south-plus",
    "tuyile2026.!2333",
    "图一乐讨厌倒狗",
    "XueFc",
    "sixpluswan",
    "米粒儿"
];
if (args.Contains("--help"))
{
    Console.WriteLine("""
        ZipExtractor · 7-Zip + UnRAR

          --input <目录>         待解压目录
          --output <目录>        完成目录
          --password <密码>      增加候选密码（日志中显示）
          --keep-sources         成功后仍保留源包
          --max-depth <层数>     嵌套解压上限，默认 20
          --non-interactive      不等待按键
          --no-normalize         不整理输出目录层级
          --update-directory-times  整理操作更新目录修改日期
          --7zip-dll <路径>      指定 7z.dll
          --unrar-dll <路径>      指定 UnRAR.dll

        Ctrl+C 取消 · 自动检测终端颜色和进度条 · NO_COLOR 禁用颜色
        """);
    return;
}

using var cancellation = new CancellationTokenSource();
Console.CancelKeyPress += (_, e) =>
{
    e.Cancel = true;
    cancellation.Cancel();
};
try
{
    var input = Option("--input") ?? @"D:\TempExt";
    var complete = Option("--output") ?? @"D:\CompleteExt";
    var sevenZipDll = Option("--7zip-dll") ?? Path.Combine(AppContext.BaseDirectory, "native", "7z.dll");
    var unrarDll = Option("--unrar-dll") ?? Path.Combine(AppContext.BaseDirectory, "native", "UnRAR.dll");
    if (!File.Exists(sevenZipDll) || !File.Exists(unrarDll))
        throw new FileNotFoundException("缺少原生解压库。请先运行 scripts/Install-NativeLibraries.ps1，或指定 --7zip-dll 和 --unrar-dll。");
    if (Option("--password") is { } supplied)
        passwords = [
            ..passwords,
            supplied
        ];
    _ = Directory.CreateDirectory(input);
    using var sevenZip = new SevenZipNativeExtractor(sevenZipDll);
    using var unrar = new UnRarNativeExtractor(unrarDll);
    var updateDirectoryLastWriteTime = true;
    var options = new PipelineOptions(input, complete, args.Contains("--keep-sources"), int.Parse(Option("--max-depth") ?? "20"), updateDirectoryLastWriteTime);
    ConsoleLog.Default.Header(options);
    var failures = new ExtractionPipeline([
        sevenZip,
        unrar
    ]).Run(options, passwords, cancellationToken: cancellation.Token);
    ConsoleLog.Default.Info("整理输出目录…");
    _ = FileSystemHelper.CleanEmptyDirectories(input, false, options.UpdateDirectoryLastWriteTime);
    var completeDir = new DirectoryInfo(complete);
    if (!args.Contains("--no-normalize"))
    {
        FileSystemHelper.NormalizeNestedFoldersContainingName(completeDir, "新建文件夹", 20, updateDirectoryLastWriteTime: options.UpdateDirectoryLastWriteTime);
        FileSystemHelper.NormalizeRedundantNestedFolders(completeDir, 20, updateDirectoryLastWriteTime: options.UpdateDirectoryLastWriteTime);
        if (!args.Contains("--non-interactive") && !Console.IsInputRedirected)
        {
            ConsoleLog.Default.Info("按任意键继续合并唯一子目录…");
            _ = Console.ReadKey(true);
            FileSystemHelper.NormalizeRedundantNestedFolders(completeDir, 20, FileSystemHelper.RedundantThreshold.Always, updateDirectoryLastWriteTime: options.UpdateDirectoryLastWriteTime);
        }
    }

    Environment.ExitCode = failures is 0 ? 0 : 1;
}
catch (OperationCanceledException)
{
    ConsoleLog.Default.Warning("已取消，当前任务源文件已保留。");
    Environment.ExitCode = 2;
}
catch (Exception e)
{
    ConsoleLog.Default.Error(e.Message);
    Environment.ExitCode = 1;
}

return;
string? Option(string key)
{
    var index = Array.IndexOf(args, key);
    return index < 0 ? null : index + 1 >= args.Length ? throw new ArgumentException($"缺少 {key} 参数值") : args[index + 1];
}

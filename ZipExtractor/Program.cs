using System;
using System.IO;
using ZipExtractor;

// 密码候选（先试空密码）
string?[] passwords =
[
    null,
    "FLYYZ",
    "yejiang",
    "yecgaa",
    "Drgon Slayer",
    "病名为祈",
    "shiki",
    "Geass",
    "hmoe.top",
    "izaya",
    "fengliyds",
    "背影",
    "hihihiha",
    "LSFS",
    "dx",
    "梅酱",
    "GS_mel",
    "xxld",
    "6666",
    "⑨",
    "nameless",
    "south-plus",
    "tuyile2025.!2333",
    "图一乐讨厌倒狗",
    "mm666",
    "XueFc"
];

FileSystemHelper.NormalizeRedundantNestedFolders(ExtractorBase.CompleteDir);
FileSystemHelper.NormalizeRedundantNestedFolders(ExtractorBase.ErrorDir);
FileSystemHelper.NormalizeRedundantNestedFolders(ExtractorBase.TempDir);
SevenZipExtractor.ExtractAll(passwords);


//while (true)
//{
//    Console.WriteLine("请输入待解压文件的完整路径（回车退出）：");
//    var inputPath = Console.ReadLine()?.Trim().Trim('"');
//    inputPath = "D:\\FuckNew";
//    if (string.IsNullOrWhiteSpace(inputPath) || inputPath.ContainsAny(Path.GetInvalidPathChars()))
//    {
//        Console.WriteLine("已退出。");
//        break;
//    }

//    if (Directory.Exists(inputPath))
//    {
//        // 先清理输入路径内的空文件夹

//        foreach (var file in Directory.EnumerateFiles(inputPath, "*", SearchOption.AllDirectories))
//        {
//            _ = FileSystemHelper.CleanEmptyDirectories(inputPath, false);
//            WinRarExtractor.ExtractRecursively(file, inputPath, passwords);
//        }
//    }
//    else
//    {
//        var rootForSingle = Path.GetDirectoryName(inputPath);
//        if (string.IsNullOrEmpty(rootForSingle))
//            rootForSingle = Directory.GetCurrentDirectory();
//        WinRarExtractor.ExtractRecursively(inputPath, rootForSingle, passwords);
//    }
//}

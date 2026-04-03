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
    "sixpluswan"
];

FileSystemHelper.CleanEmptyDirectories(ExtractorBase.TempDir, false);

// SevenZipExtractor.ExtractAll(passwords);

foreach (var file in ExtractorBase.TempDir.EnumerateFiles("*", SearchOption.AllDirectories))
    WinRarExtractor.ExtractRecursively(file, passwords);

Console.WriteLine("等待清理...");

Console.ReadKey();

FileSystemHelper.CleanEmptyDirectories(ExtractorBase.TempDir, false);

FileSystemHelper.NormalizeRedundantNestedFolders(ExtractorBase.CompleteDir, 20,
    FileSystemHelper.RedundantThreshold.Always, true);

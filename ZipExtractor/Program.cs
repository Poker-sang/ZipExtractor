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

// SevenZipExtractor.ExtractAll(passwords);

foreach (var file in ExtractorBase.TempDir.EnumerateFiles("*", SearchOption.AllDirectories))
    WinRarExtractor.ExtractRecursively(file, passwords);

Console.WriteLine("等待清理...");

Console.ReadKey();

FileSystemHelper.NormalizeRedundantNestedFolders(ExtractorBase.CompleteDir, 5);

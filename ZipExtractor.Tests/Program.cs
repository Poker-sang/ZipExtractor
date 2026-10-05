using System.Diagnostics;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using ZipExtractor;
using ZipExtractor.Extraction;

// Integration harness: generates its own archives; never scans the user's input directory.
var repo = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../"));
var root = Path.Combine(repo, "work", "tests-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
var zipExe = Environment.GetEnvironmentVariable("TEST_7ZIP_EXE") ?? @"C:\App\7-Zip\7z.exe";
var rarExe = Environment.GetEnvironmentVariable("TEST_RAR_EXE") ?? @"C:\App\WinRAR\Rar.exe";
using var zip = new SevenZipNativeExtractor(Path.Combine(repo, "ZipExtractor/native/7z.dll"));
using var rar = new UnRarNativeExtractor(Path.Combine(repo, "ZipExtractor/native/UnRAR.dll"));
var discovery = new ArchiveDiscovery();
var tests = 0;
var failures = 0;
var payload = new byte[240000];
new Random(42).NextBytes(payload);
const string password = "中文 空格\\尾部密码";
string Password(string format) => format is "zip" ? "ASCII space\\password" : password;
void Assert(bool condition, string message)
{
    if (!condition)
        throw new Exception(message);
}

void Test(string name, Action action)
{
    tests++;
    try
    {
        action();
        Console.WriteLine($"PASS {name}");
    }
    catch (Exception e)
    {
        failures++;
        Console.WriteLine($"FAIL {name}: {e}");
    }
}

string Case(string name)
{
    var path = Path.Combine(root, name);
    _ = Directory.CreateDirectory(path);
    return path;
}

void Command(string exe, string cwd, params string[] arguments)
{
    var start = new ProcessStartInfo(exe)
    {
        WorkingDirectory = cwd,
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        UseShellExecute = false
    };
    foreach (var arg in arguments)
        start.ArgumentList.Add(arg);
    using var process = Process.Start(start)!;
    var output = process.StandardOutput.ReadToEndAsync();
    var error = process.StandardError.ReadToEndAsync();
    if (!process.WaitForExit(60000))
    {
        process.Kill(true);
        throw new Exception("Fixture generator timeout");
    }

    if (process.ExitCode is not 0)
        throw new Exception($"Fixture generator: {process.ExitCode} {output.Result} {error.Result}");
}

string Generate(string name, string format, bool encrypted = false, bool split = false)
{
    var dir = Case(name);
    File.WriteAllBytes(Path.Combine(dir, "内容.bin"), payload);
    var archive = "archive." + format;
    if (format is "rar")
    {
        var a = new List<string>
        {
            "a",
            "-ep",
            "-s",
            "-idq"
        };
        if (encrypted)
            a.Add("-hp" + Password(format));
        if (split)
            a.Add("-v64k");
        a.AddRange([
            archive,
            "内容.bin"
        ]);
        Command(rarExe, dir, [.. a]);
    }
    else
    {
        var a = new List<string>
        {
            "a",
            "-t" + format
        };
        if (encrypted)
        {
            a.Add("-p" + Password(format));
            a.Add(format is "7z" ? "-mhe=on" : "-mem=AES256");
        }

        if (split)
            a.Add("-v64k");
        a.AddRange([
            archive,
            "内容.bin"
        ]);
        Command(zipExe, dir, [.. a]);
    }

    File.Delete(Path.Combine(dir, "内容.bin"));
    return dir;
}

ExtractionResult Extract(ArchiveSet set, IArchiveExtractor engine, string? pass, out string output, CancellationToken token = default)
{
    output = Case("out-" + Guid.NewGuid().ToString("N"));
    return engine.Extract(new(set, output, pass, CancellationToken: token));
}

void Verify(ArchiveSet set, IArchiveExtractor engine, string? pass = null)
{
    var result = Extract(set, engine, pass, out var output);
    Assert(result.Succeeded, $"{result.Status}: {result.Message}");
    Assert(result.VerifiedParts?.Count == set.Parts.Count, "Consumed-volume tracking mismatch");
    Assert(File.ReadAllBytes(Path.Combine(output, "内容.bin")).SequenceEqual(payload), "Content mismatch");
}

foreach (var format in new[]
{
    "zip",
    "7z",
    "rar"
}

)
{
    IArchiveExtractor engine = format is "rar" ? rar : zip;
    Test($"{format}: wrong extension + Unicode filename", () =>
    {
        var dir = Generate(format + "-plain", format);
        File.Move(Directory.GetFiles(dir).Single(), Path.Combine(dir, "伪装.txt"));
        Verify(ArchiveDiscovery.Scan(dir).Single(), engine);
    });
    Test($"{format}: password with spaces, wrong password then correct", () =>
    {
        var set = ArchiveDiscovery.Scan(Generate(format + "-encrypted", format, true)).Single();
        var wrong = Extract(set, engine, "wrong", out _);
        Assert(wrong.Status is ExtractionStatus.WrongPassword or ExtractionStatus.PasswordOrCorruptData, wrong.ToString());
        Verify(set, engine, Password(format));
    });
    Test($"{format}: encrypted split archive with appended fake suffix", () =>
    {
        var dir = Generate(format + "-split", format, true, true);
        foreach (var part in Directory.GetFiles(dir))
            File.Move(part, part + ".fake");
        var set = ArchiveDiscovery.Scan(dir).Single();
        Assert(set.Parts.Count > 1, "Not grouped");
        Verify(set, engine, Password(format));
    });
    Test($"{format}: missing middle volume", () =>
    {
        var dir = Generate(format + "-missing", format, true, true);
        var parts = Directory.GetFiles(dir).Order().ToArray();
        File.Delete(parts[1]);
        var set = ArchiveDiscovery.Scan(dir).Single();
        var result = Extract(set, engine, Password(format), out _);
        Assert(result.Status == ExtractionStatus.MissingVolume, result.ToString());
    });
}

foreach (var trailerSize in new[] { 0, 1500000 })
    Test($"Embedded ZIP: large MP4 prefix and trailer={trailerSize}", () =>
    {
        var dir = Generate("embedded-" + trailerSize, "zip");
        var original = Directory.GetFiles(dir).Single();
        var bytes = File.ReadAllBytes(original);
        File.Delete(original);
        var carrier = Path.Combine(dir, "视频.mp4");
        const int prefixLength = 16 * 1024 * 1024;
        using (var output = File.Create(carrier))
        {
            output.Write([0, 0, 0, 32, 102, 116, 121, 112, 105, 115, 111, 109]);
            output.Position = prefixLength;
            output.Write(bytes);
            output.Write(new byte[trailerSize]);
        }

        var set = ArchiveDiscovery.Scan(dir, probe: zip).Single();
        Assert(set.Format is ArchiveFormat.Zip, "Native ZIP probe failed");
        Verify(set, zip);
        Assert(new FileInfo(carrier).Length == prefixLength + bytes.Length + trailerSize, "Carrier was modified");
    });
foreach (var encrypted in new[] { false, true })
    Test($"Native probe: embedded 7z, encrypted={encrypted}", () =>
    {
        var dir = Generate("native-embedded-7z-" + encrypted, "7z", encrypted);
        var original = Directory.GetFiles(dir).Single();
        var bytes = File.ReadAllBytes(original);
        File.Delete(original);
        var carrier = Path.Combine(dir, "伪装.mp4");
        using (var output = File.Create(carrier))
        {
            output.Position = 16 * 1024 * 1024;
            output.Write(bytes);
            output.Write(new byte[14354]);
        }

        Verify(ArchiveDiscovery.Scan(dir, probe: zip).Single(), zip, encrypted ? Password("7z") : null);
    });
Test("Embedded ZIP: extraction still rejects corrupt payload", () =>
{
    var dir = Generate("embedded-crc", "zip");
    var original = Directory.GetFiles(dir).Single();
    var bytes = File.ReadAllBytes(original);
    File.Delete(original);
    bytes[bytes.Length / 2] ^= 0xff;
    File.WriteAllBytes(Path.Combine(dir, "伪装.mp4"), [.. new byte[100], .. bytes, .. new byte[100]]);
    var result = Extract(ArchiveDiscovery.Scan(dir, probe: zip).Single(), zip, null, out _);
    Assert(result.Status == ExtractionStatus.InvalidArchive, "Embedded ZIP bypassed payload validation");
});
Test("Embedded ZIP: recursive pipeline recognizes a disguised inner archive", () =>
{
    var dir = Case("embedded-recursive");
    var inner = Generate("embedded-recursive-inner", "zip");
    var content = Path.Combine(dir, "content");
    _ = Directory.CreateDirectory(content);
    File.WriteAllBytes(Path.Combine(content, "视频.mp4"), [.. new byte[300], .. File.ReadAllBytes(Directory.GetFiles(inner).Single()), .. new byte[100]]);
    var input = Path.Combine(dir, "input");
    _ = Directory.CreateDirectory(input);
    var source = Path.Combine(input, "outer.zip");
    ZipFile.CreateFromDirectory(content, source);
    var options = new PipelineOptions(input, Path.Combine(dir, "complete"), true);
    Assert(new ExtractionPipeline([zip]).Run(options, []) is 0, "Nested embedded ZIP failed");
    Assert(File.ReadAllBytes(Path.Combine(options.CompleteDirectory, "outer/视频/内容.bin")).SequenceEqual(payload), "Nested content missing");
    Assert(File.Exists(source), "Source was lost");
});
Test("Embedded ZIP: false end-record and corrupt directory rejected", () =>
{
    var dir = Generate("embedded-invalid", "zip");
    var original = Directory.GetFiles(dir).Single();
    var bytes = File.ReadAllBytes(original);
    File.Delete(original);
    var central = bytes.AsSpan().IndexOf(new byte[] { 0x50, 0x4b, 1, 2 });
    Assert(central >= 0, "Missing test directory");
    bytes[central] = 0;
    var carrier = Path.Combine(dir, "伪装.mp4");
    File.WriteAllBytes(carrier, [.. new byte[300], .. bytes, .. new byte[100]]);
    Assert(ArchiveDiscovery.Scan(dir, probe: zip).Count is 0, "Unvalidated signature accepted as ZIP");
});
Test("Embedded ZIP: prefixed ZIP64 directory", () =>
{
    var dir = Generate("embedded-zip64", "zip");
    var original = Directory.GetFiles(dir).Single();
    var bytes = File.ReadAllBytes(original);
    File.Delete(original);
    var end = bytes.AsSpan().LastIndexOf(new byte[] { 0x50, 0x4b, 5, 6 });
    var count = System.Buffers.Binary.BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(end + 10));
    var size = System.Buffers.Binary.BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(end + 12));
    var offset = System.Buffers.Binary.BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(end + 16));
    var carrier = Path.Combine(dir, "伪装.mp4");
    using (var output = File.Create(carrier))
    using (var writer = new BinaryWriter(output))
    {
        writer.Write(new byte[256]);
        writer.Write(bytes.AsSpan(0, end));
        writer.Write(0x06064b50u);
        writer.Write(44UL);
        writer.Write((ushort) 45);
        writer.Write((ushort) 45);
        writer.Write(0u);
        writer.Write(0u);
        writer.Write((ulong) count);
        writer.Write((ulong) count);
        writer.Write((ulong) size);
        writer.Write((ulong) offset);
        writer.Write(0x07064b50u);
        writer.Write(0u);
        writer.Write((ulong) end);
        writer.Write(1u);
        writer.Write(0x06054b50u);
        writer.Write((ushort) 0);
        writer.Write((ushort) 0);
        writer.Write(ushort.MaxValue);
        writer.Write(ushort.MaxValue);
        writer.Write(uint.MaxValue);
        writer.Write(uint.MaxValue);
        writer.Write((ushort) 0);
        writer.Write(new byte[99]);
    }

    Verify(ArchiveDiscovery.Scan(dir, probe: zip).Single(), zip);
});

Test("RAR5: header volume numbers with completely replaced suffixes", () =>
{
    var dir = Generate("rar-headers", "rar", false, true);
    var parts = Directory.GetFiles(dir).Order().ToArray();
    for (var i = 0; i < parts.Length; i++)
        File.Move(parts[i], Path.Combine(dir, "series." + (char) ('a' + i)));
    var set = ArchiveDiscovery.Scan(dir).Single();
    Assert(set.Parts.Count == parts.Length, "Header grouping failed");
    Verify(set, rar);
});
Test("ZIP: native split disks, wrong final extension", () =>
{
    var dir = Case("native-zip");
    MakeSplitZip(dir, payload);
    var set = ArchiveDiscovery.Scan(dir).Single();
    Assert(set.Layout == VolumeLayout.Native, "Not recognized as native ZIP volumes");
    Verify(set, zip);
});
foreach (var legacyNames in new[]
{
    false,
    true
}

)
    Test($"RAR4: {(legacyNames ? ".rar/.r00" : "partNN.rar.fake")} volumes", () =>
    {
        var dir = Case("rar4-" + legacyNames);
        for (var i = 0; i < 3; i++)
        {
            var fixture = Path.Combine(repo, "ZipExtractor.Tests/Fixtures", $"test_rar_multivolume_single_file.part{i + 1}.rar.uu");
            var bytes = DecodeUu(File.ReadAllLines(fixture));
            if (legacyNames)
            {
                bytes[10] &= 0xef; // Clear RAR4 MHD_NEWNUMBERING, leaving the archive payload intact.
                var headerSize = BitConverter.ToUInt16(bytes, 12);
                var crc = Crc(bytes.AsSpan(9, headerSize - 2));
                bytes[7] = (byte) crc;
                bytes[8] = (byte) (crc >> 8);
            }

            var name = legacyNames ? (i is 0 ? "old.rar" : $"old.r{i - 1:00}") : $"old.part{i + 1:0000}.rar.fake";
            File.WriteAllBytes(Path.Combine(dir, name), bytes);
        }

        var set = ArchiveDiscovery.Scan(dir).Single();
        Assert(set.Parts.Count is 3, "RAR4 grouping failed");
        var result = Extract(set, rar, null, out var output);
        Assert(result.Succeeded, result.ToString());
        Assert(result.VerifiedParts?.Count is 3, "RAR4 volumes not tracked");
        Assert(Directory.GetFiles(output).Length is 1 && new FileInfo(Directory.GetFiles(output).Single()).Length is 20111, "RAR4 content incomplete");
    });
Test("ZIP: empty archive is success", () =>
{
    var dir = Case("empty");
    using (ZipFile.Open(Path.Combine(dir, "empty.bin"), ZipArchiveMode.Create))
    {
    }

    var result = Extract(ArchiveDiscovery.Scan(dir).Single(), zip, null, out _);
    Assert(result.Succeeded, result.ToString());
});
Test("ZIP: corrupt payload never succeeds", () =>
{
    var dir = Generate("corrupt", "zip");
    var path = Directory.GetFiles(dir).Single();
    var bytes = File.ReadAllBytes(path);
    bytes[500] ^= 0xff;
    File.WriteAllBytes(path, bytes);
    Assert(!Extract(ArchiveDiscovery.Scan(dir).Single(), zip, null, out _).Succeeded, "Corrupt archive accepted");
});
Test("Cancellation", () =>
{
    var set = ArchiveDiscovery.Scan(Generate("cancel", "7z")).Single();
    Assert(Extract(set, zip, null, out _, new CancellationToken(true)).Status == ExtractionStatus.Cancelled, "Cancellation ignored");
});
Test("7z: solid archive with several files and empty directories", () =>
{
    var dir = Case("solid");
    var source = Path.Combine(dir, "source");
    _ = Directory.CreateDirectory(Path.Combine(source, "empty"));
    for (var i = 0; i < 5; i++)
        File.WriteAllBytes(Path.Combine(source, $"{i}.bin"), payload);
    Command(zipExe, source, "a", "-t7z", "-ms=on", Path.Combine(dir, "solid.7z"), "*");
    var set = ArchiveDiscovery.Scan(dir, false).Single();
    var result = Extract(set, zip, null, out var output);
    Assert(result.Succeeded, result.ToString());
    Assert(Directory.Exists(Path.Combine(output, "empty")), "Empty directory missing");
    for (var i = 0; i < 5; i++)
        Assert(File.ReadAllBytes(Path.Combine(output, $"{i}.bin")).SequenceEqual(payload), "Solid content mismatch");
});
Test("7z: unrelated extra numbered part cannot be accepted for cleanup", () =>
{
    var dir = Generate("extra-part", "7z", false, true);
    var count = Directory.GetFiles(dir).Length;
    File.WriteAllText(Path.Combine(dir, $"archive.7z.{count + 1:000}"), "unrelated file");
    var result = Extract(ArchiveDiscovery.Scan(dir).Single(), zip, null, out _);
    Assert(!result.Succeeded, "Unrelated trailing part accepted");
});
Test("RAR: missing last volume reported by native callback", () =>
{
    var dir = Generate("rar-missing-last", "rar", true, true);
    File.Delete(Directory.GetFiles(dir).Order().Last());
    var result = Extract(ArchiveDiscovery.Scan(dir).Single(), rar, password, out _);
    Assert(result.Status == ExtractionStatus.MissingVolume, result.ToString());
});
Test("ZIP: output traversal rejected", () =>
{
    var dir = Case("traversal");
    using (var archive = ZipFile.Open(Path.Combine(dir, "bad.zip"), ZipArchiveMode.Create))
    using (var writer = new StreamWriter(archive.CreateEntry("../escape.txt").Open()))
        writer.Write("bad");
    Assert(!Extract(ArchiveDiscovery.Scan(dir).Single(), zip, null, out _).Succeeded, "Traversal accepted");
    Assert(!File.Exists(Path.Combine(root, "escape.txt")), "Escaped output");
});
Test("Nested mixed content, renamed encrypted child, transactional source preservation", () =>
{
    var dir = Case("nested");
    var content = Path.Combine(dir, "content");
    _ = Directory.CreateDirectory(content);
    var childDir = Generate("nested-child", "7z", true);
    File.Copy(Directory.GetFiles(childDir).Single(), Path.Combine(content, "child.txt"));
    File.WriteAllText(Path.Combine(content, "说明.txt"), "keep me");
    var input = Path.Combine(dir, "input");
    _ = Directory.CreateDirectory(input);
    var source = Path.Combine(input, "outer.unknown");
    ZipFile.CreateFromDirectory(content, source);
    var before = SHA256.HashData(File.ReadAllBytes(source));
    var options = new PipelineOptions(input, Path.Combine(dir, "complete"), true);
    var pipeline = new ExtractionPipeline([
        zip,
        rar
    ]);
    Assert(pipeline.Run(options, [password]) is 0, "Nested extraction failed");
    Assert(File.ReadAllBytes(Path.Combine(options.CompleteDirectory, "outer/child/内容.bin")).SequenceEqual(payload), "Missing nested payload");
    Assert(File.ReadAllText(Path.Combine(options.CompleteDirectory, "outer/说明.txt")) is "keep me", "Mixed content lost");
    Assert(!File.Exists(Path.Combine(options.CompleteDirectory, "outer/child.txt")), "Intermediate remains");
    Assert(SHA256.HashData(File.ReadAllBytes(source)).SequenceEqual(before), "Source modified");
    var failedOptions = options with
    {
        CompleteDirectory = Path.Combine(dir, "failed-complete"),
        MaxDepth = 1
    };
    Assert(pipeline.Run(failedOptions, [password]) is 1, "Depth limit ignored");
    Assert(!Directory.EnumerateFileSystemEntries(failedOptions.CompleteDirectory).Any(), "Partial output committed");
    Assert(File.Exists(source), "Source lost on failure");
    Assert(!Directory.EnumerateFiles(dir, "*.error.txt", SearchOption.AllDirectories).Any(), "Failure report was written");
});
foreach (var marker in new[]
{
    "Game.EXE",
    "Engine.DlL",
    "assets.BIN"
}

)
    Test($"Game protection: {marker}, nested archives and descendant folders", () =>
    {
        var dir = Case("game-" + marker);
        var content = Path.Combine(dir, "content");
        var game = Path.Combine(content, "Game");
        _ = Directory.CreateDirectory(Path.Combine(game, "Data/empty"));
        _ = Directory.CreateDirectory(Path.Combine(content, "Ordinary"));
        File.WriteAllText(Path.Combine(game, marker), "game marker");
        var archiveDir = Generate("game-inner-" + marker, "zip");
        var archiveBytes = File.ReadAllBytes(Directory.GetFiles(archiveDir).Single());
        File.WriteAllBytes(Path.Combine(game, "inside.zip"), archiveBytes);
        File.WriteAllBytes(Path.Combine(game, "Data/renamed.data"), archiveBytes);
        File.WriteAllBytes(Path.Combine(content, "Ordinary/renamed.data"), archiveBytes);
        var input = Path.Combine(dir, "input");
        _ = Directory.CreateDirectory(input);
        ZipFile.CreateFromDirectory(content, Path.Combine(input, "outer.zip"));
        var options = new PipelineOptions(input, Path.Combine(dir, "complete"), true);
        Assert(new ExtractionPipeline([
            zip,
            rar
        ]).Run(options, []) is 0, "Game transaction failed");
        var outputGame = Path.Combine(options.CompleteDirectory, "outer/Game");
        _ = FileSystemHelper.CleanEmptyDirectories(options.CompleteDirectory, false);
        FileSystemHelper.NormalizeNestedFoldersContainingName(new(options.CompleteDirectory), "新建文件夹", 20);
        FileSystemHelper.NormalizeRedundantNestedFolders(new(options.CompleteDirectory), 20);
        Assert(File.ReadAllBytes(Path.Combine(outputGame, "inside.zip")).SequenceEqual(archiveBytes), "Game archive modified");
        Assert(File.ReadAllBytes(Path.Combine(outputGame, "Data/renamed.data")).SequenceEqual(archiveBytes), "Descendant archive expanded");
        Assert(Directory.Exists(Path.Combine(outputGame, "Data/empty")), "Game empty directory removed");
        Assert(!Directory.Exists(Path.Combine(outputGame, "inside")), "Game archive extracted");
        Assert(File.Exists(Path.Combine(options.CompleteDirectory, "outer/Ordinary/renamed/内容.bin")), "Ordinary sibling not extracted");
    });
Test("Game protection: one-file game folder cannot be flattened by any merge rule", () =>
{
    var dir = Case("game-merge");
    var game = Path.Combine(dir, "新建文件夹/Game");
    _ = Directory.CreateDirectory(game);
    File.WriteAllText(Path.Combine(game, "play.EXE"), "game");
    FileSystemHelper.NormalizeNestedFoldersContainingName(new(dir), "新建文件夹", 20);
    FileSystemHelper.NormalizeRedundantNestedFolders(new(dir), 20, FileSystemHelper.RedundantThreshold.Always);
    Assert(File.Exists(Path.Combine(game, "play.EXE")), "Single-file game folder flattened");
});
Test("Game protection: root input and normalization root", () =>
{
    var dir = Case("game-root");
    var input = Path.Combine(dir, "input");
    _ = Directory.CreateDirectory(Path.Combine(input, "wrapper/empty"));
    File.WriteAllText(Path.Combine(input, "engine.DLL"), "game");
    using (ZipFile.Open(Path.Combine(input, "wrapper/data.zip"), ZipArchiveMode.Create))
    {
    }

    var options = new PipelineOptions(input, Path.Combine(dir, "complete"), true);
    Assert(new ExtractionPipeline([
        zip,
        rar
    ]).Run(options, []) is 0, "Skipped game counted as failure");
    Assert(!Directory.EnumerateFileSystemEntries(options.CompleteDirectory).Any(), "Protected input extracted");
    FileSystemHelper.NormalizeRedundantNestedFolders(new(input), 20, FileSystemHelper.RedundantThreshold.Always);
    Assert(File.Exists(Path.Combine(input, "wrapper/data.zip")), "Game root's child flattened");
    Assert(Directory.Exists(Path.Combine(input, "wrapper/empty")), "Game root's empty child removed");
});
Test("Folder normalization helpers retained", () =>
{
    var dir = Case("normalize");
    _ = Directory.CreateDirectory(Path.Combine(dir, "Example/新建文件夹/Example"));
    File.WriteAllText(Path.Combine(dir, "Example/新建文件夹/Example/a.txt"), "a");
    File.WriteAllText(Path.Combine(dir, "Example/新建文件夹/Example/b.txt"), "b");
    FileSystemHelper.NormalizeNestedFoldersContainingName(new(dir), "新建文件夹", 20);
    FileSystemHelper.NormalizeRedundantNestedFolders(new(dir), 20);
    Assert(Directory.GetFiles(dir, "*.txt", SearchOption.AllDirectories).Length is 2, "Normalization lost files");
    Assert(Directory.GetDirectories(dir, "*新建文件夹*", SearchOption.AllDirectories).Length is 0, "Wrapper remains");
});
foreach (var format in new[]
{
    "zip",
    "rar"
}

)
    foreach (var interactive in new[]
    {
        false,
        true
    }

    )
        Test($"Logging: {format}, interactive={interactive}", () =>
        {
            var dir = Generate($"logging-{format}-{interactive}", format, true);
            var archive = ArchiveDiscovery.Scan(dir).Single();
            LoggingChecks.Verify(format is "rar" ? rar : zip, archive, Password(format), dir, interactive);
        });
Console.WriteLine($"{tests - failures}/{tests} passed. Fixtures: {root}");
Environment.ExitCode = failures is 0 ? 0 : 1;
static uint Crc(ReadOnlySpan<byte> data)
{
    var crc = uint.MaxValue;
    foreach (var b in data)
    {
        crc ^= b;
        for (var j = 0; j < 8; j++)
            crc = (crc >> 1) ^ ((crc & 1) is not 0 ? 0xedb88320u : 0);
    }

    return ~crc;
}

static byte[] DecodeUu(string[] lines)
{
    using var output = new MemoryStream();
    foreach (var line in lines.SkipWhile(l => !l.StartsWith("begin ")).Skip(1))
    {
        if (line is "end")
            break;
        if (line.Length is 0)
            continue;
        var count = (line[0] - 32) & 63;
        for (var pos = 1; count > 0; pos += 4)
        {
            var a = (line[pos] - 32) & 63;
            var b = (line[pos + 1] - 32) & 63;
            var c = (line[pos + 2] - 32) & 63;
            var d = (line[pos + 3] - 32) & 63;
            output.WriteByte((byte) ((a << 2) | (b >> 4)));
            count--;
            if (count > 0)
            {
                output.WriteByte((byte) ((b << 4) | (c >> 2)));
                count--;
            }

            if (count > 0)
            {
                output.WriteByte((byte) ((c << 6) | d));
                count--;
            }
        }
    }

    return output.ToArray();
}

static void MakeSplitZip(string dir, byte[] data)
{
    // One stored entry crossing disks, central directory on the final disk (PKZIP split format).
    var crc = uint.MaxValue;
    foreach (var b in data)
    {
        crc ^= b;
        for (var j = 0; j < 8; j++)
            crc = (crc >> 1) ^ ((crc & 1) is not 0 ? 0xedb88320u : 0);
    }

    crc = ~crc;
    var name = Encoding.UTF8.GetBytes("内容.bin");
    using var raw = new MemoryStream();
    using var writer = new BinaryWriter(raw, Encoding.UTF8, true);
    writer.Write(0x08074b50u); // split marker
    writer.Write(0x04034b50u);
    writer.Write((ushort) 20);
    writer.Write((ushort) 0x800);
    writer.Write((ushort) 0);
    writer.Write(0u);
    writer.Write(crc);
    writer.Write((uint) data.Length);
    writer.Write((uint) data.Length);
    writer.Write((ushort) name.Length);
    writer.Write((ushort) 0);
    writer.Write(name);
    writer.Write(data);
    var bytes = raw.ToArray();
    var diskCount = (bytes.Length + 65535) / 65536;
    for (var i = 0; i < diskCount; i++)
        File.WriteAllBytes(Path.Combine(dir, $"archive.z{i + 1:00}.fake"), bytes.AsSpan(i * 65536, Math.Min(65536, bytes.Length - i * 65536)).ToArray());
    using var last = new MemoryStream();
    using var w = new BinaryWriter(last, Encoding.UTF8, true);
    w.Write(0x02014b50u);
    w.Write((ushort) 20);
    w.Write((ushort) 20);
    w.Write((ushort) 0x800);
    w.Write((ushort) 0);
    w.Write(0u);
    w.Write(crc);
    w.Write((uint) data.Length);
    w.Write((uint) data.Length);
    w.Write((ushort) name.Length);
    w.Write((ushort) 0);
    w.Write((ushort) 0);
    w.Write((ushort) 0);
    w.Write((ushort) 0);
    w.Write(0u);
    w.Write(4u);
    w.Write(name);
    var centralSize = (uint) last.Length;
    w.Write(0x06054b50u);
    w.Write((ushort) diskCount);
    w.Write((ushort) diskCount);
    w.Write((ushort) 1);
    w.Write((ushort) 1);
    w.Write(centralSize);
    w.Write(0u);
    w.Write((ushort) 0);
    File.WriteAllBytes(Path.Combine(dir, "archive.wrong"), last.ToArray());
}

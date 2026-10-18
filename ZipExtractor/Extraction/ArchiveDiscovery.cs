using System;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using ZipExtractor.Logging;

namespace ZipExtractor.Extraction;

public sealed record ArchiveHeader(ArchiveFormat Format, bool IsVolume = false, int? VolumeIndex = null, bool ZipSplitMarker = false, int? ZipLastDisk = null, bool TailOnly = false, bool EncryptedHeaders = false);
/// <summary>Signatures decide format; names only supply volume association/order hints.</summary>
public sealed partial class ArchiveDiscovery
{
    private sealed record Item(string Path, ArchiveHeader Header, string? Key, int? Index, string Kind);

    public static IReadOnlyList<ArchiveSet> Scan(string directory, bool recursive = true, bool skipGameFolders = false, IArchiveProbe? probe = null, Action<DiscoveryProgress>? progress = null, CancellationToken token = default)
    {
        var result = new List<ArchiveSet>();
        var filesScanned = 0;
        void Report(string stage, string path) => progress?.Invoke(new(stage, path, filesScanned, result.Count));
        foreach (var group in EnumerateDirectories(new(directory), recursive, skipGameFolders, Report, token))
        {
            var items = new List<Item>();
            foreach (var file in group.OfType<FileInfo>())
            {
                token.ThrowIfCancellationRequested();
                if (file.Name.EndsWith(".downloading", StringComparison.OrdinalIgnoreCase))
                    continue;
                Report("读取文件头", file.FullName);
                items.Add(ReadItem(file.FullName, probe, Report, token));
                filesScanned++;
            }
            Report("关联分卷", Path.GetDirectoryName(group.FirstOrDefault()?.FullName) ?? directory);
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in items.OrderBy(i => i.Index ?? 0))
            {
                token.ThrowIfCancellationRequested();
                if (used.Contains(item.Path) || item.Header.Format == ArchiveFormat.Unknown)
                    continue;
                var members = item.Key is null ? [item] : items.Where(i => string.Equals(i.Key, item.Key, StringComparison.OrdinalIgnoreCase)).OrderBy(i => i.Index).ToArray();
                var format = item.Header.Format;
                // A numeric filename is not enough to join independent, complete archives.
                if (format != ArchiveFormat.Rar && members.Count(i => i.Header.Format != ArchiveFormat.Unknown && !i.Header.TailOnly) > 1 && !members.Any(i => i.Header.ZipSplitMarker || i.Header.ZipLastDisk > 0))
                    members = [item];
                if (format == ArchiveFormat.Rar && item.Header is { IsVolume: false, EncryptedHeaders: false })
                    members = [item];
                var parts = members.Select(i => i.Path).ToArray();
                var aliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var m in members)
                    aliases[Path.GetFileName(m.Path)] = m.Path;
                string? error = null;
                var entry = members[0];
                var layout = members.Length is 1 ? VolumeLayout.Single : VolumeLayout.Concatenated;
                var display = Path.GetFileNameWithoutExtension(item.Path);
                if (item.Key is not null)
                    display = Path.GetFileName(item.Key.Split('|')[0]);
                if (format == ArchiveFormat.Rar)
                {
                    if (members.All(i => i.Header.VolumeIndex.HasValue))
                    {
                        members = [.. members.OrderBy(i => i.Header.VolumeIndex)];
                        parts = [.. members.Select(i => i.Path)];
                        if (members.Select(i => i.Header.VolumeIndex!.Value).Where((n, i) => n != i).Any())
                            error = "RAR 文件头中的卷号不连续或重复。";
                    }

                    layout = item.Header.IsVolume || members.Length > 1 ? VolumeLayout.Native : VolumeLayout.Single;
                    entry = members.FirstOrDefault(i => i.Header.VolumeIndex is 0) ?? members[0];
                    if (entry.Header.VolumeIndex > 0 || (entry.Header.VolumeIndex is null && entry is { Index: > 1, Kind: "part" }) || entry.Kind is "oldrar")
                        error = "缺少 RAR 第一卷，或第一卷无法识别。";
                    if (members.Any(i => i.Header.Format != ArchiveFormat.Rar))
                        error = "RAR 卷组中存在无法识别的文件。";
                    foreach (var m in members)
                    {
                        if (m.Index is not { } n)
                            continue;
                        aliases[$"{display}.part{n}.rar"] = m.Path;
                        aliases[$"{display}.part{n:00}.rar"] = m.Path;
                        aliases[$"{display}.part{n:000}.rar"] = m.Path;
                        // The actual next-volume name is resolved by ordinal as well (see ResolveRar).
                    }
                }
                else if (format == ArchiveFormat.Zip && members.Any(i => i.Header.ZipSplitMarker || i.Header.ZipLastDisk > 0))
                {
                    layout = VolumeLayout.Native;
                    entry = members.FirstOrDefault(i => i.Kind is "zipend" || i.Header.ZipLastDisk > 0) ?? members[^1];
                    aliases[$"{display}.zip"] = entry.Path;
                    var disks = members.Where(i => i.Path != entry.Path).ToArray();
                    for (var i = 0; i < disks.Length; i++)
                        aliases[$"{display}.z{i + 1:00}"] = disks[i].Path;
                    if (entry.Header.ZipLastDisk is { } last && last != disks.Length)
                        error = $"ZIP 分卷缺失：目录记录要求 {last + 1} 卷，发现 {members.Length} 卷。";
                }

                var numberedMembers = members.Where(i => i.Kind is not "zipend").ToArray();
                if (!(format == ArchiveFormat.Rar && members.All(i => i.Header.VolumeIndex.HasValue)) && numberedMembers.Length > 1 && numberedMembers.All(i => i.Index.HasValue))
                {
                    var indices = numberedMembers.Select(i => i.Index!.Value).ToArray();
                    if (indices.Distinct().Count() != indices.Length || indices[^1] - indices[0] + 1 != indices.Length)
                        error = "分卷编号重复或不连续，无法可靠确定卷顺序。";
                }

                if (format != ArchiveFormat.Rar && entry.Index > 1 && layout != VolumeLayout.Native)
                    error = "缺少编号为 001 的第一卷。";
                foreach (var m in members)
                    _ = used.Add(m.Path);
                result.Add(new(format, entry.Path, display, parts, layout, aliases, error));
            }
        }

        Report("扫描完成", directory);
        return result;
    }

    private static IEnumerable<FileSystemInfo[]> EnumerateDirectories(DirectoryInfo directory, bool recursive, bool skipGameFolders, Action<string, string> report, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        report("枚举目录", directory.FullName);
        var entries = new List<FileSystemInfo>();
        foreach (var entry in directory.EnumerateFileSystemInfos())
        {
            token.ThrowIfCancellationRequested();
            entries.Add(entry);
        }
        if (skipGameFolders && FileSystemHelper.IsGameFolder(entries))
        {
            ConsoleLog.Default.GameSkipped(directory.FullName);
            yield break;
        }

        yield return [.. entries];
        if (recursive)
            foreach (var child in entries.OfType<DirectoryInfo>())
            {
                token.ThrowIfCancellationRequested();
                // Junctions can point back to an ancestor or outside the input tree.
                if ((child.Attributes & FileAttributes.ReparsePoint) is not 0)
                    continue;
                foreach (var files in EnumerateDirectories(child, true, skipGameFolders, report, token))
                    yield return files;
            }
    }

    private static Item ReadItem(string path, IArchiveProbe? probe, Action<string, string> report, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        var header = ReadHeader(path);
        if (header.Format is ArchiveFormat.Unknown && probe is not null)
        {
            report("深度探测（可能读取整个文件）", path);
            token.ThrowIfCancellationRequested();
            var format = probe.Probe(path, token);
            if (format is not ArchiveFormat.Unknown)
                return new(path, new(format), null, null, "single");
        }
        return MakeItem(path, header);
    }

    private static Item MakeItem(string path, ArchiveHeader header)
    {
        var name = Path.GetFileName(path);
        var dir = Path.GetDirectoryName(path)!;
        // Accept wrong/appended suffixes: name.part01.bin, name.7z.001.txt, name.z01.download.
        var part = PartVolumeRegex.Match(name);
        if (part.Success && int.TryParse(part.Groups[2].Value, out var p))
            return new(path, header, Path.Combine(dir, part.Groups[1].Value) + "|part", p, "part");
        var zip = ZipVolumeRegex.Match(name);
        if (header.Format != ArchiveFormat.Rar && zip.Success && (zip.Groups[2].Success || header.ZipLastDisk > 0))
            return new(path, header, Path.Combine(dir, zip.Groups[1].Value) + "|zip", zip.Groups[2].Success ? int.Parse(zip.Groups[2].Value) : int.MaxValue, zip.Groups[2].Success ? "zipvol" : "zipend");
        var numeric = NumericVolumeRegex.Match(name);
        if (numeric.Success && int.TryParse(numeric.Groups[2].Value, out var n))
            return new(path, header, Path.Combine(dir, numeric.Groups[1].Value) + "|numeric", n, "numeric");
        if (header.ZipLastDisk > 0)
            return new(path, header, Path.Combine(dir, Path.GetFileNameWithoutExtension(name)) + "|zip", int.MaxValue, "zipend");
        var old = LegacyRarVolumeRegex.Match(name);
        if (old.Success && header.Format == ArchiveFormat.Rar)
            return new(path, header, Path.Combine(dir, old.Groups[1].Value) + "|oldrar", (char.ToLowerInvariant(old.Groups[2].Value[0]) - 'r') * 100 + int.Parse(old.Groups[3].Value) + 1, "oldrar");
        if (header is { Format: ArchiveFormat.Rar, IsVolume: true })
        {
            var stem = RarSuffixRegex.Replace(name, "");
            if (stem == name)
                stem = Path.GetFileNameWithoutExtension(name);
            return new(path, header, Path.Combine(dir, stem) + "|oldrar", header.VolumeIndex ?? 0, "rarstart");
        }

        return new(path, header, null, null, "single");
    }

    public static ArchiveHeader ReadHeader(string path)
    {
        using var file = File.OpenRead(path);
        var data = new byte[(int) Math.Min(file.Length, 64)];
        file.ReadExactly(data);
        var sfx = data is [0x4d, 0x5a, ..];
        if (sfx)
        {
            Array.Resize(ref data, (int) Math.Min(file.Length, 1024 * 1024));
            file.Position = 0;
            file.ReadExactly(data);
        }

        var span = data.AsSpan();
        for (var offset = 0; offset < (sfx ? span.Length : Math.Min(1, span.Length)); offset++)
        {
            var s = span[offset..];
            if (s.StartsWith(new byte[] { 0x37, 0x7a, 0xbc, 0xaf, 0x27, 0x1c }))
                return new(ArchiveFormat.SevenZip);
            if (s.StartsWith(new byte[] { 0x52, 0x61, 0x72, 0x21, 0x1a, 0x07, 0x00 }))
            {
                if (s.Length < 14)
                    return new(ArchiveFormat.Rar);
                var flags = BinaryPrimitives.ReadUInt16LittleEndian(s[10..]);
                return new(ArchiveFormat.Rar, (flags & 1) is not 0, (flags & 1) is 0 || (flags & 0x100) is not 0 ? 0 : null);
            }

            if (s.StartsWith(new byte[] { 0x52, 0x61, 0x72, 0x21, 0x1a, 0x07, 0x01, 0x00 }))
            {
                try
                {
                    var at = 12;
                    _ = VInt(s, ref at);
                    var type = VInt(s, ref at);
                    var flags = VInt(s, ref at);
                    if ((flags & 1) is not 0)
                        _ = VInt(s, ref at);
                    if ((flags & 2) is not 0)
                        _ = VInt(s, ref at);
                    if (type is not 1)
                        return new(ArchiveFormat.Rar, EncryptedHeaders: type is 4);
                    var arcFlags = VInt(s, ref at);
                    var index = (arcFlags & 2) is not 0 ? checked((int) VInt(s, ref at)) : 0;
                    return new(ArchiveFormat.Rar, (arcFlags & 1) is not 0, index);
                }
                catch (Exception e) when (e is InvalidDataException or OverflowException)
                {
                    return new(ArchiveFormat.Rar);
                }
            }

            if (s.StartsWith(new byte[] { 0x50, 0x4b, 3, 4 }) || s.StartsWith(new byte[] { 0x50, 0x4b, 5, 6 }) || s.StartsWith(new byte[] { 0x50, 0x4b, 7, 8 }))
                return new(ArchiveFormat.Zip, ZipSplitMarker: s[2] is 7, ZipLastDisk: ZipDisk(file));
        }

        // The final disk of a split ZIP need not start with a ZIP signature.
        var disk = ZipDisk(file);
        return disk.HasValue ? new(ArchiveFormat.Zip, ZipLastDisk: disk, TailOnly: true) : new(ArchiveFormat.Unknown);
    }

    private static int? ZipDisk(FileStream file)
    {
        var size = (int) Math.Min(file.Length, 65557);
        var tail = new byte[size];
        file.Position = file.Length - size;
        file.ReadExactly(tail);
        for (var i = tail.Length - 22; i >= 0; i--)
            if (tail.AsSpan(i).StartsWith(new byte[] { 0x50, 0x4b, 5, 6 }) && i + 22 + BinaryPrimitives.ReadUInt16LittleEndian(tail.AsSpan(i + 20)) == tail.Length)
                return BinaryPrimitives.ReadUInt16LittleEndian(tail.AsSpan(i + 4));
        return null;
    }

    private static ulong VInt(ReadOnlySpan<byte> data, ref int at)
    {
        ulong result = 0;
        for (var shift = 0; shift < 64 && at < data.Length; shift += 7)
        {
            var b = data[at++];
            result |= (ulong) (b & 127) << shift;
            if ((b & 128) is 0)
                return result;
        }

        throw new InvalidDataException("无效 RAR 头");
    }

    [GeneratedRegex(@"^(.*?)[._ -]part(\d+)(?:\..*)?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex PartVolumeRegex { get; }

    [GeneratedRegex(@"^(.*?)\.(?:z(\d{2,})|zip)(?:\..*)?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex ZipVolumeRegex { get; }

    [GeneratedRegex(@"^(.*?)\.(\d{3,})(?:\..*)?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex NumericVolumeRegex { get; }

    [GeneratedRegex(@"^(.*?)\.([r-z])(\d{2})(?:\..*)?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex LegacyRarVolumeRegex { get; }

    [GeneratedRegex(@"\.rar(?:\..*)?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex RarSuffixRegex { get; }
}

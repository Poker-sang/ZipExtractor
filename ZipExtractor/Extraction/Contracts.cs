using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace ZipExtractor.Extraction;

public enum ArchiveFormat
{
    Unknown,
    Zip,
    SevenZip,
    Rar
}

public enum VolumeLayout
{
    Single,
    Concatenated,
    Native
}

public enum ExtractionStatus
{
    Success,
    WrongPassword,
    PasswordOrCorruptData,
    MissingVolume,
    InvalidArchive,
    Unsupported,
    IoError,
    Cancelled
}

public sealed record ArchiveSet(ArchiveFormat Format, string EntryPath, string DisplayName, IReadOnlyList<string> Parts, VolumeLayout Layout, IReadOnlyDictionary<string, string> VolumeAliases, string? DiscoveryError = null);
public sealed record ExtractionProgress(string Engine, string? Entry, ulong Completed, ulong? Total);
public sealed record ExtractionResult(ExtractionStatus Status, string Message, IReadOnlyList<string>? VerifiedParts = null)
{
    public bool Succeeded => Status == ExtractionStatus.Success;

    public static ExtractionResult Success(IReadOnlyList<string>? verifiedParts = null) => new(ExtractionStatus.Success, "校验通过", verifiedParts);
}

/// <summary>Each call is one password attempt, writing into a new, empty directory.</summary>
public sealed record ExtractionRequest(ArchiveSet Archive, string OutputDirectory, string? Password, IProgress<ExtractionProgress>? Progress = null, CancellationToken CancellationToken = default);
/// <summary>Replaceable extraction backend. Implementations validate every extracted item.</summary>
public interface IArchiveExtractor : IDisposable
{
    string Name { get; }

    bool Supports(ArchiveFormat format);

    ExtractionResult Extract(ExtractionRequest request);
}

/// <summary>Native format probing for files not identified by their outer headers.</summary>
public interface IArchiveProbe
{
    ArchiveFormat Probe(string path, CancellationToken token = default);
}

/// <summary>Replace this to support a user-supplied mapping for arbitrarily renamed volumes.</summary>
public interface IVolumeResolver
{
    string? Resolve(ArchiveSet archive, string requestedName);
}

public sealed class VolumeResolver : IVolumeResolver
{
    public string? Resolve(ArchiveSet archive, string requestedName)
    {
        var name = Path.GetFileName(requestedName);
        if (archive.VolumeAliases.TryGetValue(name, out var path) && File.Exists(path))
            return path;
        // Never search other directories or guess among unrelated files.
        foreach (var part in archive.Parts)
            if (string.Equals(Path.GetFileName(part), name, StringComparison.OrdinalIgnoreCase) && File.Exists(part))
                return part;
        return null;
    }
}

internal static class OutputPaths
{
    public static string Get(string root, string entry)
    {
        entry = entry.Replace('/', Path.DirectorySeparatorChar);
        if (Path.IsPathRooted(entry) || entry.Contains(':') || entry.Contains('\0'))
            throw new InvalidDataException($"无效的压缩包路径：{entry}");
        var components = entry.Split(Path.DirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries).Where(c => c is not ".").ToArray();
        foreach (var component in components)
            if (component is ".." || component.EndsWith(' ') || component.EndsWith('.'))
                throw new InvalidDataException($"无效的压缩包路径：{entry}");
        entry = string.Join(Path.DirectorySeparatorChar, components);
        var fullRoot = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        if (entry.Length is 0)
            return fullRoot;
        var path = Path.GetFullPath(Path.Combine(fullRoot, entry));
        return !path.StartsWith(fullRoot, StringComparison.OrdinalIgnoreCase) ? throw new InvalidDataException($"压缩包路径越界：{entry}") : path;
    }
}

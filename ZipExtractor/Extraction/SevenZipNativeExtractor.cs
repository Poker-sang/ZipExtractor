using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using System.Threading;
using ZipExtractor.Extraction.Native;

namespace ZipExtractor.Extraction;

public sealed unsafe partial class SevenZipNativeExtractor : IArchiveExtractor, IArchiveProbe
{
    private static readonly StrategyBasedComWrappers _ComWrappers = new();
    private readonly nint _library;
    private readonly CreateObject _create;
    private readonly IVolumeResolver _resolver;

    public string Name => "7-Zip";

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int CreateObject(ref Guid clsid, ref Guid iid, out nint instance);

    public SevenZipNativeExtractor(string dllPath, IVolumeResolver? resolver = null)
    {
        _resolver = resolver ?? new VolumeResolver();
        _library = NativeLibrary.Load(Path.GetFullPath(dllPath));
        _create = Marshal.GetDelegateForFunctionPointer<CreateObject>(NativeLibrary.GetExport(_library, "CreateObject"));
    }

    public bool Supports(ArchiveFormat format) => format is ArchiveFormat.Zip or ArchiveFormat.SevenZip;

    public ExtractionResult Extract(ExtractionRequest request)
    {
        if (!Supports(request.Archive.Format))
            return new(ExtractionStatus.Unsupported, "7-Zip 后端未配置此格式。");
        if (request.Archive.DiscoveryError is { } problem)
            return new(ExtractionStatus.MissingVolume, problem);
        IInArchive? archive = null;
        ComObject? archiveOwner = null;
        using var open = new OpenCallback(request, _resolver);
        try
        {
            request.CancellationToken.ThrowIfCancellationRequested();
            using var input = new NativeInputStream(request.Archive.Layout == VolumeLayout.Concatenated ? new ConcatenatedStream(request.Archive.Parts) : File.OpenRead(request.Archive.EntryPath), request.CancellationToken);
            archiveOwner = CreateArchive(request.Archive.Format);
            archive = (IInArchive) (object) archiveOwner;
            // Let the native handler locate embedded data anywhere in this input stream.
            var maxSearch = (ulong) input.Length;
            var hr = archive.Open(input, &maxSearch, open);
            if (request.CancellationToken.IsCancellationRequested)
                return new(ExtractionStatus.Cancelled, "已取消");
            if (open.Missing is { } missing)
                return new(ExtractionStatus.MissingVolume, $"缺少分卷：{missing}");
            if (open.Error is { } openError)
                return new(ExtractionStatus.IoError, openError.Message);
            if (input.Error is { } readError)
                return new(ExtractionStatus.IoError, readError.Message);
            if (hr is not 0)
                return new(open.PasswordRequested ? ExtractionStatus.PasswordOrCorruptData : ExtractionStatus.InvalidArchive, $"无法打开压缩包（0x{hr:X8}）" + (open.PasswordRequested ? "，密码不正确或加密头损坏。" : "。"));
            if (request.Archive.Layout == VolumeLayout.Concatenated)
            {
                using var physicalSize = GetArchiveProperty(archive, ArchiveProperty.PhysicalSize);
                using var offset = GetArchiveProperty(archive, ArchiveProperty.Offset);
                if (physicalSize.Unsigned is { } size && size + (offset.Unsigned ?? 0) != (ulong) input.Length)
                    return new(ExtractionStatus.InvalidArchive, "分卷总长度与引擎识别的压缩包长度不一致，可能包含无关分片。");
            }

            using var callback = new ExtractCallback(archive, request);
            hr = archive.Extract(null, uint.MaxValue, ArchiveExtractMode.Extract, callback);
            if (request.CancellationToken.IsCancellationRequested)
                return new(ExtractionStatus.Cancelled, "已取消");
            if (open.Missing is { } missing2)
                return new(ExtractionStatus.MissingVolume, $"缺少分卷：{missing2}");
            if (callback.Error is { } error)
                return new(ExtractionStatus.IoError, error.Message);
            if (input.Error is { } inputError)
                return new(ExtractionStatus.IoError, inputError.Message);
            if (callback.Failure is { } failure)
                return failure;
            if (hr is not 0)
                return new(ExtractionStatus.InvalidArchive, $"解压失败（0x{hr:X8}）。");
            using var flags = GetArchiveProperty(archive, ArchiveProperty.ErrorFlags); // kpidErrorFlags
            if (flags.Unsigned is > 0)
                return new(ExtractionStatus.InvalidArchive, $"压缩包结构校验失败：0x{flags.Number:X}");
            using var warnings = GetArchiveProperty(archive, ArchiveProperty.WarningFlags);
            if (warnings.Unsigned is > 0)
                return new(ExtractionStatus.InvalidArchive, $"压缩包结构警告：0x{warnings.Number:X}，保留源文件。");
            var verified = request.Archive.Layout == VolumeLayout.Concatenated ? request.Archive.Parts : [.. open.OpenedParts.Append(request.Archive.EntryPath).Distinct(StringComparer.OrdinalIgnoreCase)];
            return ExtractionResult.Success(verified);
        }
        catch (OperationCanceledException)
        {
            return new(ExtractionStatus.Cancelled, "已取消");
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or COMException)
        {
            return new(ExtractionStatus.IoError, e.Message);
        }
        finally
        {
            try
            {
                if (archive is not null)
                    _ = archive.Close();
            }
            finally
            {
                // Release our unique wrapper after native calls finish, before unloading the DLL.
                archiveOwner?.FinalRelease();
            }
        }
    }

    public ArchiveFormat Probe(string path, CancellationToken token = default)
    {
        foreach (var format in new[] { ArchiveFormat.Zip, ArchiveFormat.SevenZip })
        {
            token.ThrowIfCancellationRequested();
            var set = new ArchiveSet(format, path, Path.GetFileNameWithoutExtension(path), [path], VolumeLayout.Single, new Dictionary<string, string>());
            using var input = new NativeInputStream(File.OpenRead(path), token);
            using var open = new OpenCallback(new(set, "", null, CancellationToken: token), _resolver);
            var owner = CreateArchive(format);
            IInArchive? archive = null;
            try
            {
                archive = (IInArchive) (object) owner;
                var maxSearch = (ulong) input.Length;
                var hr = archive.Open(input, &maxSearch, open);
                token.ThrowIfCancellationRequested();
                if (input.Error is { } readError)
                    throw new IOException($"无法探测压缩包：{path}", readError);
                // Encrypted headers can identify the handler even when opening needs a password.
                if (open.PasswordRequested)
                    return format;
                if (hr is not 0 || open.Error is not null || open.Missing is not null)
                    continue;
                using var errors = GetArchiveProperty(archive, ArchiveProperty.ErrorFlags);
                if (errors.Unsigned is null or 0)
                    return format;
            }
            finally
            {
                try
                {
                    if (archive is not null)
                        _ = archive.Close();
                }
                finally
                {
                    owner.FinalRelease();
                }
            }
        }

        return ArchiveFormat.Unknown;
    }

    private ComObject CreateArchive(ArchiveFormat format)
    {
        var clsid = new Guid(format == ArchiveFormat.Zip ? "23170F69-40C1-278A-1000-000110010000" : "23170F69-40C1-278A-1000-000110070000");
        var iid = typeof(IInArchive).GUID;
        Marshal.ThrowExceptionForHR(_create(ref clsid, ref iid, out var ptr));
        try
        {
            // Each operation owns its unique wrapper and releases it before unloading the DLL.
            return (ComObject) _ComWrappers.GetOrCreateObjectForComInstance(ptr, CreateObjectFlags.UniqueInstance);
        }
        finally
        {
            _ = Marshal.Release(ptr);
        }
    }

    internal static PropVariant GetProperty(IInArchive archive, uint index, ArchiveProperty property)
    {
        var value = new PropVariant();
        Marshal.ThrowExceptionForHR(archive.GetProperty(index, property, ref value));
        return value;
    }

    private static PropVariant GetArchiveProperty(IInArchive archive, ArchiveProperty property)
    {
        var value = new PropVariant();
        Marshal.ThrowExceptionForHR(archive.GetArchiveProperty(property, ref value));
        return value;
    }

    public void Dispose() => NativeLibrary.Free(_library);

    [GeneratedComClass]
    public sealed partial class OpenCallback(ExtractionRequest request, IVolumeResolver resolver) : IArchiveOpenCallback, IArchiveOpenVolumeCallback, ICryptoGetTextPassword, IDisposable
    {
        private readonly List<NativeInputStream> _streams = [];

        public bool PasswordRequested { get; private set; }

        public string? Missing { get; private set; }

        public Exception? Error { get; private set; }

        public List<string> OpenedParts { get; } = [];

        public int SetTotal(ulong* files, ulong* bytes) => request.CancellationToken.IsCancellationRequested ? unchecked((int) 0x80004004) : 0;

        public int SetCompleted(ulong* files, ulong* bytes) => SetTotal(files, bytes);

        public int CryptoGetTextPassword(out string password)
        {
            PasswordRequested = true;
            password = request.Password ?? "";
            return request.Password is null ? unchecked((int) 0x80004004) : 0;
        }

        public int GetProperty(ArchiveProperty propId, ref PropVariant value)
        {
            value = default;
            if (propId is ArchiveProperty.Name)
            {
                // Give native split ZIP a canonical virtual name, regardless of its physical suffix.
                var name = request.Archive.Layout == VolumeLayout.Native ? request.Archive.DisplayName + ".zip" : Path.GetFileName(request.Archive.EntryPath);
                value.Type = PropVariantType.BStr;
                value.Pointer = Marshal.StringToBSTR(name);
            }

            if (propId is ArchiveProperty.Size)
            {
                value.Type = PropVariantType.UInt64;
                value.Number = (ulong) new FileInfo(request.Archive.EntryPath).Length;
            }

            return 0;
        }

        public int GetStream(string name, out IInStream? stream)
        {
            stream = null;
            try
            {
                request.CancellationToken.ThrowIfCancellationRequested();
                var path = resolver.Resolve(request.Archive, name);
                if (path is null)
                {
                    Missing = name;
                    return 1;
                }

                var input = new NativeInputStream(File.OpenRead(path), request.CancellationToken);
                _streams.Add(input);
                OpenedParts.Add(path);
                stream = input;
                return 0;
            }
            catch (Exception e)
            {
                Error = e;
                return Marshal.GetHRForException(e);
            }
        }

        public void Dispose()
        {
            foreach (var stream in _streams)
                stream.Dispose();
        }
    }

    [GeneratedComClass]
    public sealed partial class ExtractCallback : IArchiveExtractCallback, ICryptoGetTextPassword, IDisposable
    {
        private readonly IInArchive _archive;
        private readonly ExtractionRequest _request;
        private NativeOutputStream? _output;
        private ulong? _total;
        private string? _entry;
        private string? _targetPath;
        private long? _modified;
        private bool _encrypted;

        public Exception? Error { get; private set; }

        public ExtractionResult? Failure { get; private set; }

        internal ExtractCallback(IInArchive archive, ExtractionRequest request)
        {
            _archive = archive;
            _request = request;
        }

        public int SetTotal(ulong value)
        {
            _total = value;
            return Check();
        }

        public int SetCompleted(ulong* complete)
        {
            try
            {
                _request.Progress?.Report(new("7-Zip", _entry, complete is null ? 0 : *complete, _total));
                return Check();
            }
            catch (Exception e)
            {
                Error = e;
                return Marshal.GetHRForException(e);
            }
        }

        private int Check() => _request.CancellationToken.IsCancellationRequested ? unchecked((int) 0x80004004) : 0;

        public int CryptoGetTextPassword(out string password)
        {
            password = _request.Password ?? "";
            if (_request.Password is null)
            {
                Failure = new(ExtractionStatus.WrongPassword, "压缩包需要密码。");
                return unchecked((int) 0x80004004);
            }

            return Check();
        }

        public int GetStream(uint index, out ISequentialOutStream? stream, ArchiveAskMode mode)
        {
            stream = null;
            try
            {
                _output?.Dispose();
                _output = null;
                _request.CancellationToken.ThrowIfCancellationRequested();
                using var path = GetProperty(_archive, index, ArchiveProperty.Path);
                using var dir = GetProperty(_archive, index, ArchiveProperty.IsDirectory);
                using var encryption = GetProperty(_archive, index, ArchiveProperty.Encrypted);
                using var time = GetProperty(_archive, index, ArchiveProperty.ModifiedTime);
                _entry = path.Text ?? $"item-{index}";
                _encrypted = encryption.Boolean;
                _modified = time.Type is PropVariantType.FileTime ? (long) time.Number : null;
                _targetPath = null;
                if (mode is not ArchiveAskMode.Extract)
                    return 0;
                using var symlink = GetProperty(_archive, index, ArchiveProperty.SymbolicLink);
                using var hardlink = GetProperty(_archive, index, ArchiveProperty.HardLink);
                if (!string.IsNullOrEmpty(symlink.Text) || !string.IsNullOrEmpty(hardlink.Text))
                {
                    Failure = new(ExtractionStatus.Unsupported, $"暂不提取链接：{_entry}");
                    return unchecked((int) 0x80004005);
                }

                _targetPath = OutputPaths.Get(_request.OutputDirectory, _entry);
                if (dir.Boolean)
                    _ = Directory.CreateDirectory(_targetPath);
                else
                {
                    _ = Directory.CreateDirectory(Path.GetDirectoryName(_targetPath)!);
                    // An archive containing duplicate names fails explicitly instead of silently overwriting.
                    _output = new(new FileStream(_targetPath, FileMode.CreateNew, FileAccess.Write), _request.CancellationToken);
                    stream = _output;
                }

                return 0;
            }
            catch (Exception e)
            {
                Error = e;
                return Marshal.GetHRForException(e);
            }
        }

        public int PrepareOperation(ArchiveAskMode mode) => Check();

        public int SetOperationResult(ArchiveOperationResult result)
        {
            try
            {
                if (_output?.Error is { } e)
                    Error = e;
                _output?.Dispose();
                _output = null;
                if (result is not ArchiveOperationResult.Ok)
                {
                    var status = result switch
                    {
                        ArchiveOperationResult.WrongPassword => ExtractionStatus.WrongPassword,
                        ArchiveOperationResult.UnsupportedMethod => ExtractionStatus.Unsupported,
                        ArchiveOperationResult.DataError or ArchiveOperationResult.CrcError when _encrypted => ExtractionStatus.PasswordOrCorruptData,
                        _ => ExtractionStatus.InvalidArchive
                    };
                    Failure ??= new(status, $"{_entry}：7-Zip 校验错误 {result}。");
                }
                else if (_targetPath is not null && _modified is > 0 && File.Exists(_targetPath))
                    File.SetLastWriteTimeUtc(_targetPath, DateTime.FromFileTimeUtc(_modified.Value));
                return Check();
            }
            catch (Exception e)
            {
                Error = e;
                return Marshal.GetHRForException(e);
            }
        }

        public void Dispose()
        {
            _output?.Dispose();
        }
    }
}

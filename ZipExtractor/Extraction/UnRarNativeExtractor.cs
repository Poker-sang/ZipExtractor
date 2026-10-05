using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace ZipExtractor.Extraction;

public sealed unsafe class UnRarNativeExtractor : IArchiveExtractor
{
    private readonly nint _library;
    private readonly Open _open;
    private readonly Close _close;
    private readonly ReadHeader _read;
    private readonly ProcessFile _process;
    private readonly IVolumeResolver _resolver;

    public string Name => "UnRAR";

    public UnRarNativeExtractor(string dllPath, IVolumeResolver? resolver = null)
    {
        _resolver = resolver ?? new VolumeResolver();
        _library = NativeLibrary.Load(Path.GetFullPath(dllPath));
        _open = Load<Open>("RAROpenArchiveEx");
        _close = Load<Close>("RARCloseArchive");
        _read = Load<ReadHeader>("RARReadHeaderEx");
        _process = Load<ProcessFile>("RARProcessFileW");
    }

    private T Load<T>(string name)
            where T : Delegate => Marshal.GetDelegateForFunctionPointer<T>(NativeLibrary.GetExport(_library, name));

    public bool Supports(ArchiveFormat format) => format is ArchiveFormat.Rar;

    public ExtractionResult Extract(ExtractionRequest request)
    {
        if (!Supports(request.Archive.Format))
            return new(ExtractionStatus.Unsupported, "UnRAR 仅支持 RAR。");
        if (request.Archive.DiscoveryError is { } problem)
            return new(ExtractionStatus.MissingVolume, problem);
        var longName = (char*) NativeMemory.AllocZeroed(32768, sizeof(char));
        var handle = nint.Zero;
        ExtractionResult? callbackFailure = null;
        string? entry = null;
        ulong completed = 0;
        var encrypted = false;
        var currentPart = Array.FindIndex([.. request.Archive.Parts], p => p == request.Archive.EntryPath);
        var openedParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            request.Archive.EntryPath
        };
        var suppliedParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Callback callback = (msg, _1, p1, p2) =>
        {
            try
            {
                request.CancellationToken.ThrowIfCancellationRequested();
                switch (msg)
                {
                    case RarCallbackMessage.NeedPasswordWide:
                        if (request.Password is null)
                        {
                            callbackFailure = new(ExtractionStatus.WrongPassword, "压缩包需要密码。");
                            return -1;
                        }

                        WriteString((char*) p1, request.Password, p2.ToInt32());
                        return 1;
                    case RarCallbackMessage.ChangeVolumeWide:
                    {
                        var requested = Marshal.PtrToStringUni(p1)!;
                        if ((RarVolumeMode) p2.ToInt32() is RarVolumeMode.Notify)
                        {
                            var index = Array.FindIndex([.. request.Archive.Parts], p => string.Equals(p, requested, StringComparison.OrdinalIgnoreCase));
                            if (index >= 0)
                                currentPart = index;
                            if (File.Exists(requested))
                                _ = openedParts.Add(Path.GetFullPath(requested));
                            return 1;
                        }

                        var resolved = _resolver.Resolve(request.Archive, requested);
                        if (resolved is null && currentPart + 1 < request.Archive.Parts.Count)
                        {
                            // Discovery has already checked this group's ordinal continuity and RAR headers.
                            var candidate = request.Archive.Parts[currentPart + 1];
                            var header = ArchiveDiscovery.ReadHeader(candidate);
                            if (header is { Format: ArchiveFormat.Rar, IsVolume: true } && (header.VolumeIndex is null || header.VolumeIndex == currentPart + 1))
                                resolved = candidate;
                        }

                        if (resolved is null)
                        {
                            callbackFailure = new(ExtractionStatus.MissingVolume, $"缺少分卷：{requested}");
                            return -1;
                        }

                        if (!suppliedParts.Add(resolved))
                        {
                            callbackFailure = new(ExtractionStatus.InvalidArchive, $"引擎无法使用分卷：{resolved}");
                            return -1;
                        }

                        WriteString((char*) p1, resolved, 1024);
                        return 1;
                    }

                    case RarCallbackMessage.ProcessData:
                        completed += (ulong) p2.ToInt64();
                        request.Progress?.Report(new(Name, entry, completed, null));
                        return 1;
                    case RarCallbackMessage.LargeDictionary:
                        callbackFailure = new(ExtractionStatus.Unsupported, $"RAR 字典超过引擎内存限制：{p1.ToInt64()} KiB。");
                        return -1;
                    default:
                        return 0;
                }
            }
            catch (OperationCanceledException)
            {
                callbackFailure = new(ExtractionStatus.Cancelled, "已取消");
                return -1;
            }
            catch (Exception e)
            {
                callbackFailure = new(ExtractionStatus.IoError, e.Message);
                return -1;
            }
        };
        try
        {
            request.CancellationToken.ThrowIfCancellationRequested();
            var data = new OpenData
            {
                ArcNameW = request.Archive.EntryPath,
                OpenMode = RarOpenMode.Extract,
                Callback = callback
            };
            handle = _open(ref data);
            if (callbackFailure is not null)
                return callbackFailure;
            if (data.OpenResult is not RarError.Success || handle == nint.Zero)
                return Map(data.OpenResult, (data.Flags & RarArchiveFlags.EncryptedHeaders) is not 0);
            // Fail if a continuation was supplied; partial extraction must never count as success.
            if ((data.Flags & RarArchiveFlags.Volume) is not 0 && (data.Flags & RarArchiveFlags.FirstVolume) is 0)
                return new(ExtractionStatus.MissingVolume, "需要 RAR 第一卷。");
            while (true)
            {
                request.CancellationToken.ThrowIfCancellationRequested();
                *longName = '\0';
                var header = new HeaderData
                {
                    FileNameEx = longName,
                    FileNameExSize = 32768
                };
                var rc = _read(handle, ref header);
                if (callbackFailure is not null)
                    return callbackFailure;
                if (rc is RarError.EndArchive)
                    return ExtractionResult.Success([.. openedParts]);
                if (rc is not RarError.Success)
                    return Map(rc, encrypted || (data.Flags & RarArchiveFlags.EncryptedHeaders) is not 0);
                entry = *longName is not '\0' ? new(longName) : new string(header.FileNameW);
                encrypted = (header.Flags & RarHeaderFlags.Encrypted) is not 0;
                if (header.RedirType is not 0)
                    return new(ExtractionStatus.Unsupported, $"暂不提取链接：{entry}");
                var output = OutputPaths.Get(request.OutputDirectory, entry);
                if ((header.Flags & RarHeaderFlags.Directory) is 0 && File.Exists(output))
                    return new(ExtractionStatus.IoError, $"压缩包包含重名文件：{entry}");
                _ = Directory.CreateDirectory(Path.GetDirectoryName(output)!);
                rc = _process(handle, RarOperation.Extract, null, output);
                if (callbackFailure is not null)
                    return callbackFailure;
                if (rc is not RarError.Success)
                    return Map(rc, encrypted);
            }
        }
        catch (OperationCanceledException)
        {
            return new(ExtractionStatus.Cancelled, "已取消");
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or ArgumentException)
        {
            return new(ExtractionStatus.IoError, e.Message);
        }
        finally
        {
            if (handle != nint.Zero)
                _ = _close(handle);
            GC.KeepAlive(callback);
            NativeMemory.Free(longName);
        }
    }

    private static void WriteString(char* ptr, string value, int capacity)
    {
        if (value.Length >= capacity)
            throw new IOException("原生接口字符串缓冲区不足。");
        value.CopyTo(new(ptr, capacity));
        ptr[value.Length] = '\0';
    }

    private static ExtractionResult Map(RarError error, bool encrypted) => new(error switch
    {
        RarError.MissingPassword or RarError.BadPassword => ExtractionStatus.WrongPassword,
        RarError.BadData when encrypted => ExtractionStatus.PasswordOrCorruptData,
        RarError.BadData or RarError.BadArchive => ExtractionStatus.InvalidArchive,
        RarError.UnknownFormat or RarError.Reference or RarError.LargeDictionary => ExtractionStatus.Unsupported,
        _ => ExtractionStatus.IoError
    }, $"UnRAR 错误 {error}（{(uint) error}）" + (error is RarError.BadData && encrypted ? "：密码不正确或加密数据损坏。" : "。"));

    public void Dispose() => NativeLibrary.Free(_library);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate nint Open(ref OpenData data);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate RarError Close(nint archive);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate RarError ReadHeader(nint archive, ref HeaderData data);

    [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Unicode)]
    private delegate RarError ProcessFile(nint archive, RarOperation operation, string? destination, string? name);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int Callback(RarCallbackMessage message, nint user, nint p1, nint p2);

    private enum RarOpenMode : uint
    {
        List = 0,
        Extract = 1,
        ListIncludingSplit = 2
    }

    private enum RarOperation : int
    {
        Skip = 0,
        Test = 1,
        Extract = 2
    }

    private enum RarCallbackMessage : uint
    {
        ChangeVolume = 0,
        ProcessData = 1,
        NeedPassword = 2,
        ChangeVolumeWide = 3,
        NeedPasswordWide = 4,
        LargeDictionary = 5
    }

    private enum RarVolumeMode : int
    {
        Ask = 0,
        Notify = 1
    }

    private enum RarError : uint
    {
        Success = 0,
        EndArchive = 10,
        NoMemory = 11,
        BadData = 12,
        BadArchive = 13,
        UnknownFormat = 14,
        Open = 15,
        Create = 16,
        Close = 17,
        Read = 18,
        Write = 19,
        SmallBuffer = 20,
        Unknown = 21,
        MissingPassword = 22,
        Reference = 23,
        BadPassword = 24,
        LargeDictionary = 25
    }

    [Flags]
    private enum RarArchiveFlags : uint
    {
        Volume = 1,
        EncryptedHeaders = 0x80,
        FirstVolume = 0x100
    }

    [Flags]
    private enum RarHeaderFlags : uint
    {
        Encrypted = 4,
        Directory = 0x20
    }

    // UnRAR SDK 7.20, #pragma pack(1). The reserved tail preserves the published ABI.
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    private struct OpenData
    {
        public byte* ArcName;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? ArcNameW;
        public RarOpenMode OpenMode;
        public RarError OpenResult;
        public byte* CmtBuf;
        public uint CmtBufSize, CmtSize, CmtState;
        public RarArchiveFlags Flags;
        [MarshalAs(UnmanagedType.FunctionPtr)]
        public Callback? Callback;
        public nint UserData;
        public uint OpFlags;
        public char* CmtBufW;
        public byte* MarkOfTheWeb;
        public fixed uint Reserved[23];
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1, CharSet = CharSet.Unicode)]
    private struct HeaderData
    {
        public fixed byte ArcName[1024];
        public fixed char ArcNameW[1024];
        public fixed byte FileName[1024];
        public fixed char FileNameW[1024];
        public RarHeaderFlags Flags;
        public uint PackSize, PackSizeHigh, UnpSize, UnpSizeHigh, HostOS, FileCRC, FileTime, UnpVer, Method, FileAttr;
        public byte* CmtBuf;
        public uint CmtBufSize, CmtSize, CmtState, DictSize, HashType;
        public fixed byte Hash[32];
        public uint RedirType;
        public char* RedirName;
        public uint RedirNameSize, DirTarget, MtimeLow, MtimeHigh, CtimeLow, CtimeHigh, AtimeLow, AtimeHigh;
        public char* ArcNameEx;
        public uint ArcNameExSize;
        public char* FileNameEx;
        public uint FileNameExSize;
        public fixed uint Reserved[982];
    }
}

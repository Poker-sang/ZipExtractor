using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace ZipExtractor.Extraction.Native;

// ABI declarations from the public 7-Zip IArchive.h / IStream.h / IPassword.h interfaces.
// Kept internal to the backend; no COM registration or external managed wrapper is required.
[GeneratedComInterface, Guid("23170F69-40C1-278A-0000-000600600000")]
internal unsafe partial interface IInArchive
{
    [PreserveSig]
    int Open(IInStream stream, ulong* maxCheckStartPosition, IArchiveOpenCallback callback);

    [PreserveSig]
    int Close();

    [PreserveSig]
    int GetNumberOfItems(out uint count);

    [PreserveSig]
    int GetProperty(uint index, ArchiveProperty propId, ref PropVariant value);

    [PreserveSig]
    int Extract(uint* indices, uint count, ArchiveExtractMode testMode, IArchiveExtractCallback callback);

    [PreserveSig]
    int GetArchiveProperty(ArchiveProperty propId, ref PropVariant value);

    [PreserveSig]
    int GetNumberOfProperties(out uint count);

    [PreserveSig]
    int GetPropertyInfo(uint index, char** name, ArchiveProperty* propId, PropVariantType* varType);

    [PreserveSig]
    int GetNumberOfArchiveProperties(out uint count);

    [PreserveSig]
    int GetArchivePropertyInfo(uint index, char** name, ArchiveProperty* propId, PropVariantType* varType);
}

[GeneratedComInterface, Guid("23170F69-40C1-278A-0000-000300030000")]
public unsafe partial interface IInStream
{
    [PreserveSig]
    int Read(byte* data, uint size, uint* processed);

    [PreserveSig]
    int Seek(long offset, SeekOrigin origin, ulong* position);
}

[GeneratedComInterface, Guid("23170F69-40C1-278A-0000-000300020000")]
public unsafe partial interface ISequentialOutStream
{
    [PreserveSig]
    int Write(byte* data, uint size, uint* processed);
}

[GeneratedComInterface, Guid("23170F69-40C1-278A-0000-000600100000")]
public unsafe partial interface IArchiveOpenCallback
{
    [PreserveSig]
    int SetTotal(ulong* files, ulong* bytes);

    [PreserveSig]
    int SetCompleted(ulong* files, ulong* bytes);
}

[GeneratedComInterface, Guid("23170F69-40C1-278A-0000-000600200000")]
public unsafe partial interface IArchiveExtractCallback
{
    [PreserveSig]
    int SetTotal(ulong total);

    [PreserveSig]
    int SetCompleted(ulong* complete);

    [PreserveSig]
    int GetStream(uint index, out ISequentialOutStream? stream, ArchiveAskMode mode);

    [PreserveSig]
    int PrepareOperation(ArchiveAskMode mode);

    [PreserveSig]
    int SetOperationResult(ArchiveOperationResult result);
}

[GeneratedComInterface, Guid("23170F69-40C1-278A-0000-000600300000")]
public partial interface IArchiveOpenVolumeCallback
{
    [PreserveSig]
    int GetProperty(ArchiveProperty propId, ref PropVariant value);

    [PreserveSig]
    int GetStream([MarshalAs(UnmanagedType.LPWStr)] string name, out IInStream? stream);
}

[GeneratedComInterface, Guid("23170F69-40C1-278A-0000-000500100000")]
public partial interface ICryptoGetTextPassword
{
    [PreserveSig]
    int CryptoGetTextPassword([MarshalAs(UnmanagedType.BStr)] out string password);
}

[StructLayout(LayoutKind.Explicit, Size = 24)]
public partial struct PropVariant : IDisposable
{
    [FieldOffset(0)]
    public PropVariantType Type;
    [FieldOffset(8)]
    public ulong Number;
    [FieldOffset(8)]
    public nint Pointer;

    public readonly string? Text => Type is PropVariantType.BStr ? Marshal.PtrToStringBSTR(Pointer) : null;

    public readonly bool Boolean => Type is PropVariantType.Boolean && (short) Number is not 0;

    public readonly ulong? Unsigned => Type is PropVariantType.UInt32 or PropVariantType.UInt64 ? Number : null;

    public void Dispose()
    {
        _ = PropVariantClear(ref this);
    }

    [LibraryImport("ole32.dll")]
    private static partial int PropVariantClear(ref PropVariant value);
}

// Values are defined by the 7-Zip SDK; preserve their native widths and vtable order.
public enum ArchiveProperty : uint
{
    Path = 3,
    Name = 4,
    IsDirectory = 6,
    Size = 7,
    ModifiedTime = 12,
    Encrypted = 15,
    Offset = 36,
    PhysicalSize = 44,
    SymbolicLink = 54,
    ErrorFlags = 71,
    WarningFlags = 72,
    HardLink = 90
}

public enum ArchiveExtractMode : int
{
    Extract = 0,
    Test = 1
}

public enum ArchiveAskMode : int
{
    Extract = 0,
    Test = 1,
    Skip = 2,
    ReadExternal = 3
}

public enum ArchiveOperationResult : int
{
    Ok = 0,
    UnsupportedMethod = 1,
    DataError = 2,
    CrcError = 3,
    Unavailable = 4,
    UnexpectedEnd = 5,
    DataAfterEnd = 6,
    IsNotArchive = 7,
    HeadersError = 8,
    WrongPassword = 9
}

// VARTYPE occupies 16 bits even though System.Runtime.InteropServices.VarEnum uses Int32.
public enum PropVariantType : ushort
{
    Empty = 0,
    BStr = 8,
    Boolean = 11,
    UInt32 = 19,
    UInt64 = 21,
    FileTime = 64
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using System.Threading;

namespace ZipExtractor.Extraction.Native;

/// <summary>Seekable concatenation for byte-split 7z/ZIP; opens only one volume at a time.</summary>
public sealed class ConcatenatedStream : Stream
{
    private readonly string[] _paths;
    private readonly long[] _starts;
    private FileStream? _current;
    private int _currentIndex = -1;
    private long _position;

    public ConcatenatedStream(IReadOnlyList<string> parts)
    {
        _paths = [.. parts];
        _starts = new long[_paths.Length + 1];
        for (var i = 0; i < _paths.Length; i++)
            _starts[i + 1] = checked(_starts[i] + new FileInfo(_paths[i]).Length);
    }

    public override bool CanRead => true;

    public override bool CanSeek => true;

    public override bool CanWrite => false;

    public override long Length => _starts[^1];

    public override long Position
    {
        get => _position;
        set => Seek(value, SeekOrigin.Begin);
    }

    public override int Read(byte[] buffer, int offset, int count) => Read(buffer.AsSpan(offset, count));

    public override int Read(Span<byte> buffer)
    {
        if (buffer.IsEmpty || _position >= Length)
            return 0;
        var index = Array.BinarySearch(_starts, _position);
        if (index < 0)
            index = ~index - 1;
        while (index < _paths.Length && _starts[index + 1] == _position)
            index++;
        if (index >= _paths.Length)
            return 0;
        if (_currentIndex != index)
        {
            _current?.Dispose();
            _current = File.OpenRead(_paths[index]);
            _currentIndex = index;
        }

        _current!.Position = _position - _starts[index];
        var n = _current.Read(buffer[..(int) Math.Min(buffer.Length, _starts[index + 1] - _position)]);
        if (n is 0)
            throw new EndOfStreamException("解压过程中分卷长度发生变化。");
        _position += n;
        return n;
    }

    public override long Seek(long offset, SeekOrigin origin)
    {
        var next = checked((origin switch
        {
            SeekOrigin.Begin => 0,
            SeekOrigin.Current => _position,
            SeekOrigin.End => Length,
            _ => throw new ArgumentOutOfRangeException(nameof(origin))
        }) + offset);
        return next < 0 ? throw new IOException("负偏移") : (_position = next);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
            _current?.Dispose();
        base.Dispose(disposing);
    }

    public override void Flush()
    {
    }

    public override void SetLength(long value) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
}

[GeneratedComClass]
public sealed unsafe partial class NativeInputStream(Stream source, CancellationToken token) : IInStream, IDisposable
{
    public long Length => source.Length;

    public Exception? Error { get; private set; }

    public int Read(byte* data, uint size, uint* processed)
    {
        if (processed is not null)
            *processed = 0;
        try
        {
            token.ThrowIfCancellationRequested();
            var n = source.Read(new(data, checked((int) size)));
            if (processed is not null)
                *processed = (uint) n;
            return 0;
        }
        catch (Exception e)
        {
            Error = e;
            return Marshal.GetHRForException(e);
        }
    }

    public int Seek(long offset, SeekOrigin origin, ulong* position)
    {
        try
        {
            token.ThrowIfCancellationRequested();
            var result = source.Seek(offset, origin);
            if (position is not null)
                *position = (ulong) result;
            return 0;
        }
        catch (Exception e)
        {
            Error = e;
            return Marshal.GetHRForException(e);
        }
    }

    public void Dispose() => source.Dispose();
}

[GeneratedComClass]
public sealed unsafe partial class NativeOutputStream(Stream target, CancellationToken token) : ISequentialOutStream, IDisposable
{
    public Exception? Error { get; private set; }

    public int Write(byte* data, uint size, uint* processed)
    {
        if (processed is not null)
            *processed = 0;
        try
        {
            token.ThrowIfCancellationRequested();
            target.Write(new(data, checked((int) size)));
            if (processed is not null)
                *processed = size;
            return 0;
        }
        catch (Exception e)
        {
            Error = e;
            return Marshal.GetHRForException(e);
        }
    }

    public void Dispose() => target.Dispose();
}

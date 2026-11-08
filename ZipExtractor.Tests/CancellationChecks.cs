using ZipExtractor.Extraction;

internal static class CancellationChecks
{
    public static void Verify(string directory, int mode)
    {
        var input = Directory.CreateDirectory(Path.Combine(directory, "input")).FullName;
        var output = Path.Combine(directory, "output");
        var source = Path.Combine(input, "source.7z");
        File.WriteAllBytes(source, [0x37, 0x7a, 0xbc, 0xaf, 0x27, 0x1c]);
        using var cancellation = new CancellationTokenSource();
        using var engine = new CancellingEngine(cancellation, mode);
        try
        {
            _ = new ExtractionPipeline([engine]).Run(new(input, output), [], cancellationToken: cancellation.Token);
            throw new InvalidOperationException("Cancellation was not propagated.");
        }
        catch (OperationCanceledException)
        {
        }
        if (!File.Exists(source))
            throw new InvalidOperationException("Cancellation removed the source archive.");
        var jobs = Directory.GetDirectories(output, ".extract-*");
        if (mode < 2)
        {
            if (jobs.Length is not 0)
                throw new InvalidOperationException("Incomplete extraction was not cleaned up.");
            return;
        }
        var job = jobs.Single();
        var contents = Directory.GetFiles(job, "*.txt", SearchOption.AllDirectories);
        var expectedCount = mode is 5 ? 2 : 1;
        if (contents.Length != expectedCount || contents.Any(p => File.ReadAllText(p) is not "preserve me"))
            throw new InvalidOperationException("Completed content was lost or partial content was retained.");
        if (mode is 4 && Directory.GetFiles(job, "partial.txt", SearchOption.AllDirectories).Length is not 0)
            throw new InvalidOperationException("Partially extracted nested archive was retained.");
    }

    private sealed class CancellingEngine(CancellationTokenSource cancellation, int mode) : IArchiveExtractor, IArchiveProbe
    {
        public string Name => "Cancellation test";

        public bool Supports(ArchiveFormat format) => true;

        public ExtractionResult Extract(ExtractionRequest request)
        {
            if (mode >= 4 && Path.GetFileName(request.Archive.EntryPath) is "source.7z")
            {
                File.WriteAllText(Path.Combine(request.OutputDirectory, "outer.txt"), "preserve me");
                File.WriteAllBytes(Path.Combine(request.OutputDirectory, "child.7z"), [0x37, 0x7a, 0xbc, 0xaf, 0x27, 0x1c]);
                return ExtractionResult.Success();
            }
            File.WriteAllText(Path.Combine(request.OutputDirectory, "partial.txt"), "preserve me");
            switch (mode)
            {
                case 0:
                    return new(ExtractionStatus.Cancelled, "cancelled by backend");
                case 1:
                    throw new OperationCanceledException();
                case 2:
                case 5:
                    cancellation.Cancel();
                    return ExtractionResult.Success();
                case 4:
                    cancellation.Cancel();
                    return new(ExtractionStatus.Cancelled, "nested extraction cancelled");
                default:
                    return ExtractionResult.Success();
            }
        }

        public ArchiveFormat Probe(string path, CancellationToken token = default)
        {
            if (mode >= 4)
                return ArchiveFormat.Unknown;
            cancellation.Cancel();
            token.ThrowIfCancellationRequested();
            return ArchiveFormat.Unknown;
        }

        public void Dispose()
        {
        }
    }
}

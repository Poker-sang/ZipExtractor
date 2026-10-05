using System;
using ZipExtractor.Extraction;

namespace ZipExtractor.Logging;

/// <summary>Native callbacks must update progress synchronously, without posting to a UI context.</summary>
internal sealed class CallbackProgress(Action<ExtractionProgress> callback) : IProgress<ExtractionProgress>
{
    public void Report(ExtractionProgress value) => callback(value);
}

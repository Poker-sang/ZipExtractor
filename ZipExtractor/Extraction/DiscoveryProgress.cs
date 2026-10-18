namespace ZipExtractor.Extraction;

public sealed record DiscoveryProgress(string Stage, string Path, int FilesScanned, int ArchivesFound);

namespace DocToolkit.Tests;

/// <summary>
/// A uniquely named temp file that deletes itself. Every file-path test needs one, and a test that
/// leaks files into the runner's temp directory is a test that eventually fills a CI disk.
/// </summary>
internal sealed class TempFile : IDisposable
{
    public string Path { get; } =
        System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"doctoolkit-{Guid.NewGuid():N}.tmp");

    public void Dispose()
    {
        try { if (File.Exists(Path)) File.Delete(Path); } catch { /* best effort */ }
    }
}

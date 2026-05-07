using System.IO.Packaging;
using System.Runtime.InteropServices.JavaScript;
using DocxSax;
using DocumentFormat.OpenXml.Packaging;

namespace DocxSax.Browser;

internal static class Program
{
    public static void Main()
    {
    }
}

public static partial class BrowserBridge
{
    [JSExport]
    public static string ParseBytesJsonBatchFrames(byte[] bytes, int batchSize)
    {
        ArgumentNullException.ThrowIfNull(bytes);

        try
        {
            using var stream = new MemoryStream(bytes, writable: false);
            var reader = new DocxSaxReader();
            return DocxEventJson.ToJsonBatchFrames(reader.Read(stream), batchSize);
        }
        catch (Exception exception) when (exception is InvalidDataException or FileFormatException or OpenXmlPackageException)
        {
            throw new InvalidOperationException($"docx-sax browser bridge failed to parse DOCX bytes: {exception.Message}", exception);
        }
    }
}

using System.IO.Packaging;
using System.Runtime.InteropServices.JavaScript;
using System.Text;
using DocxSax;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxSax.Browser;

internal static class Program
{
    public static void Main()
    {
    }
}

public static partial class BrowserBridge
{
    private static readonly object ParseSessionsLock = new();
    private static readonly Dictionary<int, ParseSession> ParseSessions = new();
    private static int nextParseSessionId;

    [JSExport]
    public static int Warmup()
    {
        using var stream = CreateWarmupDocx();
        var reader = new DocxSaxReader();
        var eventCount = 0;

        foreach (var docxEvent in reader.Read(stream))
        {
            _ = DocxEventJson.ToJson(docxEvent);
            eventCount++;
        }

        _ = DocxEventJson.ToJsonBatchFrames(
            new DocxEvent[]
            {
                new PackageEvent(DocxEventKind.PackageStart, 0),
                new DiagnosticEvent(1, "warmup"),
                new PackageEvent(DocxEventKind.PackageEnd, 2),
            },
            batchSize: 2);

        return eventCount;
    }

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

    [JSExport]
    public static int BeginParseBytesJsonBatches(byte[] bytes, int batchSize)
    {
        ArgumentNullException.ThrowIfNull(bytes);

        var stream = new MemoryStream(bytes, writable: false);
        var reader = new DocxSaxReader();
        var session = new ParseSession(stream, reader.Read(stream).GetEnumerator(), NormalizeBatchSize(batchSize));

        lock (ParseSessionsLock)
        {
            var id = ++nextParseSessionId;
            ParseSessions.Add(id, session);
            return id;
        }
    }

    [JSExport]
    public static string? ReadNextJsonBatch(int parseSessionId)
    {
        var session = GetParseSession(parseSessionId);

        try
        {
            var builder = new StringBuilder();
            builder.Append('[');
            var count = 0;
            while (count < session.BatchSize && session.Enumerator.MoveNext())
            {
                if (count > 0)
                {
                    builder.Append(',');
                }

                builder.Append(DocxEventJson.ToJson(session.Enumerator.Current));
                count++;
            }

            if (count == 0)
            {
                EndParseBytesJsonBatches(parseSessionId);
                return null;
            }

            builder.Append(']');
            return builder.ToString();
        }
        catch (Exception exception) when (exception is InvalidDataException or FileFormatException or OpenXmlPackageException)
        {
            EndParseBytesJsonBatches(parseSessionId);
            throw new InvalidOperationException($"docx-sax browser bridge failed to parse DOCX bytes: {exception.Message}", exception);
        }
    }

    [JSExport]
    public static void EndParseBytesJsonBatches(int parseSessionId)
    {
        ParseSession? session = null;
        lock (ParseSessionsLock)
        {
            if (ParseSessions.Remove(parseSessionId, out var removed))
            {
                session = removed;
            }
        }

        session?.Dispose();
    }

    private static ParseSession GetParseSession(int parseSessionId)
    {
        lock (ParseSessionsLock)
        {
            if (ParseSessions.TryGetValue(parseSessionId, out var session))
            {
                return session;
            }
        }

        throw new InvalidOperationException($"Unknown docx-sax browser parse session: {parseSessionId}");
    }

    private static int NormalizeBatchSize(int batchSize) => batchSize > 0 ? batchSize : 128;

    private static MemoryStream CreateWarmupDocx()
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: true))
        {
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(
                new Body(
                    new Paragraph(
                        new Run(
                            new Text("docx-sax warmup")
                            {
                                Space = SpaceProcessingModeValues.Preserve,
                            }))));
            mainPart.Document.Save();
        }

        stream.Position = 0;
        return stream;
    }

    private sealed class ParseSession(MemoryStream stream, IEnumerator<DocxEvent> enumerator, int batchSize) : IDisposable
    {
        public IEnumerator<DocxEvent> Enumerator { get; } = enumerator;
        public int BatchSize { get; } = batchSize;

        public void Dispose()
        {
            Enumerator.Dispose();
            stream.Dispose();
        }
    }
}

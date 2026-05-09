using System.IO.Packaging;
using System.Runtime.InteropServices.JavaScript;
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
            _ = ToRow(docxEvent);
            eventCount++;
        }

        _ = new object?[]
        {
            ToRow(new PackageEvent(DocxEventKind.PackageStart, 0)),
            ToRow(new DiagnosticEvent(1, "warmup")),
            ToRow(new PackageEvent(DocxEventKind.PackageEnd, 2)),
        };

        return eventCount;
    }

    [JSExport]
    public static int BeginParseBytesBatches(byte[] bytes, int batchSize)
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
    [return: JSMarshalAs<JSType.Array<JSType.Any>>]
    public static object?[]? ReadNextBatch(int parseSessionId)
    {
        var session = GetParseSession(parseSessionId);

        try
        {
            var batch = new List<object?>(session.BatchSize);
            while (batch.Count < session.BatchSize && session.Enumerator.MoveNext())
            {
                batch.Add(ToRow(session.Enumerator.Current));
            }

            if (batch.Count == 0)
            {
                EndParseBytesBatches(parseSessionId);
                return null;
            }

            return batch.ToArray();
        }
        catch (Exception exception) when (exception is InvalidDataException or FileFormatException or OpenXmlPackageException)
        {
            EndParseBytesBatches(parseSessionId);
            throw new InvalidOperationException($"docx-sax browser bridge failed to parse DOCX bytes: {exception.Message}", exception);
        }
    }

    [JSExport]
    public static void EndParseBytesBatches(int parseSessionId)
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

    private static object?[] ToRow(DocxEvent docxEvent) => docxEvent switch
    {
        PackageEvent package => Row(
            docxEvent.Kind,
            docxEvent.Ordinal,
            phase: package.IsStart ? "start" : "end"),
        PartEvent part => Row(
            docxEvent.Kind,
            docxEvent.Ordinal,
            phase: part.IsStart ? "start" : "end",
            uri: part.Uri,
            contentType: part.ContentType,
            relationshipType: part.RelationshipType),
        RelationshipEvent relationship => Row(
            docxEvent.Kind,
            docxEvent.Ordinal,
            relationshipType: relationship.RelationshipType,
            sourceUri: relationship.SourceUri,
            id: relationship.Id,
            targetUri: relationship.TargetUri,
            isExternal: relationship.IsExternal),
        ElementStartEvent element => Row(
            docxEvent.Kind,
            docxEvent.Ordinal,
            partUri: element.PartUri,
            name: element.Name,
            localName: element.LocalName,
            prefix: element.Prefix,
            namespaceUri: element.NamespaceUri,
            depth: element.Depth,
            path: element.Path,
            isEmptyElement: element.IsEmptyElement,
            attributes: element.Attributes.Select(attribute => (object?)ToAttributeRow(attribute)).ToArray()),
        ElementEndEvent element => Row(
            docxEvent.Kind,
            docxEvent.Ordinal,
            partUri: element.PartUri,
            name: element.Name,
            localName: element.LocalName,
            prefix: element.Prefix,
            namespaceUri: element.NamespaceUri,
            depth: element.Depth,
            path: element.Path),
        TextEvent text => Row(
            docxEvent.Kind,
            docxEvent.Ordinal,
            partUri: text.PartUri,
            text: text.Text,
            depth: text.Depth,
            path: text.Path,
            isWhitespace: text.IsWhitespace),
        DiagnosticEvent diagnostic => Row(
            docxEvent.Kind,
            docxEvent.Ordinal,
            partUri: diagnostic.PartUri,
            message: diagnostic.Message),
        _ => throw new InvalidOperationException($"Unsupported event type: {docxEvent.GetType().Name}"),
    };

    private static object?[] Row(
        DocxEventKind kind,
        long ordinal,
        string? phase = null,
        string? uri = null,
        string? contentType = null,
        string? relationshipType = null,
        string? sourceUri = null,
        string? id = null,
        string? targetUri = null,
        bool isExternal = false,
        string? partUri = null,
        string? name = null,
        string? localName = null,
        string? prefix = null,
        string? namespaceUri = null,
        int depth = 0,
        string? path = null,
        bool isEmptyElement = false,
        object?[]? attributes = null,
        string? text = null,
        bool isWhitespace = false,
        string? message = null) =>
        [
            (int)kind,
            (double)ordinal,
            phase,
            uri,
            contentType,
            relationshipType,
            sourceUri,
            id,
            targetUri,
            isExternal,
            partUri,
            name,
            localName,
            prefix,
            namespaceUri,
            depth,
            path,
            isEmptyElement,
            attributes ?? [],
            text,
            isWhitespace,
            message,
        ];

    private static object?[] ToAttributeRow(DocxAttribute attribute) =>
        [
            attribute.Name,
            attribute.LocalName,
            attribute.Prefix,
            attribute.NamespaceUri,
            attribute.Value,
        ];

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

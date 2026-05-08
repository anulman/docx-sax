using System.Xml;
using DocumentFormat.OpenXml.Packaging;

namespace DocxSax;

/// <summary>
/// Reads DOCX/OpenXML packages as a typed, low-level event stream.
/// </summary>
public sealed class DocxSaxReader
{
    /// <summary>
    /// Reads a DOCX package from <paramref name="stream" /> and emits typed OpenXML events.
    /// </summary>
    /// <param name="stream">A readable, seekable DOCX/OpenXML package stream.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="stream" /> is <see langword="null" />.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="stream" /> is not readable or seekable.</exception>
    public IEnumerable<DocxEvent> Read(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        if (!stream.CanRead)
        {
            throw new ArgumentException("The DOCX stream must be readable.", nameof(stream));
        }

        if (!stream.CanSeek)
        {
            throw new ArgumentException("The DOCX stream must be seekable.", nameof(stream));
        }

        return ReadCore(stream);
    }

    private static IEnumerable<DocxEvent> ReadCore(Stream stream)
    {
        var ordinals = new OrdinalCounter();
        yield return new PackageEvent(DocxEventKind.PackageStart, ordinals.Next());

        using var document = WordprocessingDocument.Open(stream, false);
        var mainDocumentPart = document.MainDocumentPart;
        if (mainDocumentPart is null)
        {
            yield return new DiagnosticEvent(ordinals.Next(), "Package does not contain a main document part.");
        }
        else
        {
            foreach (var relationship in ReadPackageRelationships(document, mainDocumentPart, ordinals))
            {
                yield return relationship;
            }

            var visitedParts = new HashSet<string>(StringComparer.Ordinal);
            foreach (var docxEvent in ReadPart(mainDocumentPart, ordinals, visitedParts))
            {
                yield return docxEvent;
            }
        }

        yield return new PackageEvent(DocxEventKind.PackageEnd, ordinals.Next());
    }

    private static IEnumerable<DocxEvent> ReadPackageRelationships(
        WordprocessingDocument document,
        MainDocumentPart mainDocumentPart,
        OrdinalCounter ordinals)
    {
        yield return new RelationshipEvent(
            ordinals.Next(),
            SourceUri: "/",
            Id: document.GetIdOfPart(mainDocumentPart),
            RelationshipType: mainDocumentPart.RelationshipType,
            TargetUri: mainDocumentPart.Uri.ToString(),
            IsExternal: false);
    }

    private static IEnumerable<DocxEvent> ReadPart(OpenXmlPart part, OrdinalCounter ordinals, ISet<string> visitedParts)
    {
        var partUri = part.Uri.ToString();
        if (!visitedParts.Add(partUri))
        {
            yield break;
        }

        yield return new PartEvent(DocxEventKind.PartStart, ordinals.Next(), partUri, part.ContentType, part.RelationshipType);

        var childParts = part.Parts.OrderBy(pair => pair.RelationshipId, StringComparer.Ordinal).ToArray();
        foreach (var relationship in ReadPartRelationships(part, partUri, childParts, ordinals))
        {
            yield return relationship;
        }

        if (IsXmlContentType(part.ContentType))
        {
            foreach (var xmlEvent in ReadXml(part, partUri, ordinals))
            {
                yield return xmlEvent;
            }
        }

        foreach (var child in childParts)
        {
            foreach (var childEvent in ReadPart(child.OpenXmlPart, ordinals, visitedParts))
            {
                yield return childEvent;
            }
        }

        yield return new PartEvent(DocxEventKind.PartEnd, ordinals.Next(), partUri, part.ContentType, part.RelationshipType);
    }

    private static bool IsXmlContentType(string contentType) =>
        contentType.EndsWith("+xml", StringComparison.OrdinalIgnoreCase) ||
        contentType.EndsWith("/xml", StringComparison.OrdinalIgnoreCase) ||
        contentType.Equals("application/xml", StringComparison.OrdinalIgnoreCase) ||
        contentType.Equals("text/xml", StringComparison.OrdinalIgnoreCase);

    private static IEnumerable<RelationshipEvent> ReadPartRelationships(
        OpenXmlPart part,
        string partUri,
        IReadOnlyList<IdPartPair> childParts,
        OrdinalCounter ordinals)
    {
        foreach (var child in childParts)
        {
            yield return new RelationshipEvent(
                ordinals.Next(),
                SourceUri: partUri,
                Id: child.RelationshipId,
                RelationshipType: child.OpenXmlPart.RelationshipType,
                TargetUri: child.OpenXmlPart.Uri.ToString(),
                IsExternal: false);
        }

        foreach (var hyperlink in part.HyperlinkRelationships.OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
        {
            yield return new RelationshipEvent(
                ordinals.Next(),
                SourceUri: partUri,
                Id: hyperlink.Id,
                RelationshipType: hyperlink.RelationshipType,
                TargetUri: hyperlink.Uri.ToString(),
                IsExternal: hyperlink.IsExternal);
        }

        foreach (var external in part.ExternalRelationships.OrderBy(relationship => relationship.Id, StringComparer.Ordinal))
        {
            yield return new RelationshipEvent(
                ordinals.Next(),
                SourceUri: partUri,
                Id: external.Id,
                RelationshipType: external.RelationshipType,
                TargetUri: external.Uri.ToString(),
                IsExternal: true);
        }
    }

    private static IEnumerable<DocxEvent> ReadXml(OpenXmlPart part, string partUri, OrdinalCounter ordinals)
    {
        using var partStream = part.GetStream(FileMode.Open, FileAccess.Read);
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            IgnoreComments = false,
            IgnoreProcessingInstructions = false,
            IgnoreWhitespace = false,
        };

        using var reader = XmlReader.Create(partStream, settings);
        var path = new XmlPathStack();

        while (reader.Read())
        {
            switch (reader.NodeType)
            {
                case XmlNodeType.Element:
                    {
                        var elementName = reader.Name;
                        path.Push(elementName);
                        var currentPath = path.Current;
                        var attributes = ReadAttributes(reader);
                        var isEmpty = reader.IsEmptyElement;

                        yield return new ElementStartEvent(
                            ordinals.Next(),
                            partUri,
                            elementName,
                            reader.LocalName,
                            reader.Prefix,
                            reader.NamespaceURI,
                            attributes,
                            reader.Depth,
                            currentPath,
                            isEmpty);

                        if (isEmpty)
                        {
                            yield return new ElementEndEvent(
                                ordinals.Next(),
                                partUri,
                                elementName,
                                reader.LocalName,
                                reader.Prefix,
                                reader.NamespaceURI,
                                reader.Depth,
                                currentPath);
                            path.Pop();
                        }

                        break;
                    }

                case XmlNodeType.EndElement:
                    {
                        var currentPath = path.Current;
                        yield return new ElementEndEvent(
                            ordinals.Next(),
                            partUri,
                            reader.Name,
                            reader.LocalName,
                            reader.Prefix,
                            reader.NamespaceURI,
                            reader.Depth,
                            currentPath);

                        path.PopIfAny();

                        break;
                    }

                case XmlNodeType.Text:
                case XmlNodeType.CDATA:
                case XmlNodeType.SignificantWhitespace:
                case XmlNodeType.Whitespace:
                    yield return new TextEvent(
                        ordinals.Next(),
                        partUri,
                        reader.Value,
                        reader.Depth,
                        path.Current,
                        reader.NodeType is XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace);
                    break;
            }
        }
    }

    private static IReadOnlyList<DocxAttribute> ReadAttributes(XmlReader reader)
    {
        if (!reader.HasAttributes)
        {
            return Array.Empty<DocxAttribute>();
        }

        var attributes = new List<DocxAttribute>(reader.AttributeCount);
        while (reader.MoveToNextAttribute())
        {
            attributes.Add(new DocxAttribute(
                reader.Name,
                reader.LocalName,
                reader.Prefix,
                reader.NamespaceURI,
                reader.Value));
        }

        reader.MoveToElement();
        return attributes;
    }

    private sealed class XmlPathStack
    {
        private readonly List<string> names = [];
        private string current = "/";

        public string Current => current;

        public void Push(string name)
        {
            names.Add(name);
            current = current.Length == 1 ? string.Concat(current, name) : string.Concat(current, '/', name);
        }

        public void PopIfAny()
        {
            if (names.Count > 0)
            {
                Pop();
            }
        }

        public void Pop()
        {
            if (names.Count == 0)
            {
                return;
            }

            names.RemoveAt(names.Count - 1);
            var slashIndex = current.LastIndexOf('/');
            current = slashIndex <= 0 ? "/" : current[..slashIndex];
        }
    }

    private sealed class OrdinalCounter
    {
        private long value;

        public long Next() => value++;
    }
}

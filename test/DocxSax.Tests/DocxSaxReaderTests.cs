using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocxSax.Tests;

public sealed class DocxSaxReaderTests
{
    [Fact]
    public void Read_EmitsTypedPackagePartRelationshipAndXmlEvents()
    {
        using var stream = CreateSimpleDocx("Hello DOCX SAX");

        var events = new DocxSaxReader().Read(stream).ToArray();

        Assert.NotEmpty(events);
        Assert.Equal(DocxEventKind.PackageStart, events.First().Kind);
        Assert.Equal(DocxEventKind.PackageEnd, events.Last().Kind);
        Assert.Equal(Enumerable.Range(0, events.Length).Select(i => (long)i), events.Select(e => e.Ordinal));

        var mainPart = Assert.Single(events.OfType<PartEvent>(), e => e.Kind == DocxEventKind.PartStart);
        Assert.Equal("/word/document.xml", mainPart.Uri);
        Assert.Equal("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", mainPart.ContentType);

        var packageRelationship = Assert.Single(events.OfType<RelationshipEvent>(), e => e.SourceUri == "/");
        Assert.Equal("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", packageRelationship.RelationshipType);
        Assert.Equal("/word/document.xml", packageRelationship.TargetUri);
        Assert.False(packageRelationship.IsExternal);

        Assert.Contains(events.OfType<ElementStartEvent>(), e =>
            e.PartUri == "/word/document.xml" &&
            e.Name == "w:document" &&
            e.LocalName == "document" &&
            e.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main" &&
            e.Depth == 0 &&
            e.Path == "/w:document");

        Assert.Contains(events.OfType<ElementStartEvent>(), e =>
            e.Name == "w:t" &&
            e.Path == "/w:document/w:body/w:p/w:r/w:t");

        var text = Assert.Single(events.OfType<TextEvent>(), e => e.Text == "Hello DOCX SAX");
        Assert.Equal("/word/document.xml", text.PartUri);
        Assert.Equal("/w:document/w:body/w:p/w:r/w:t", text.Path);
        Assert.False(text.IsWhitespace);
    }

    [Fact]
    public void Read_PreservesAttributesAndStaysLowLevel()
    {
        using var stream = CreateSimpleDocx("Preserve me", includeSpacePreserve: true);

        var events = new DocxSaxReader().Read(stream).ToArray();

        var textStart = Assert.Single(events.OfType<ElementStartEvent>(), e => e.Name == "w:t");
        var space = Assert.Single(textStart.Attributes, a => a.Name == "xml:space");
        Assert.Equal("space", space.LocalName);
        Assert.Equal("xml", space.Prefix);
        Assert.Equal("http://www.w3.org/XML/1998/namespace", space.NamespaceUri);
        Assert.Equal("preserve", space.Value);

        Assert.DoesNotContain(Enum.GetNames<DocxEventKind>(), name => name.Contains("Paragraph", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(events.OfType<ElementStartEvent>(), e => e.LocalName == "p");
        Assert.DoesNotContain(events, e => e.GetType().Name.Contains("Paragraph", StringComparison.OrdinalIgnoreCase));
    }


    [Fact]
    public void Read_ThrowsArgumentNullExceptionForNullStream()
    {
        var exception = Assert.Throws<ArgumentNullException>(() => new DocxSaxReader().Read(null!));

        Assert.Equal("stream", exception.ParamName);
    }

    [Fact]
    public void Read_ThrowsArgumentExceptionForUnreadableStream()
    {
        using var stream = new UnreadableSeekableStream();

        var exception = Assert.Throws<ArgumentException>(() => new DocxSaxReader().Read(stream));

        Assert.Equal("stream", exception.ParamName);
        Assert.Contains("readable", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Read_ThrowsArgumentExceptionForNonSeekableStream()
    {
        using var inner = CreateSimpleDocx("Non-seekable");
        using var stream = new NonSeekableReadStream(inner);

        var exception = Assert.Throws<ArgumentException>(() => new DocxSaxReader().Read(stream));

        Assert.Equal("stream", exception.ParamName);
        Assert.Contains("seekable", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static MemoryStream CreateSimpleDocx(string text, bool includeSpacePreserve = false)
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, autoSave: true))
        {
            var mainPart = document.AddMainDocumentPart();
            var textElement = new Text(text);
            if (includeSpacePreserve)
            {
                textElement.Space = SpaceProcessingModeValues.Preserve;
            }

            mainPart.Document = new Document(new Body(new Paragraph(new Run(textElement))));
            mainPart.Document.Save();
        }

        stream.Position = 0;
        return stream;
    }

    private sealed class UnreadableSeekableStream : Stream
    {
        public override bool CanRead => false;

        public override bool CanSeek => true;

        public override bool CanWrite => false;

        public override long Length => 0;

        public override long Position { get; set; }

        public override void Flush()
        {
        }

        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        public override long Seek(long offset, SeekOrigin origin) => 0;

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    private sealed class NonSeekableReadStream(Stream inner) : Stream
    {
        public override bool CanRead => inner.CanRead;

        public override bool CanSeek => false;

        public override bool CanWrite => false;

        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() => inner.Flush();

        public override int Read(byte[] buffer, int offset, int count) => inner.Read(buffer, offset, count);

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}

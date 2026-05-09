namespace DocxSax;

/// <summary>
/// Base type for every typed event emitted by <see cref="DocxSaxReader" />.
/// </summary>
/// <param name="Kind">The concrete event kind.</param>
/// <param name="Ordinal">Zero-based ordinal in the emitted stream.</param>
public abstract record DocxEvent(DocxEventKind Kind, long Ordinal);

/// <summary>Marks the start or end of the DOCX package stream.</summary>
public sealed record PackageEvent(DocxEventKind Kind, long Ordinal) : DocxEvent(Kind, Ordinal)
{
    public bool IsStart => Kind == DocxEventKind.PackageStart;
}

/// <summary>Marks the start or end of an OpenXML package part.</summary>
public sealed record PartEvent(
    DocxEventKind Kind,
    long Ordinal,
    string Uri,
    string ContentType,
    string RelationshipType) : DocxEvent(Kind, Ordinal)
{
    public bool IsStart => Kind == DocxEventKind.PartStart;
}

/// <summary>Describes an OpenXML relationship from the package or current part.</summary>
public sealed record RelationshipEvent(
    long Ordinal,
    string SourceUri,
    string Id,
    string RelationshipType,
    string TargetUri,
    bool IsExternal) : DocxEvent(DocxEventKind.Relationship, Ordinal);

/// <summary>Represents an XML attribute exactly as observed on an element start.</summary>
public sealed record DocxAttribute(
    string Name,
    string LocalName,
    string Prefix,
    string NamespaceUri,
    string Value);

/// <summary>Base type for XML node events inside a package part.</summary>
public abstract record XmlPartEvent(
    DocxEventKind Kind,
    long Ordinal,
    string PartUri,
    int Depth,
    string Path) : DocxEvent(Kind, Ordinal);

/// <summary>Represents a low-level XML element start.</summary>
public sealed record ElementStartEvent(
    long Ordinal,
    string PartUri,
    string Name,
    string LocalName,
    string Prefix,
    string NamespaceUri,
    IReadOnlyList<DocxAttribute> Attributes,
    int Depth,
    string Path,
    bool IsEmptyElement) : XmlPartEvent(DocxEventKind.ElementStart, Ordinal, PartUri, Depth, Path);

/// <summary>Represents a low-level XML element end.</summary>
public sealed record ElementEndEvent(
    long Ordinal,
    string PartUri,
    string Name,
    string LocalName,
    string Prefix,
    string NamespaceUri,
    int Depth,
    string Path) : XmlPartEvent(DocxEventKind.ElementEnd, Ordinal, PartUri, Depth, Path);

/// <summary>Represents an XML text-like node inside a package part.</summary>
public sealed record TextEvent(
    long Ordinal,
    string PartUri,
    string Text,
    int Depth,
    string Path,
    bool IsWhitespace) : XmlPartEvent(DocxEventKind.Text, Ordinal, PartUri, Depth, Path);

/// <summary>Reports non-fatal parser diagnostics without projecting app-level semantics.</summary>
public sealed record DiagnosticEvent(
    long Ordinal,
    string Message,
    string? PartUri = null) : DocxEvent(DocxEventKind.Diagnostic, Ordinal);

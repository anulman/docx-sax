namespace DocxSax;

/// <summary>
/// Identifies the shape of a low-level event emitted while reading a DOCX package.
/// </summary>
public enum DocxEventKind
{
    PackageStart,
    PackageEnd,
    PartStart,
    PartEnd,
    Relationship,
    ElementStart,
    ElementEnd,
    Text,
    Diagnostic,
}

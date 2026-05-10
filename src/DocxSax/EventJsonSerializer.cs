using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Json.Serialization.Metadata;

namespace DocxSax;

/// <summary>
/// Serializes typed <see cref="DocxEvent" /> values to the stable JSON event shape used by adapter layers.
/// This is intentionally internal so JSONL remains transport plumbing, not the core library API contract.
/// </summary>
internal static class DocxEventJson
{
    public static string ToJson(DocxEvent docxEvent) => docxEvent switch
    {
        PackageEvent package => Serialize(
            new PackagePayload("package", package.IsStart ? "start" : "end", package.Ordinal),
            DocxEventJsonContext.Default.PackagePayload),
        PartEvent part => Serialize(
            new PartPayload("part", part.IsStart ? "start" : "end", part.Ordinal, part.Uri, part.ContentType, part.RelationshipType),
            DocxEventJsonContext.Default.PartPayload),
        RelationshipEvent relationship => Serialize(
            new RelationshipPayload("relationship", relationship.Ordinal, relationship.SourceUri, relationship.Id, relationship.RelationshipType, relationship.TargetUri, relationship.IsExternal),
            DocxEventJsonContext.Default.RelationshipPayload),
        ElementStartEvent element => Serialize(
            new ElementStartPayload(
                "element",
                element.Ordinal,
                element.PartUri,
                element.Name,
                element.LocalName,
                element.Prefix,
                element.NamespaceUri,
                element.Depth,
                element.Path,
                element.IsEmptyElement,
                element.Attributes.Select(ToAttributePayload).ToArray()),
            DocxEventJsonContext.Default.ElementStartPayload),
        ElementEndEvent element => Serialize(
            new ElementEndPayload("end", element.Ordinal, element.PartUri, element.Name, element.LocalName, element.Prefix, element.NamespaceUri, element.Depth, element.Path),
            DocxEventJsonContext.Default.ElementEndPayload),
        TextEvent text => Serialize(
            new TextPayload("text", text.Ordinal, text.PartUri, text.Text, text.Depth, text.Path, text.IsWhitespace),
            DocxEventJsonContext.Default.TextPayload),
        DiagnosticEvent diagnostic => Serialize(
            new DiagnosticPayload("diagnostic", diagnostic.Ordinal, diagnostic.Message, diagnostic.PartUri),
            DocxEventJsonContext.Default.DiagnosticPayload),
        _ => throw new InvalidOperationException($"Unsupported event type: {docxEvent.GetType().Name}"),
    };

    public static IEnumerable<string> ToJsonBatches(IEnumerable<DocxEvent> events, int batchSize = 128)
    {
        ArgumentNullException.ThrowIfNull(events);
        if (batchSize <= 0)
        {
            batchSize = 128;
        }

        var batch = new List<string>(batchSize);
        foreach (var docxEvent in events)
        {
            batch.Add(ToJson(docxEvent));
            if (batch.Count >= batchSize)
            {
                yield return Flush(batch);
            }
        }

        if (batch.Count > 0)
        {
            yield return Flush(batch);
        }
    }

    public static string ToJsonBatchFrames(IEnumerable<DocxEvent> events, int batchSize = 128) =>
        string.Join('\n', ToJsonBatches(events, batchSize));

    private static string Flush(List<string> batch)
    {
        var json = "[" + string.Join(",", batch) + "]";
        batch.Clear();
        return json;
    }

    private static string Serialize<TPayload>(TPayload payload, JsonTypeInfo<TPayload> jsonTypeInfo) =>
        JsonSerializer.Serialize(payload, jsonTypeInfo);

    private static AttributePayload ToAttributePayload(DocxAttribute attribute) =>
        new(attribute.Name, attribute.LocalName, attribute.Prefix, attribute.NamespaceUri, attribute.Value);

    internal sealed record PackagePayload(
        [property: JsonPropertyName("type")] string Type,
        [property: JsonPropertyName("phase")] string Phase,
        [property: JsonPropertyName("ordinal")] long Ordinal);

    internal sealed record PartPayload(
        [property: JsonPropertyName("type")] string Type,
        [property: JsonPropertyName("phase")] string Phase,
        [property: JsonPropertyName("ordinal")] long Ordinal,
        [property: JsonPropertyName("uri")] string Uri,
        [property: JsonPropertyName("contentType")] string ContentType,
        [property: JsonPropertyName("relationshipType")] string RelationshipType);

    internal sealed record RelationshipPayload(
        [property: JsonPropertyName("type")] string Type,
        [property: JsonPropertyName("ordinal")] long Ordinal,
        [property: JsonPropertyName("sourceUri")] string SourceUri,
        [property: JsonPropertyName("id")] string Id,
        [property: JsonPropertyName("relationshipType")] string RelationshipType,
        [property: JsonPropertyName("targetUri")] string TargetUri,
        [property: JsonPropertyName("isExternal")] bool IsExternal);

    internal sealed record ElementStartPayload(
        [property: JsonPropertyName("type")] string Type,
        [property: JsonPropertyName("ordinal")] long Ordinal,
        [property: JsonPropertyName("partUri")] string PartUri,
        [property: JsonPropertyName("name")] string Name,
        [property: JsonPropertyName("localName")] string LocalName,
        [property: JsonPropertyName("prefix")] string Prefix,
        [property: JsonPropertyName("namespaceUri")] string NamespaceUri,
        [property: JsonPropertyName("depth")] int Depth,
        [property: JsonPropertyName("path")] string Path,
        [property: JsonPropertyName("isEmptyElement")] bool IsEmptyElement,
        [property: JsonPropertyName("attributes")] AttributePayload[] Attributes);

    internal sealed record ElementEndPayload(
        [property: JsonPropertyName("type")] string Type,
        [property: JsonPropertyName("ordinal")] long Ordinal,
        [property: JsonPropertyName("partUri")] string PartUri,
        [property: JsonPropertyName("name")] string Name,
        [property: JsonPropertyName("localName")] string LocalName,
        [property: JsonPropertyName("prefix")] string Prefix,
        [property: JsonPropertyName("namespaceUri")] string NamespaceUri,
        [property: JsonPropertyName("depth")] int Depth,
        [property: JsonPropertyName("path")] string Path);

    internal sealed record TextPayload(
        [property: JsonPropertyName("type")] string Type,
        [property: JsonPropertyName("ordinal")] long Ordinal,
        [property: JsonPropertyName("partUri")] string PartUri,
        [property: JsonPropertyName("text")] string Text,
        [property: JsonPropertyName("depth")] int Depth,
        [property: JsonPropertyName("path")] string Path,
        [property: JsonPropertyName("isWhitespace")] bool IsWhitespace);

    internal sealed record DiagnosticPayload(
        [property: JsonPropertyName("type")] string Type,
        [property: JsonPropertyName("ordinal")] long Ordinal,
        [property: JsonPropertyName("message")] string Message,
        [property: JsonPropertyName("partUri")] string? PartUri);

    internal sealed record AttributePayload(
        [property: JsonPropertyName("name")] string Name,
        [property: JsonPropertyName("localName")] string LocalName,
        [property: JsonPropertyName("prefix")] string Prefix,
        [property: JsonPropertyName("namespaceUri")] string NamespaceUri,
        [property: JsonPropertyName("value")] string Value);
}

[JsonSourceGenerationOptions(WriteIndented = false)]
[JsonSerializable(typeof(DocxEventJson.PackagePayload))]
[JsonSerializable(typeof(DocxEventJson.PartPayload))]
[JsonSerializable(typeof(DocxEventJson.RelationshipPayload))]
[JsonSerializable(typeof(DocxEventJson.ElementStartPayload))]
[JsonSerializable(typeof(DocxEventJson.ElementEndPayload))]
[JsonSerializable(typeof(DocxEventJson.TextPayload))]
[JsonSerializable(typeof(DocxEventJson.DiagnosticPayload))]
internal sealed partial class DocxEventJsonContext : JsonSerializerContext;

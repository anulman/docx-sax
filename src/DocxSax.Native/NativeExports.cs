using System.IO.Packaging;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Json.Serialization.Metadata;
using DocxSax;
using DocumentFormat.OpenXml.Packaging;

namespace DocxSax.Native;

public static unsafe class NativeExports
{
    public const int Success = 0;
    public const int InvalidArgument = 2;
    public const int ParseFailure = 3;
    public const int CallbackFailure = 4;

    [UnmanagedCallersOnly(EntryPoint = "docx_sax_parse_file_json_batches", CallConvs = [typeof(System.Runtime.CompilerServices.CallConvCdecl)])]
    public static int ParseFileJsonBatches(byte* inputPathUtf8, int batchSize, delegate* unmanaged[Cdecl]<byte*, int, void*, int> onBatch, void* userData)
    {
        if (inputPathUtf8 is null || onBatch is null)
        {
            return InvalidArgument;
        }

        if (batchSize <= 0)
        {
            batchSize = 128;
        }

        try
        {
            var inputPath = Marshal.PtrToStringUTF8((nint)inputPathUtf8);
            if (string.IsNullOrWhiteSpace(inputPath))
            {
                return InvalidArgument;
            }

            using var stream = File.OpenRead(inputPath);
            var reader = new DocxSaxReader();
            var batch = new List<string>(batchSize);

            foreach (var docxEvent in reader.Read(stream))
            {
                batch.Add(ToJson(docxEvent));
                if (batch.Count >= batchSize && !Flush(batch, onBatch, userData))
                {
                    return CallbackFailure;
                }
            }

            return Flush(batch, onBatch, userData) ? Success : CallbackFailure;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or InvalidDataException or FileFormatException or OpenXmlPackageException or JsonException)
        {
            return ParseFailure;
        }
    }

    private static bool Flush(List<string> batch, delegate* unmanaged[Cdecl]<byte*, int, void*, int> onBatch, void* userData)
    {
        if (batch.Count == 0)
        {
            return true;
        }

        var json = "[" + string.Join(",", batch) + "]";
        var bytes = Encoding.UTF8.GetBytes(json);
        fixed (byte* pointer = bytes)
        {
            var result = onBatch(pointer, bytes.Length, userData);
            batch.Clear();
            return result == 0;
        }
    }

    private static string ToJson(DocxEvent docxEvent) => docxEvent switch
    {
        PackageEvent package => Serialize(
            new PackagePayload("package", package.IsStart ? "start" : "end", package.Ordinal),
            NativeJsonContext.Default.PackagePayload),
        PartEvent part => Serialize(
            new PartPayload("part", part.IsStart ? "start" : "end", part.Ordinal, part.Uri, part.ContentType, part.RelationshipType),
            NativeJsonContext.Default.PartPayload),
        RelationshipEvent relationship => Serialize(
            new RelationshipPayload("relationship", relationship.Ordinal, relationship.SourceUri, relationship.Id, relationship.RelationshipType, relationship.TargetUri, relationship.IsExternal),
            NativeJsonContext.Default.RelationshipPayload),
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
            NativeJsonContext.Default.ElementStartPayload),
        ElementEndEvent element => Serialize(
            new ElementEndPayload("end", element.Ordinal, element.PartUri, element.Name, element.LocalName, element.Prefix, element.NamespaceUri, element.Depth, element.Path),
            NativeJsonContext.Default.ElementEndPayload),
        TextEvent text => Serialize(
            new TextPayload("text", text.Ordinal, text.PartUri, text.Text, text.Depth, text.Path, text.IsWhitespace),
            NativeJsonContext.Default.TextPayload),
        DiagnosticEvent diagnostic => Serialize(
            new DiagnosticPayload("diagnostic", diagnostic.Ordinal, diagnostic.Message, diagnostic.PartUri),
            NativeJsonContext.Default.DiagnosticPayload),
        _ => throw new InvalidOperationException($"Unsupported event type: {docxEvent.GetType().Name}"),
    };

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
[JsonSerializable(typeof(NativeExports.PackagePayload))]
[JsonSerializable(typeof(NativeExports.PartPayload))]
[JsonSerializable(typeof(NativeExports.RelationshipPayload))]
[JsonSerializable(typeof(NativeExports.ElementStartPayload))]
[JsonSerializable(typeof(NativeExports.ElementEndPayload))]
[JsonSerializable(typeof(NativeExports.TextPayload))]
[JsonSerializable(typeof(NativeExports.DiagnosticPayload))]
internal sealed partial class NativeJsonContext : JsonSerializerContext;

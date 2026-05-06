using System.IO.Packaging;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Json.Serialization.Metadata;
using DocxSax;
using DocumentFormat.OpenXml.Packaging;

namespace DocxSax.Tool;

internal static class Program
{
    public static Task<int> Main(string[] args) => MainAsync(args, Console.Out, Console.Error);

    public static async Task<int> MainAsync(string[] args, TextWriter stdout, TextWriter stderr)
    {
        if (!TryParseArgs(args, out var inputPath, out var parseError))
        {
            await stderr.WriteLineAsync(parseError);
            await WriteUsageAsync(stderr);
            return 2;
        }

        var tempPath = Path.GetTempFileName();
        try
        {
            await using (var stream = File.OpenRead(inputPath))
            await using (var tempStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
            await using (var tempWriter = new StreamWriter(tempStream))
            {
                var reader = new DocxSaxReader();
                foreach (var docxEvent in reader.Read(stream))
                {
                    await tempWriter.WriteLineAsync(ToJsonLine(docxEvent));
                }
            }

            using var tempReader = File.OpenText(tempPath);
            while (await tempReader.ReadLineAsync() is { } line)
            {
                await stdout.WriteLineAsync(line);
            }

            return 0;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or InvalidDataException or FileFormatException or OpenXmlPackageException or ArgumentException)
        {
            await stderr.WriteLineAsync($"docx-sax: failed to parse '{inputPath}': {exception.Message}");
            return 1;
        }
        finally
        {
            try
            {
                File.Delete(tempPath);
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }
    }

    private static bool TryParseArgs(string[] args, out string inputPath, out string error)
    {
        inputPath = string.Empty;
        error = string.Empty;

        if (args.Length is 3 && args[0] == "parse" && args[2] == "--jsonl")
        {
            inputPath = args[1];
            return true;
        }

        error = "docx-sax: expected command shape: docx-sax parse <input.docx> --jsonl";
        return false;
    }

    private static Task WriteUsageAsync(TextWriter stderr) => stderr.WriteLineAsync("usage: docx-sax parse <input.docx> --jsonl");

    private static string ToJsonLine(DocxEvent docxEvent) => docxEvent switch
    {
        PackageEvent package => Serialize(
            new PackagePayload("package", package.IsStart ? "start" : "end", package.Ordinal),
            JsonLineJsonContext.Default.PackagePayload),
        PartEvent part => Serialize(
            new PartPayload("part", part.IsStart ? "start" : "end", part.Ordinal, part.Uri, part.ContentType, part.RelationshipType),
            JsonLineJsonContext.Default.PartPayload),
        RelationshipEvent relationship => Serialize(
            new RelationshipPayload("relationship", relationship.Ordinal, relationship.SourceUri, relationship.Id, relationship.RelationshipType, relationship.TargetUri, relationship.IsExternal),
            JsonLineJsonContext.Default.RelationshipPayload),
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
            JsonLineJsonContext.Default.ElementStartPayload),
        ElementEndEvent element => Serialize(
            new ElementEndPayload("end", element.Ordinal, element.PartUri, element.Name, element.LocalName, element.Prefix, element.NamespaceUri, element.Depth, element.Path),
            JsonLineJsonContext.Default.ElementEndPayload),
        TextEvent text => Serialize(
            new TextPayload("text", text.Ordinal, text.PartUri, text.Text, text.Depth, text.Path, text.IsWhitespace),
            JsonLineJsonContext.Default.TextPayload),
        DiagnosticEvent diagnostic => Serialize(
            new DiagnosticPayload("diagnostic", diagnostic.Ordinal, diagnostic.Message, diagnostic.PartUri),
            JsonLineJsonContext.Default.DiagnosticPayload),
        _ => throw new InvalidOperationException($"Unsupported event type: {docxEvent.GetType().Name}"),
    };

    private static AttributePayload ToAttributePayload(DocxAttribute attribute) =>
        new(attribute.Name, attribute.LocalName, attribute.Prefix, attribute.NamespaceUri, attribute.Value);

    private static string Serialize<TPayload>(TPayload payload, JsonTypeInfo<TPayload> jsonTypeInfo) =>
        JsonSerializer.Serialize(payload, jsonTypeInfo);

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
[JsonSerializable(typeof(Program.PackagePayload))]
[JsonSerializable(typeof(Program.PartPayload))]
[JsonSerializable(typeof(Program.RelationshipPayload))]
[JsonSerializable(typeof(Program.ElementStartPayload))]
[JsonSerializable(typeof(Program.ElementEndPayload))]
[JsonSerializable(typeof(Program.TextPayload))]
[JsonSerializable(typeof(Program.DiagnosticPayload))]
internal sealed partial class JsonLineJsonContext : JsonSerializerContext;

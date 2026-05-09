using System.IO.Packaging;
using System.Text.Json;
using DocxSax;
using DocumentFormat.OpenXml.Packaging;

namespace DocxSax.Tool;

internal static class Program
{
    public static Task<int> Main(string[] args) => MainAsync(args, Console.Out, Console.Error);

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = false,
    };

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

    private static string ToJsonLine(DocxEvent docxEvent)
    {
        object payload = docxEvent switch
        {
            PackageEvent package => new
            {
                type = "package",
                phase = package.IsStart ? "start" : "end",
                ordinal = package.Ordinal,
            },
            PartEvent part => new
            {
                type = "part",
                phase = part.IsStart ? "start" : "end",
                ordinal = part.Ordinal,
                uri = part.Uri,
                contentType = part.ContentType,
                relationshipType = part.RelationshipType,
            },
            RelationshipEvent relationship => new
            {
                type = "relationship",
                ordinal = relationship.Ordinal,
                sourceUri = relationship.SourceUri,
                id = relationship.Id,
                relationshipType = relationship.RelationshipType,
                targetUri = relationship.TargetUri,
                isExternal = relationship.IsExternal,
            },
            ElementStartEvent element => new
            {
                type = "element",
                ordinal = element.Ordinal,
                partUri = element.PartUri,
                name = element.Name,
                localName = element.LocalName,
                prefix = element.Prefix,
                namespaceUri = element.NamespaceUri,
                depth = element.Depth,
                path = element.Path,
                isEmptyElement = element.IsEmptyElement,
                attributes = element.Attributes.Select(attribute => new
                {
                    name = attribute.Name,
                    localName = attribute.LocalName,
                    prefix = attribute.Prefix,
                    namespaceUri = attribute.NamespaceUri,
                    value = attribute.Value,
                }),
            },
            ElementEndEvent element => new
            {
                type = "end",
                ordinal = element.Ordinal,
                partUri = element.PartUri,
                name = element.Name,
                localName = element.LocalName,
                prefix = element.Prefix,
                namespaceUri = element.NamespaceUri,
                depth = element.Depth,
                path = element.Path,
            },
            TextEvent text => new
            {
                type = "text",
                ordinal = text.Ordinal,
                partUri = text.PartUri,
                text = text.Text,
                depth = text.Depth,
                path = text.Path,
                isWhitespace = text.IsWhitespace,
            },
            DiagnosticEvent diagnostic => new
            {
                type = "diagnostic",
                ordinal = diagnostic.Ordinal,
                message = diagnostic.Message,
                partUri = diagnostic.PartUri,
            },
            _ => throw new InvalidOperationException($"Unsupported event type: {docxEvent.GetType().Name}"),
        };

        return JsonSerializer.Serialize(payload, JsonOptions);
    }
}

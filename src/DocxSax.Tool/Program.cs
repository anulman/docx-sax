using System.IO.Packaging;
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
                    await tempWriter.WriteLineAsync(DocxEventJson.ToJson(docxEvent));
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
}

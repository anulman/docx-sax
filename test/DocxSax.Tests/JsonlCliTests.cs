using System.Diagnostics;
using System.Text.Json;
using Xunit;

namespace DocxSax.Tests;

public sealed class JsonlCliTests
{
    [Fact]
    public async Task ParseJsonl_MatchesGoldenAndEveryLineIsValidJson()
    {
        var fixture = FixturePath("simple.docx");
        var golden = await File.ReadAllTextAsync(GoldenPath("simple.jsonl"));

        var result = await RunCliAsync(fixture);

        Assert.Equal(0, result.ExitCode);
        Assert.Equal(string.Empty, result.StandardError);
        Assert.Equal(NormalizeNewlines(golden), NormalizeNewlines(result.StandardOutput));

        foreach (var line in JsonLines(result.StandardOutput))
        {
            using var _ = JsonDocument.Parse(line);
        }
    }

    [Fact]
    public async Task ParseJsonl_EmitsExpectedStructuralOrderAndIsDeterministic()
    {
        var fixture = FixturePath("simple.docx");

        var first = await RunCliAsync(fixture);
        var second = await RunCliAsync(fixture);

        Assert.Equal(0, first.ExitCode);
        Assert.Equal(0, second.ExitCode);
        Assert.Equal(first.StandardOutput, second.StandardOutput);

        var events = JsonLines(first.StandardOutput)
            .Select(line => JsonDocument.Parse(line))
            .ToArray();

        try
        {
            Assert.Equal("package", StringAt(events[0], "type"));
            Assert.Equal("start", StringAt(events[0], "phase"));
            Assert.Equal("package", StringAt(events[^1], "type"));
            Assert.Equal("end", StringAt(events[^1], "phase"));

            AssertOrder(events, "relationship", "part", "element", "text", "end");
            Assert.Contains(events, e => StringAt(e, "type") == "part" && StringAt(e, "phase") == "end");
        }
        finally
        {
            foreach (var document in events)
            {
                document.Dispose();
            }
        }
    }

    [Fact]
    public async Task ParseJsonl_PreservesUnknownXmlElementAndAttribute()
    {
        var fixture = FixturePath("unknown-inline.docx");
        var golden = await File.ReadAllTextAsync(GoldenPath("unknown-inline.jsonl"));

        var result = await RunCliAsync(fixture);

        Assert.Equal(0, result.ExitCode);
        Assert.Equal(NormalizeNewlines(golden), NormalizeNewlines(result.StandardOutput));

        var lines = JsonLines(result.StandardOutput).ToArray();
        Assert.Contains(lines, line => line.Contains("\"name\":\"x:unknown\"", StringComparison.Ordinal));
        Assert.Contains(lines, line => line.Contains("\"namespaceUri\":\"urn:docx-sax:test\"", StringComparison.Ordinal));
        Assert.Contains(lines, line => line.Contains("\"name\":\"x:flag\"", StringComparison.Ordinal));
        Assert.Contains(lines, line => line.Contains("\"text\":\"Mystery node\"", StringComparison.Ordinal));
    }


    [Fact]
    public async Task ParseJsonl_RequiresExplicitJsonlFlag()
    {
        var result = await RunCliAsync(FixturePath("simple.docx"), includeJsonlFlag: false);

        Assert.Equal(2, result.ExitCode);
        Assert.Equal(string.Empty, result.StandardOutput);
        Assert.Contains("--jsonl", result.StandardError, StringComparison.Ordinal);
        Assert.Contains("usage:", result.StandardError, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ParseJsonl_CorruptDocxExitsNonzeroAndWritesDiagnosticToStderr()
    {
        var result = await RunCliAsync(FixturePath("corrupt.docx"));

        Assert.NotEqual(0, result.ExitCode);
        Assert.Equal(string.Empty, result.StandardOutput);
        Assert.Contains("docx-sax: failed to parse", result.StandardError, StringComparison.Ordinal);
        Assert.Contains("corrupt.docx", result.StandardError, StringComparison.Ordinal);
    }

    private static void AssertOrder(JsonDocument[] events, params string[] expectedTypes)
    {
        var cursor = -1;
        foreach (var expectedType in expectedTypes)
        {
            var next = Array.FindIndex(events, cursor + 1, e => StringAt(e, "type") == expectedType);
            Assert.True(next > cursor, $"Expected event type '{expectedType}' after index {cursor}.");
            cursor = next;
        }
    }

    private static string StringAt(JsonDocument document, string propertyName) =>
        document.RootElement.GetProperty(propertyName).GetString() ?? string.Empty;

    private static IEnumerable<string> JsonLines(string text) =>
        NormalizeNewlines(text).Split('\n', StringSplitOptions.RemoveEmptyEntries);

    private static string FixturePath(string fileName) => Path.Combine(RepoRoot(), "test", "DocxSax.Tests", "Fixtures", fileName);

    private static string GoldenPath(string fileName) => Path.Combine(RepoRoot(), "test", "DocxSax.Tests", "Golden", fileName);

    private static string RepoRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "DocxSax.sln")))
        {
            directory = directory.Parent;
        }

        return directory?.FullName ?? throw new InvalidOperationException("Could not locate repository root.");
    }

    private static Task<CliResult> RunCliAsync(string inputPath) => RunCliAsync(inputPath, includeJsonlFlag: true);

    private static async Task<CliResult> RunCliAsync(string inputPath, bool includeJsonlFlag)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "dotnet",
            WorkingDirectory = RepoRoot(),
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        startInfo.ArgumentList.Add("run");
        startInfo.ArgumentList.Add("--project");
        startInfo.ArgumentList.Add(Path.Combine(RepoRoot(), "src", "DocxSax.Tool", "DocxSax.Tool.csproj"));
        startInfo.ArgumentList.Add("--configuration");
        startInfo.ArgumentList.Add("Release");
        startInfo.ArgumentList.Add("--");
        startInfo.ArgumentList.Add("parse");
        startInfo.ArgumentList.Add(inputPath);
        if (includeJsonlFlag)
        {
            startInfo.ArgumentList.Add("--jsonl");
        }

        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start dotnet CLI.");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();

        await process.WaitForExitAsync();
        var stdout = await stdoutTask;
        var stderr = await stderrTask;

        return new CliResult(process.ExitCode, stdout, stderr);
    }

    private static string NormalizeNewlines(string text) => text.Replace("\r\n", "\n", StringComparison.Ordinal);

    private sealed record CliResult(int ExitCode, string StandardOutput, string StandardError);
}

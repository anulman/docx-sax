using System.IO.Packaging;
using System.Runtime.InteropServices;
using DocxSax;
using DocumentFormat.OpenXml.Packaging;

namespace DocxSax.Native;

public static unsafe class NativeExports
{
    public const int Success = 0;
    public const int InvalidArgument = 2;
    public const int ParseFailure = 3;
    public const int CallbackFailure = 4;

    [UnmanagedCallersOnly(EntryPoint = "docx_sax_parse_file_events", CallConvs = [typeof(System.Runtime.CompilerServices.CallConvCdecl)])]
    public static int ParseFileEvents(byte* inputPathUtf8, int batchSize, delegate* unmanaged[Cdecl]<NativeDocxEvent*, void*, int> onEvent, void* userData)
    {
        if (inputPathUtf8 is null || onEvent is null)
        {
            return InvalidArgument;
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

            foreach (var docxEvent in reader.Read(stream))
            {
                using var marshaled = MarshaledDocxEvent.From(docxEvent);
                var payload = marshaled.Payload;
                if (onEvent(&payload, userData) != 0)
                {
                    return CallbackFailure;
                }
            }

            return Success;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or InvalidDataException or FileFormatException or OpenXmlPackageException)
        {
            return ParseFailure;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct NativeString
    {
        public byte* Data;
        public int Length;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct NativeAttribute
    {
        public NativeString Name;
        public NativeString LocalName;
        public NativeString Prefix;
        public NativeString NamespaceUri;
        public NativeString Value;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct NativeDocxEvent
    {
        public int Kind;
        public long Ordinal;
        public NativeString Phase;
        public NativeString Uri;
        public NativeString ContentType;
        public NativeString RelationshipType;
        public NativeString SourceUri;
        public NativeString Id;
        public NativeString TargetUri;
        public int IsExternal;
        public NativeString PartUri;
        public NativeString Name;
        public NativeString LocalName;
        public NativeString Prefix;
        public NativeString NamespaceUri;
        public int Depth;
        public NativeString Path;
        public int IsEmptyElement;
        public NativeAttribute* Attributes;
        public int AttributeCount;
        public NativeString Text;
        public int IsWhitespace;
        public NativeString Message;
    }

    private sealed class MarshaledDocxEvent : IDisposable
    {
        private readonly List<nint> allocations = [];
        private nint attributesAllocation;

        private MarshaledDocxEvent()
        {
        }

        public NativeDocxEvent Payload;

        public static MarshaledDocxEvent From(DocxEvent docxEvent)
        {
            var marshaled = new MarshaledDocxEvent();
            marshaled.Payload.Ordinal = docxEvent.Ordinal;

            switch (docxEvent)
            {
                case PackageEvent package:
                    marshaled.Payload.Kind = (int)docxEvent.Kind;
                    marshaled.Payload.Phase = marshaled.String(package.IsStart ? "start" : "end");
                    break;
                case PartEvent part:
                    marshaled.Payload.Kind = (int)docxEvent.Kind;
                    marshaled.Payload.Phase = marshaled.String(part.IsStart ? "start" : "end");
                    marshaled.Payload.Uri = marshaled.String(part.Uri);
                    marshaled.Payload.ContentType = marshaled.String(part.ContentType);
                    marshaled.Payload.RelationshipType = marshaled.String(part.RelationshipType);
                    break;
                case RelationshipEvent relationship:
                    marshaled.Payload.Kind = (int)docxEvent.Kind;
                    marshaled.Payload.SourceUri = marshaled.String(relationship.SourceUri);
                    marshaled.Payload.Id = marshaled.String(relationship.Id);
                    marshaled.Payload.RelationshipType = marshaled.String(relationship.RelationshipType);
                    marshaled.Payload.TargetUri = marshaled.String(relationship.TargetUri);
                    marshaled.Payload.IsExternal = relationship.IsExternal ? 1 : 0;
                    break;
                case ElementStartEvent element:
                    marshaled.Payload.Kind = (int)docxEvent.Kind;
                    marshaled.Payload.PartUri = marshaled.String(element.PartUri);
                    marshaled.Payload.Name = marshaled.String(element.Name);
                    marshaled.Payload.LocalName = marshaled.String(element.LocalName);
                    marshaled.Payload.Prefix = marshaled.String(element.Prefix);
                    marshaled.Payload.NamespaceUri = marshaled.String(element.NamespaceUri);
                    marshaled.Payload.Depth = element.Depth;
                    marshaled.Payload.Path = marshaled.String(element.Path);
                    marshaled.Payload.IsEmptyElement = element.IsEmptyElement ? 1 : 0;
                    marshaled.SetAttributes(element.Attributes);
                    break;
                case ElementEndEvent element:
                    marshaled.Payload.Kind = (int)docxEvent.Kind;
                    marshaled.Payload.PartUri = marshaled.String(element.PartUri);
                    marshaled.Payload.Name = marshaled.String(element.Name);
                    marshaled.Payload.LocalName = marshaled.String(element.LocalName);
                    marshaled.Payload.Prefix = marshaled.String(element.Prefix);
                    marshaled.Payload.NamespaceUri = marshaled.String(element.NamespaceUri);
                    marshaled.Payload.Depth = element.Depth;
                    marshaled.Payload.Path = marshaled.String(element.Path);
                    break;
                case TextEvent text:
                    marshaled.Payload.Kind = (int)docxEvent.Kind;
                    marshaled.Payload.PartUri = marshaled.String(text.PartUri);
                    marshaled.Payload.Text = marshaled.String(text.Text);
                    marshaled.Payload.Depth = text.Depth;
                    marshaled.Payload.Path = marshaled.String(text.Path);
                    marshaled.Payload.IsWhitespace = text.IsWhitespace ? 1 : 0;
                    break;
                case DiagnosticEvent diagnostic:
                    marshaled.Payload.Kind = (int)docxEvent.Kind;
                    marshaled.Payload.Message = marshaled.String(diagnostic.Message);
                    marshaled.Payload.PartUri = marshaled.String(diagnostic.PartUri);
                    break;
                default:
                    throw new InvalidOperationException($"Unsupported event type: {docxEvent.GetType().Name}");
            }

            return marshaled;
        }

        public void Dispose()
        {
            foreach (var allocation in allocations)
            {
                Marshal.FreeCoTaskMem(allocation);
            }

            if (attributesAllocation != 0)
            {
                Marshal.FreeCoTaskMem(attributesAllocation);
            }
        }

        private NativeString String(string? value)
        {
            if (value is null)
            {
                return default;
            }

            var bytes = System.Text.Encoding.UTF8.GetByteCount(value);
            var allocation = Marshal.AllocCoTaskMem(bytes + 1);
            allocations.Add(allocation);

            var span = new Span<byte>((void*)allocation, bytes + 1);
            System.Text.Encoding.UTF8.GetBytes(value, span[..bytes]);
            span[bytes] = 0;
            return new NativeString { Data = (byte*)allocation, Length = bytes };
        }

        private void SetAttributes(IReadOnlyList<DocxAttribute> attributes)
        {
            if (attributes.Count == 0)
            {
                return;
            }

            var size = sizeof(NativeAttribute) * attributes.Count;
            attributesAllocation = Marshal.AllocCoTaskMem(size);
            var nativeAttributes = new Span<NativeAttribute>((void*)attributesAllocation, attributes.Count);

            for (var i = 0; i < attributes.Count; i++)
            {
                var attribute = attributes[i];
                nativeAttributes[i] = new NativeAttribute
                {
                    Name = String(attribute.Name),
                    LocalName = String(attribute.LocalName),
                    Prefix = String(attribute.Prefix),
                    NamespaceUri = String(attribute.NamespaceUri),
                    Value = String(attribute.Value),
                };
            }

            Payload.Attributes = (NativeAttribute*)attributesAllocation;
            Payload.AttributeCount = attributes.Count;
        }
    }
}

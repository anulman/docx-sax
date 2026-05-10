export type DocxSaxEvent =
  | DocxSaxPackageEvent
  | DocxSaxPartEvent
  | DocxSaxRelationshipEvent
  | DocxSaxElementEvent
  | DocxSaxTextEvent
  | DocxSaxElementEndEvent
  | DocxSaxDiagnosticEvent;

export type DocxSaxBytesInput = Uint8Array | ArrayBuffer | ArrayBufferView | Blob;

export interface DocxSaxBatchOptions {
  /** Maximum number of events per yielded batch. Defaults to 128 when omitted or invalid. */
  batchSize?: number;
  /** Minimum milliseconds between cooperative main-thread yields. Defaults to 64. */
  mainThreadYieldIntervalMs?: number;
}

export interface BrowserRuntimeOptions {
  /** URL for the published .NET browser runtime module. Defaults to ./dist/wasm/wwwroot/_framework/dotnet.js. */
  dotnetModuleUrl?: string;
}

export interface BrowserParseOptions extends DocxSaxBatchOptions, BrowserRuntimeOptions {}

export interface DocxSaxPackageEvent {
  type: 'package';
  phase: 'start' | 'end';
  /** Zero-based ordinal in the transport-neutral event stream. */
  ordinal: number;
}

export interface DocxSaxPartEvent {
  type: 'part';
  phase: 'start' | 'end';
  ordinal: number;
  uri: string;
  contentType: string;
  relationshipType: string;
}

export interface DocxSaxRelationshipEvent {
  type: 'relationship';
  ordinal: number;
  sourceUri: string;
  id: string;
  relationshipType: string;
  targetUri: string;
  isExternal: boolean;
}

export interface DocxSaxAttribute {
  name: string;
  localName: string;
  prefix: string;
  namespaceUri: string;
  value: string;
}

export interface DocxSaxXmlEventBase {
  ordinal: number;
  partUri: string;
  name: string;
  localName: string;
  prefix: string;
  namespaceUri: string;
  depth: number;
  path: string;
}

export interface DocxSaxElementEvent extends DocxSaxXmlEventBase {
  type: 'element';
  isEmptyElement: boolean;
  attributes: DocxSaxAttribute[];
}

export interface DocxSaxElementEndEvent extends DocxSaxXmlEventBase {
  type: 'end';
}

export interface DocxSaxTextEvent {
  type: 'text';
  ordinal: number;
  partUri: string;
  text: string;
  depth: number;
  path: string;
  isWhitespace: boolean;
}

export interface DocxSaxDiagnosticEvent {
  type: 'diagnostic';
  ordinal: number;
  message: string;
  partUri?: string | null;
}

/** Start loading and initializing the browser WASM runtime before the first parse. */
export function preloadRuntime(options?: BrowserRuntimeOptions): Promise<void>;

/** Warm OpenXML/package parsing paths after the WASM runtime loads. */
export function warmupRuntime(options?: BrowserRuntimeOptions): Promise<void>;

/** Parse DOCX bytes/blob through the browser WASM bridge and yield transport-neutral DocxSax events. */
export function parseBytes(input: DocxSaxBytesInput, options?: BrowserParseOptions): AsyncIterable<DocxSaxEvent>;

/** Parse DOCX bytes/blob through the browser WASM bridge and yield arrays of transport-neutral DocxSax events. */
export function parseBytesBatches(input: DocxSaxBytesInput, options?: BrowserParseOptions): AsyncIterable<DocxSaxEvent[]>;

declare const defaultExport: {
  parseBytes: typeof parseBytes;
  parseBytesBatches: typeof parseBytesBatches;
  preloadRuntime: typeof preloadRuntime;
  warmupRuntime: typeof warmupRuntime;
};

export default defaultExport;

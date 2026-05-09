export type DocxSaxEvent =
  | DocxSaxPackageEvent
  | DocxSaxPartEvent
  | DocxSaxRelationshipEvent
  | DocxSaxElementEvent
  | DocxSaxTextEvent
  | DocxSaxElementEndEvent
  | DocxSaxDiagnosticEvent;

export interface DocxSaxBatchOptions {
  /** Maximum number of events per yielded batch. Defaults to 128 when omitted or invalid. */
  batchSize?: number;
}

export interface NodeParseOptions extends DocxSaxBatchOptions {
  /** Override path to the Native AOT DocxSax.Native shared library. Mostly useful for local validation. */
  nativeLibraryPath?: string;
}

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

/** Parse a DOCX file path through the Native bridge and yield transport-neutral DocxSax events. */
export function parseFile(path: string, options?: NodeParseOptions): AsyncIterable<DocxSaxEvent>;

/** Parse a DOCX file path through the Native bridge and yield arrays of transport-neutral DocxSax events. */
export function parseFileBatches(path: string, options?: NodeParseOptions): AsyncIterable<DocxSaxEvent[]>;

declare const defaultExport: {
  parseFile: typeof parseFile;
  parseFileBatches: typeof parseFileBatches;
};

export default defaultExport;

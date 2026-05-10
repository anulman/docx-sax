export interface DocxSaxWasmViteOptions {
  /** Public mount path for runtime assets. Defaults to /docx-sax. */
  mount?: string;
  /** Clean the output _framework directory before copying during build. Defaults to true. */
  clean?: boolean;
}

export function docxSaxWasm(options?: DocxSaxWasmViteOptions): unknown;
export default docxSaxWasm;

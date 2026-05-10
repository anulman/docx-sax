export interface DocxSaxWasmWebpackOptions {
  /** Public output path for runtime assets. Defaults to docx-sax. */
  mount?: string;
}

export class DocxSaxWasmWebpackPlugin {
  constructor(options?: DocxSaxWasmWebpackOptions);
  apply(compiler: unknown): void;
}

export function docxSaxWasm(options?: DocxSaxWasmWebpackOptions): DocxSaxWasmWebpackPlugin;
export default DocxSaxWasmWebpackPlugin;

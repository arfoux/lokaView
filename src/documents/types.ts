import type { ComponentType } from "react";

export type DocumentKind = "pdf" | "docx" | "xlsx" | "pptx" | "csv";
export type LegacyOfficeKind = "doc" | "xls" | "ppt";

export type ViewerCapability =
  | "zoom"
  | "search"
  | "pagination"
  | "fullscreen"
  | "download-original"
  | "rotate"
  | "sheet-navigation"
  | "slide-navigation"
  | "text-extraction"
  | "export-csv"
  | "export-json";

export interface OpenDocumentContext {
  file: File;
  arrayBuffer: ArrayBuffer;
  signature: Uint8Array;
  detection: FileTypeDetection;
  signal?: AbortSignal;
}

export interface DocumentAdapter<TDocument extends OpenedDocument = OpenedDocument> {
  readonly kind: DocumentKind;
  readonly label: string;
  readonly acceptedExtensions: string[];
  readonly acceptedMimeTypes: string[];
  canHandle(file: File, signature?: Uint8Array, detection?: FileTypeDetection): boolean;
  open(context: OpenDocumentContext): Promise<TDocument>;
}

export interface OpenedDocumentBase {
  readonly kind: DocumentKind;
  readonly fileName: string;
  readonly fileSize: number;
  readonly fileType: string;
  readonly originalFile: File;
  readonly originalUrl: string;
  readonly capabilities: Set<ViewerCapability>;
  readonly warnings: string[];
  readonly detectedBy: FileTypeDetection["confidence"];
  readonly openedAt: Date;
  dispose(): void;
}

export interface PdfOpenedDocument extends OpenedDocumentBase {
  readonly kind: "pdf";
  readonly arrayBuffer: ArrayBuffer;
}

export interface DocxOpenedDocument extends OpenedDocumentBase {
  readonly kind: "docx";
  readonly arrayBuffer: ArrayBuffer;
}

export interface PptxOpenedDocument extends OpenedDocumentBase {
  readonly kind: "pptx";
  readonly arrayBuffer: ArrayBuffer;
}

export interface SpreadsheetMerge {
  startRow: number;
  startColumn: number;
  rowSpan: number;
  columnSpan: number;
}

export interface SpreadsheetCellStyle {
  backgroundColor?: string;
  color?: string;
  fontWeight?: "bold";
  fontStyle?: "italic";
  textDecoration?: "underline";
  horizontalAlign?: "left" | "center" | "right";
}

export interface SpreadsheetSheet {
  name: string;
  rows: string[][];
  rowCount: number;
  columnCount: number;
  truncatedRows: boolean;
  truncatedColumns: boolean;
  merges: SpreadsheetMerge[];
  styles: Record<string, SpreadsheetCellStyle>;
}

export interface XlsxOpenedDocument extends OpenedDocumentBase {
  readonly kind: "xlsx";
  readonly sheets: SpreadsheetSheet[];
  readonly workbookWarnings: string[];
}

export interface CsvOpenedDocument extends OpenedDocumentBase {
  readonly kind: "csv";
  readonly sheet: SpreadsheetSheet;
  readonly delimiter: string;
  readonly delimiterConfidence: "high" | "medium" | "low";
}

export type OpenedDocument =
  | PdfOpenedDocument
  | DocxOpenedDocument
  | XlsxOpenedDocument
  | PptxOpenedDocument
  | CsvOpenedDocument;

export type DocumentViewerComponent<TDocument extends OpenedDocument = OpenedDocument> =
  ComponentType<{
    document: TDocument;
  }>;

export type FileTypeConfidence = "signature" | "zip-manifest" | "mime" | "extension" | "unknown";

export interface FileTypeDetection {
  kind?: DocumentKind;
  legacyKind?: LegacyOfficeKind;
  confidence: FileTypeConfidence;
  extension?: string;
  mimeType?: string;
  isZip: boolean;
  isLegacyOffice: boolean;
  message?: string;
  warnings: string[];
}

export interface DocumentErrorOptions {
  code: DocumentErrorCode;
  title: string;
  message: string;
  cause?: unknown;
  recoverable?: boolean;
}

export type DocumentErrorCode =
  | "unsupported-format"
  | "legacy-format"
  | "file-too-large"
  | "unsafe-zip"
  | "read-aborted"
  | "read-failed"
  | "parse-failed"
  | "encrypted-document"
  | "renderer-failed";

export class DocumentError extends Error {
  readonly code: DocumentErrorCode;
  readonly title: string;
  readonly recoverable: boolean;
  override readonly cause?: unknown;

  constructor(options: DocumentErrorOptions) {
    super(options.message);
    this.name = "DocumentError";
    this.code = options.code;
    this.title = options.title;
    this.cause = options.cause;
    this.recoverable = options.recoverable ?? true;
  }
}

export interface OpenLocalDocumentOptions {
  signal?: AbortSignal;
  allowLargeFile?: boolean;
  onProgress?: (progress: ReadProgress) => void;
}

export interface ReadProgress {
  loadedBytes: number;
  totalBytes: number;
  ratio: number;
}

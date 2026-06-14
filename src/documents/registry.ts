import { csvAdapter } from "./csv/csvAdapter";
import { docxAdapter } from "./docx/docxAdapter";
import { pdfAdapter } from "./pdf/pdfAdapter";
import { pptxAdapter } from "./pptx/pptxAdapter";
import type { DocumentAdapter, DocumentKind, FileTypeDetection, OpenDocumentContext } from "./types";
import { DocumentError } from "./types";
import { xlsxAdapter } from "./xlsx/xlsxAdapter";

const adapters: DocumentAdapter[] = [pdfAdapter, docxAdapter, xlsxAdapter, pptxAdapter, csvAdapter];

export function getDocumentAdapters(): readonly DocumentAdapter[] {
  return adapters;
}

export function getAdapterByKind(kind: DocumentKind): DocumentAdapter {
  const adapter = adapters.find((candidate) => candidate.kind === kind);

  if (!adapter) {
    throw new DocumentError({
      code: "unsupported-format",
      title: "No viewer is available",
      message: `No viewer adapter is registered for ${kind.toUpperCase()} files.`,
    });
  }

  return adapter;
}

export function findAdapter(
  file: File,
  signature: Uint8Array,
  detection: FileTypeDetection,
): DocumentAdapter {
  if (detection.kind) {
    const adapter = getAdapterByKind(detection.kind);
    if (adapter.canHandle(file, signature, detection)) {
      return adapter;
    }
  }

  const adapter = adapters.find((candidate) => candidate.canHandle(file, signature, detection));

  if (!adapter) {
    throw new DocumentError({
      code: "unsupported-format",
      title: "Unsupported file",
      message: detection.message ?? "Open a PDF, DOCX, XLSX, PPTX, or CSV file.",
    });
  }

  return adapter;
}

export async function openWithRegisteredAdapter(context: OpenDocumentContext) {
  const adapter = findAdapter(context.file, context.signature, context.detection);
  return adapter.open(context);
}

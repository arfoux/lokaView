import type { DocumentKind, FileTypeDetection, LegacyOfficeKind } from "./types";
import { inspectZip } from "./zipInspection";

const MIME_KIND_MAP: Record<string, DocumentKind> = {
  "application/pdf": "pdf",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
  "application/vnd.openxmlformats-officedocument.presentationml.presentation": "pptx",
  "text/csv": "csv",
  "application/csv": "csv",
};

const EXTENSION_KIND_MAP: Record<string, DocumentKind> = {
  pdf: "pdf",
  docx: "docx",
  xlsx: "xlsx",
  pptx: "pptx",
  csv: "csv",
};

const LEGACY_EXTENSION_MAP: Record<string, LegacyOfficeKind> = {
  doc: "doc",
  xls: "xls",
  ppt: "ppt",
};

function getExtension(fileName: string): string | undefined {
  const dotIndex = fileName.lastIndexOf(".");
  if (dotIndex < 0 || dotIndex === fileName.length - 1) {
    return undefined;
  }
  return fileName.slice(dotIndex + 1).toLowerCase();
}

function isPdf(signature: Uint8Array): boolean {
  return (
    signature[0] === 0x25 &&
    signature[1] === 0x50 &&
    signature[2] === 0x44 &&
    signature[3] === 0x46 &&
    signature[4] === 0x2d
  );
}

function isCompoundBinaryFile(signature: Uint8Array): boolean {
  const expected = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  return expected.every((value, index) => signature[index] === value);
}

function detectOfficeZipKind(arrayBuffer: ArrayBuffer): {
  kind?: DocumentKind;
  warnings: string[];
  isZip: boolean;
} {
  const inspection = inspectZip(arrayBuffer);
  const names = new Set(inspection.entries.map((entry) => entry.name.toLowerCase()));
  const hasPrefix = (prefix: string) =>
    inspection.entries.some((entry) => entry.name.toLowerCase().startsWith(prefix));

  if (names.has("word/document.xml") || hasPrefix("word/")) {
    return { kind: "docx", warnings: inspection.warnings, isZip: inspection.isZip };
  }

  if (names.has("xl/workbook.xml") || hasPrefix("xl/worksheets/")) {
    return { kind: "xlsx", warnings: inspection.warnings, isZip: inspection.isZip };
  }

  if (names.has("ppt/presentation.xml") || hasPrefix("ppt/slides/")) {
    return { kind: "pptx", warnings: inspection.warnings, isZip: inspection.isZip };
  }

  return { warnings: inspection.warnings, isZip: inspection.isZip };
}

export function detectFileType(
  file: File,
  signature: Uint8Array,
  arrayBuffer?: ArrayBuffer,
): FileTypeDetection {
  const extension = getExtension(file.name);
  const mimeType = file.type.toLowerCase() || undefined;
  const warnings: string[] = [];
  const extensionKind = extension ? EXTENSION_KIND_MAP[extension] : undefined;
  const mimeKind = mimeType ? MIME_KIND_MAP[mimeType] : undefined;
  const legacyKind = extension ? LEGACY_EXTENSION_MAP[extension] : undefined;
  const isLegacyOffice = Boolean(legacyKind) || isCompoundBinaryFile(signature);

  if (legacyKind || isCompoundBinaryFile(signature)) {
    return {
      legacyKind,
      confidence: "signature",
      extension,
      mimeType,
      isZip: false,
      isLegacyOffice,
      warnings,
      message: "This older Office format is not supported yet. Please save the file as DOCX, XLSX, or PPTX and try again.",
    };
  }

  if (isPdf(signature)) {
    if (extensionKind && extensionKind !== "pdf") {
      warnings.push("The file content looks like a PDF even though the extension says otherwise.");
    }

    return {
      kind: "pdf",
      confidence: "signature",
      extension,
      mimeType,
      isZip: false,
      isLegacyOffice,
      warnings,
    };
  }

  if (arrayBuffer) {
    const zipDetection = detectOfficeZipKind(arrayBuffer);

    if (zipDetection.warnings.length > 0) {
      warnings.push(...zipDetection.warnings);
    }

    if (zipDetection.kind) {
      if (extensionKind && extensionKind !== zipDetection.kind) {
        warnings.push("The file package content does not match the extension. The detected package type will be used.");
      }

      return {
        kind: zipDetection.kind,
        confidence: "zip-manifest",
        extension,
        mimeType,
        isZip: zipDetection.isZip,
        isLegacyOffice,
        warnings,
      };
    }
  }

  if (extensionKind) {
    return {
      kind: extensionKind,
      confidence: "extension",
      extension,
      mimeType,
      isZip: signature[0] === 0x50 && signature[1] === 0x4b,
      isLegacyOffice,
      warnings,
    };
  }

  if (mimeKind) {
    return {
      kind: mimeKind,
      confidence: "mime",
      extension,
      mimeType,
      isZip: signature[0] === 0x50 && signature[1] === 0x4b,
      isLegacyOffice,
      warnings,
    };
  }

  return {
    confidence: "unknown",
    extension,
    mimeType,
    isZip: signature[0] === 0x50 && signature[1] === 0x4b,
    isLegacyOffice,
    warnings,
    message: "This file type is not supported yet. Open a PDF, DOCX, XLSX, PPTX, or CSV file.",
  };
}

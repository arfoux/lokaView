import { createOpenedDocumentBase } from "../openedDocument";
import type { DocxOpenedDocument, DocumentAdapter } from "../types";
import { DocumentError } from "../types";
import { assertZipWithinLimits, inspectZip } from "../zipInspection";

export const docxAdapter: DocumentAdapter<DocxOpenedDocument> = {
  kind: "docx",
  label: "Word",
  acceptedExtensions: [".docx"],
  acceptedMimeTypes: ["application/vnd.openxmlformats-officedocument.wordprocessingml.document"],
  canHandle(_file, _signature, detection) {
    return detection?.kind === "docx";
  },
  open({ file, arrayBuffer, detection }) {
    const zip = inspectZip(arrayBuffer);
    assertZipWithinLimits(zip);

    if (!zip.isZip) {
      throw new DocumentError({
        code: "parse-failed",
        title: "This is not a valid DOCX package",
        message: "DOCX files must be ZIP-based Office Open XML packages. This file could not be inspected as one.",
      });
    }

    if (zip.isZip && zip.entries.length > 0) {
      const hasDocument = zip.entries.some((entry) => entry.name.toLowerCase() === "word/document.xml");

      if (!hasDocument) {
        throw new DocumentError({
          code: "parse-failed",
          title: "This DOCX package is missing document content",
          message: "The file looks like a DOCX package, but the main Word document part was not found.",
        });
      }
    }

    const base = createOpenedDocumentBase({
      kind: "docx",
      file,
      detection,
      capabilities: ["download-original", "zoom", "text-extraction", "local-editing"],
      warnings: zip.warnings,
    });

    return Promise.resolve({
      ...base,
      kind: "docx",
      arrayBuffer,
    });
  },
};

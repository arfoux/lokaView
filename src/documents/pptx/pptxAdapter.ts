import { createOpenedDocumentBase } from "../openedDocument";
import type { DocumentAdapter, PptxOpenedDocument } from "../types";
import { DocumentError } from "../types";
import { assertZipWithinLimits, inspectZip } from "../zipInspection";

export const pptxAdapter: DocumentAdapter<PptxOpenedDocument> = {
  kind: "pptx",
  label: "PowerPoint",
  acceptedExtensions: [".pptx"],
  acceptedMimeTypes: ["application/vnd.openxmlformats-officedocument.presentationml.presentation"],
  canHandle(_file, _signature, detection) {
    return detection?.kind === "pptx";
  },
  open({ file, arrayBuffer, detection }) {
    const zip = inspectZip(arrayBuffer);
    assertZipWithinLimits(zip);

    if (!zip.isZip) {
      throw new DocumentError({
        code: "parse-failed",
        title: "This is not a valid PPTX package",
        message: "PPTX files must be ZIP-based Office Open XML packages. This file could not be inspected as one.",
      });
    }

    if (zip.isZip && zip.entries.length > 0) {
      const hasPresentation = zip.entries.some(
        (entry) => entry.name.toLowerCase() === "ppt/presentation.xml",
      );

      if (!hasPresentation) {
        throw new DocumentError({
          code: "parse-failed",
          title: "This PPTX package is missing presentation content",
          message: "The file looks like a PPTX package, but the main presentation part was not found.",
        });
      }
    }

    const base = createOpenedDocumentBase({
      kind: "pptx",
      file,
      detection,
      capabilities: ["download-original", "slide-navigation", "fullscreen", "zoom", "search"],
      warnings: zip.warnings,
    });

    return Promise.resolve({
      ...base,
      kind: "pptx",
      arrayBuffer,
    });
  },
};

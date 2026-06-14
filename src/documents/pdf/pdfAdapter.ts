import { createOpenedDocumentBase } from "../openedDocument";
import type { DocumentAdapter, PdfOpenedDocument } from "../types";

export const pdfAdapter: DocumentAdapter<PdfOpenedDocument> = {
  kind: "pdf",
  label: "PDF",
  acceptedExtensions: [".pdf"],
  acceptedMimeTypes: ["application/pdf"],
  canHandle(_file, signature, detection) {
    return (
      detection?.kind === "pdf" ||
      (signature?.[0] === 0x25 &&
        signature?.[1] === 0x50 &&
        signature?.[2] === 0x44 &&
        signature?.[3] === 0x46)
    );
  },
  open({ file, arrayBuffer, detection }) {
    const base = createOpenedDocumentBase({
      kind: "pdf",
      file,
      detection,
      capabilities: ["download-original", "zoom", "pagination", "rotate", "fullscreen"],
    });

    return Promise.resolve({
      ...base,
      kind: "pdf",
      arrayBuffer,
    });
  },
};

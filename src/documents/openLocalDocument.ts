import { detectFileType } from "./detectFileType";
import { assertFileSize, validateFileSize } from "./limits";
import { openWithRegisteredAdapter } from "./registry";
import { getSignature, readLocalFile } from "./readLocalFile";
import { DocumentError } from "./types";
import type { OpenLocalDocumentOptions, OpenedDocument } from "./types";

export async function openLocalDocument(
  file: File,
  options: OpenLocalDocumentOptions = {},
): Promise<OpenedDocument> {
  const signature = await getSignature(file);
  const earlyDetection = detectFileType(file, signature);
  const sizeValidation = validateFileSize(file.size, earlyDetection.kind);

  if (sizeValidation.status === "block") {
    assertFileSize(file.size, earlyDetection.kind);
  }

  if (sizeValidation.status === "warn" && !options.allowLargeFile) {
    throw new DocumentError({
      code: "file-too-large",
      title: sizeValidation.title ?? "Large local file",
      message: sizeValidation.message ?? "This file is large enough to require confirmation before local parsing.",
      recoverable: true,
    });
  }

  const arrayBuffer = await readLocalFile(file, {
    signal: options.signal,
    onProgress: options.onProgress,
  });
  const detection = detectFileType(file, signature, arrayBuffer);

  if (detection.legacyKind) {
    throw new DocumentError({
      code: "legacy-format",
      title: "Older Office format is not supported",
      message:
        detection.message ??
        "This older Office format is not supported yet. Please save the file as DOCX, XLSX, or PPTX and try again.",
    });
  }

  if (!detection.kind) {
    throw new DocumentError({
      code: "unsupported-format",
      title: "Unsupported file",
      message: detection.message ?? "Open a PDF, DOCX, XLSX, PPTX, or CSV file.",
    });
  }

  return openWithRegisteredAdapter({
    file,
    arrayBuffer,
    signature,
    detection,
    signal: options.signal,
  });
}

import type { FileTypeDetection, OpenedDocumentBase, ViewerCapability } from "./types";

export function createOpenedDocumentBase(options: {
  kind: OpenedDocumentBase["kind"];
  file: File;
  detection: FileTypeDetection;
  capabilities: ViewerCapability[];
  warnings?: string[];
  dispose?: () => void;
}): OpenedDocumentBase {
  const originalUrl = URL.createObjectURL(options.file);
  let disposed = false;

  return {
    kind: options.kind,
    fileName: options.file.name,
    fileSize: options.file.size,
    fileType: options.file.type || "unknown",
    originalFile: options.file,
    originalUrl,
    capabilities: new Set(options.capabilities),
    warnings: [...options.detection.warnings, ...(options.warnings ?? [])],
    detectedBy: options.detection.confidence,
    openedAt: new Date(),
    dispose() {
      if (disposed) {
        return;
      }

      disposed = true;
      URL.revokeObjectURL(originalUrl);
      options.dispose?.();
    },
  };
}

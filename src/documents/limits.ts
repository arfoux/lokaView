import { DocumentError } from "./types";
import type { DocumentKind } from "./types";

export const FILE_LIMITS = {
  largeWarningBytes: 25 * 1024 * 1024,
  hardBlockBytes: 120 * 1024 * 1024,
  signatureBytes: 4096,
  csvPreviewRows: 2000,
  csvPreviewColumns: 80,
  sheetPreviewRows: 1000,
  sheetPreviewColumns: 80,
  zip: {
    maxEntries: 4500,
    maxSingleEntryBytes: 80 * 1024 * 1024,
    maxTotalUncompressedBytes: 220 * 1024 * 1024,
  },
} as const;

export interface FileSizeValidation {
  status: "ok" | "warn" | "block";
  title?: string;
  message?: string;
}

export function formatBytes(bytes: number): string {
  if (bytes === 0) {
    return "0 B";
  }

  const units = ["B", "KB", "MB", "GB"];
  const unitIndex = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / 1024 ** unitIndex;
  return `${value >= 10 || unitIndex === 0 ? value.toFixed(0) : value.toFixed(1)} ${units[unitIndex]}`;
}

export function validateFileSize(size: number, kind?: DocumentKind): FileSizeValidation {
  const label = kind ? `${kind.toUpperCase()} file` : "file";

  if (size > FILE_LIMITS.hardBlockBytes) {
    return {
      status: "block",
      title: "File is too large for this local preview",
      message: `This ${label} is ${formatBytes(size)}. To keep the browser responsive, Local Office Viewer blocks files larger than ${formatBytes(
        FILE_LIMITS.hardBlockBytes,
      )}.`,
    };
  }

  if (size > FILE_LIMITS.largeWarningBytes) {
    return {
      status: "warn",
      title: "Large local file",
      message: `This ${label} is ${formatBytes(size)}. It can still be opened locally, but parsing may take a while or use significant memory.`,
    };
  }

  return { status: "ok" };
}

export function assertFileSize(size: number, kind?: DocumentKind): void {
  const validation = validateFileSize(size, kind);

  if (validation.status === "block") {
    throw new DocumentError({
      code: "file-too-large",
      title: validation.title ?? "File is too large",
      message: validation.message ?? "This file is too large for safe local processing.",
    });
  }
}

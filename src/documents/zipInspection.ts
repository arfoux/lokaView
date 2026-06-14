import { FILE_LIMITS, formatBytes } from "./limits";
import { DocumentError } from "./types";

export interface ZipEntrySummary {
  name: string;
  compressedSize: number;
  uncompressedSize: number;
}

export interface ZipInspection {
  isZip: boolean;
  entries: ZipEntrySummary[];
  totalUncompressedSize: number;
  warnings: string[];
}

const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50;

function readUInt16(bytes: Uint8Array, offset: number): number {
  return (bytes[offset] ?? 0) | ((bytes[offset + 1] ?? 0) << 8);
}

function readUInt32(bytes: Uint8Array, offset: number): number {
  return (
    (bytes[offset] ?? 0) |
    ((bytes[offset + 1] ?? 0) << 8) |
    ((bytes[offset + 2] ?? 0) << 16) |
    ((bytes[offset + 3] ?? 0) << 24)
  ) >>> 0;
}

export function hasZipSignature(signature: Uint8Array): boolean {
  return signature[0] === 0x50 && signature[1] === 0x4b;
}

export function inspectZip(arrayBuffer: ArrayBuffer): ZipInspection {
  const bytes = new Uint8Array(arrayBuffer);
  const warnings: string[] = [];

  if (!hasZipSignature(bytes)) {
    return { isZip: false, entries: [], totalUncompressedSize: 0, warnings };
  }

  const minEocdOffset = Math.max(0, bytes.length - 22 - 65535);
  let eocdOffset = -1;

  for (let offset = bytes.length - 22; offset >= minEocdOffset; offset -= 1) {
    if (readUInt32(bytes, offset) === EOCD_SIGNATURE) {
      eocdOffset = offset;
      break;
    }
  }

  if (eocdOffset === -1) {
    warnings.push("The ZIP directory could not be inspected; the file may be malformed or use ZIP64.");
    return { isZip: true, entries: [], totalUncompressedSize: 0, warnings };
  }

  const entryCount = readUInt16(bytes, eocdOffset + 10);
  const centralDirectoryOffset = readUInt32(bytes, eocdOffset + 16);
  const decoder = new TextDecoder("utf-8", { fatal: false });
  const entries: ZipEntrySummary[] = [];
  let offset = centralDirectoryOffset;
  let totalUncompressedSize = 0;

  for (let index = 0; index < entryCount; index += 1) {
    if (offset + 46 > bytes.length || readUInt32(bytes, offset) !== CENTRAL_DIRECTORY_SIGNATURE) {
      warnings.push("The ZIP directory ended unexpectedly; only a partial safety inspection was possible.");
      break;
    }

    const compressedSize = readUInt32(bytes, offset + 20);
    const uncompressedSize = readUInt32(bytes, offset + 24);
    const nameLength = readUInt16(bytes, offset + 28);
    const extraLength = readUInt16(bytes, offset + 30);
    const commentLength = readUInt16(bytes, offset + 32);
    const nameStart = offset + 46;
    const nameEnd = nameStart + nameLength;

    if (nameEnd > bytes.length) {
      warnings.push("A ZIP entry name was truncated; the file may be malformed.");
      break;
    }

    const name = decoder.decode(bytes.subarray(nameStart, nameEnd));
    entries.push({ name, compressedSize, uncompressedSize });
    totalUncompressedSize += uncompressedSize;
    offset = nameEnd + extraLength + commentLength;
  }

  return { isZip: true, entries, totalUncompressedSize, warnings };
}

export function assertZipWithinLimits(inspection: ZipInspection): void {
  if (!inspection.isZip) {
    return;
  }

  if (inspection.entries.length > FILE_LIMITS.zip.maxEntries) {
    throw new DocumentError({
      code: "unsafe-zip",
      title: "Document package is too complex",
      message: `This Office file contains ${inspection.entries.length} ZIP entries. The local safety limit is ${FILE_LIMITS.zip.maxEntries}.`,
    });
  }

  const largeEntry = inspection.entries.find(
    (entry) => entry.uncompressedSize > FILE_LIMITS.zip.maxSingleEntryBytes,
  );

  if (largeEntry) {
    throw new DocumentError({
      code: "unsafe-zip",
      title: "Document package contains an oversized part",
      message: `The entry "${largeEntry.name}" expands to ${formatBytes(
        largeEntry.uncompressedSize,
      )}, which is larger than the local safety limit.`,
    });
  }

  if (inspection.totalUncompressedSize > FILE_LIMITS.zip.maxTotalUncompressedBytes) {
    throw new DocumentError({
      code: "unsafe-zip",
      title: "Document package expands too much",
      message: `This Office package expands to about ${formatBytes(
        inspection.totalUncompressedSize,
      )}. The local safety limit is ${formatBytes(FILE_LIMITS.zip.maxTotalUncompressedBytes)}.`,
    });
  }
}

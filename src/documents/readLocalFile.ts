import { FILE_LIMITS } from "./limits";
import { DocumentError } from "./types";
import type { ReadProgress } from "./types";

export function getSignature(file: File, bytes = FILE_LIMITS.signatureBytes): Promise<Uint8Array> {
  return file.slice(0, bytes).arrayBuffer().then((buffer) => new Uint8Array(buffer));
}

export async function readLocalFile(
  file: File,
  options: {
    signal?: AbortSignal;
    onProgress?: (progress: ReadProgress) => void;
  } = {},
): Promise<ArrayBuffer> {
  const { signal, onProgress } = options;

  if (signal?.aborted) {
    throw new DocumentError({
      code: "read-aborted",
      title: "Opening was canceled",
      message: "The previous document load was canceled.",
    });
  }

  if (!file.stream) {
    return file.arrayBuffer();
  }

  const reader = file.stream().getReader();
  const chunks: Uint8Array[] = [];
  let loadedBytes = 0;

  try {
    while (true) {
      if (signal?.aborted) {
        await reader.cancel();
        throw new DocumentError({
          code: "read-aborted",
          title: "Opening was canceled",
          message: "The previous document load was canceled.",
        });
      }

      const result = await reader.read();

      if (result.done) {
        break;
      }

      chunks.push(result.value);
      loadedBytes += result.value.byteLength;
      onProgress?.({
        loadedBytes,
        totalBytes: file.size,
        ratio: file.size > 0 ? loadedBytes / file.size : 1,
      });
    }
  } catch (error) {
    if (error instanceof DocumentError) {
      throw error;
    }

    throw new DocumentError({
      code: "read-failed",
      title: "Could not read this file",
      message: "The browser could not read the selected local file.",
      cause: error,
    });
  } finally {
    reader.releaseLock();
  }

  const output = new Uint8Array(loadedBytes);
  let offset = 0;
  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.byteLength;
  }

  return output.buffer;
}

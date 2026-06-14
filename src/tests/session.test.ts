import { describe, expect, it, vi } from "vitest";
import { DocumentSession } from "../documents/session";
import type { OpenedDocument } from "../documents/types";

function makeDocument(kind: OpenedDocument["kind"], dispose = vi.fn()): OpenedDocument {
  return {
    kind,
    fileName: `${kind}-file`,
    fileSize: 12,
    fileType: "test",
    originalFile: new File(["x"], `${kind}.bin`),
    originalUrl: "blob:test",
    capabilities: new Set(),
    warnings: [],
    detectedBy: "extension",
    openedAt: new Date(),
    dispose,
  } as unknown as OpenedDocument;
}

describe("DocumentSession", () => {
  it("disposes the previous document when switching", () => {
    const firstDispose = vi.fn();
    const secondDispose = vi.fn();
    const session = new DocumentSession();

    session.setActive(makeDocument("pdf", firstDispose));
    session.setActive(makeDocument("csv", secondDispose));

    expect(firstDispose).toHaveBeenCalledTimes(1);
    expect(secondDispose).not.toHaveBeenCalled();

    session.clear();
    expect(secondDispose).toHaveBeenCalledTimes(1);
  });
});

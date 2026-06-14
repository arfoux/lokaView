import { describe, expect, it } from "vitest";
import { detectFileType } from "../documents/detectFileType";
import { getSignature } from "../documents/readLocalFile";
import { makeFile, makeLegacyOfficeFile, makeOfficeZipFile, makePdfFile } from "./fixtures";

describe("detectFileType", () => {
  it("detects PDF by signature even when the extension is wrong", async () => {
    const file = makePdfFile("document.bin");
    const detection = detectFileType(file, await getSignature(file));

    expect(detection.kind).toBe("pdf");
    expect(detection.confidence).toBe("signature");
  });

  it("detects DOCX from the ZIP manifest", async () => {
    const file = makeOfficeZipFile(
      "letter.docx",
      ["[Content_Types].xml", "word/document.xml"],
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
    const signature = await getSignature(file);
    const detection = detectFileType(file, signature, await file.arrayBuffer());

    expect(detection.kind).toBe("docx");
    expect(detection.confidence).toBe("zip-manifest");
  });

  it("detects older Office binary files as intentionally unsupported", async () => {
    const file = makeLegacyOfficeFile("old.xls");
    const detection = detectFileType(file, await getSignature(file));

    expect(detection.legacyKind).toBe("xls");
    expect(detection.kind).toBeUndefined();
    expect(detection.message).toContain("older Office format");
  });

  it("returns a friendly unsupported result for unknown files", async () => {
    const file = makeFile("notes.txt", ["hello"], "text/plain");
    const detection = detectFileType(file, await getSignature(file));

    expect(detection.kind).toBeUndefined();
    expect(detection.confidence).toBe("unknown");
    expect(detection.message).toContain("not supported");
  });
});

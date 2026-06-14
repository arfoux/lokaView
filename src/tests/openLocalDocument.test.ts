import { describe, expect, it, vi } from "vitest";
import { openLocalDocument } from "../documents/openLocalDocument";
import { DocumentError } from "../documents/types";
import { makeFile, makeMinimalXlsxFile, makeOfficeZipFile, makePdfFile } from "./fixtures";

describe("openLocalDocument", () => {
  it("opens each supported adapter without calling fetch", async () => {
    const { DOMParser } = await import("@xmldom/xmldom");
    vi.stubGlobal("DOMParser", DOMParser);
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    const files = [
      makePdfFile(),
      makeOfficeZipFile(
        "letter.docx",
        ["[Content_Types].xml", "word/document.xml"],
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      ),
      makeMinimalXlsxFile(),
      makeOfficeZipFile(
        "deck.pptx",
        ["[Content_Types].xml", "ppt/presentation.xml", "ppt/slides/slide1.xml"],
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      ),
      makeFile("data.csv", ["name,count\nAlpha,3\nBeta,7"], "text/csv"),
    ];

    const opened = await Promise.all(files.map((file) => openLocalDocument(file)));

    expect(opened.map((document) => document.kind)).toEqual(["pdf", "docx", "xlsx", "pptx", "csv"]);
    const workbook = opened.find((document) => document.kind === "xlsx");
    expect(workbook?.kind === "xlsx" ? workbook.sheets[0]?.styles["0:0"]?.backgroundColor : undefined).toBe("#1f4e79");
    expect(fetchSpy).not.toHaveBeenCalled();
    opened.forEach((document) => document.dispose());
    fetchSpy.mockRestore();
    vi.unstubAllGlobals();
  });

  it("fails gracefully for unsupported files", async () => {
    await expect(openLocalDocument(makeFile("notes.txt", ["hello"], "text/plain"))).rejects.toMatchObject({
      code: "unsupported-format",
    });
  });

  it("fails gracefully for malformed Office packages", async () => {
    await expect(
      openLocalDocument(makeFile("broken.docx", ["not a zip"], "application/vnd.openxmlformats-officedocument.wordprocessingml.document")),
    ).rejects.toBeInstanceOf(DocumentError);
  });
});

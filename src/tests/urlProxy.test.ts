import { describe, expect, it } from "vitest";
import {
  githubBlobUrlToContentsApiUrl,
  isGitHubBlobSourceAllowed,
  resolveDocumentContentType,
  resolveUrlDocument,
} from "../worker/urlProxy";

describe("urlProxy", () => {
  it("resolves the fileku alias without allowing user-provided repo details", () => {
    const resolved = resolveUrlDocument("/url/fileku/datamahasiswa.docx");

    expect(resolved).toMatchObject({
      sourceKind: "alias",
      alias: "fileku",
      relativePath: "datamahasiswa.docx",
      sourceUrl: "https://github.com/arfoux/simpenan/blob/main/datamahasiswa.docx",
      fileName: "datamahasiswa.docx",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
  });

  it("resolves full HTTP(S) document URLs from the /url route", () => {
    const resolved = resolveUrlDocument("/url/https://calibre-ebook.com/downloads/demos/demo.docx");

    expect(resolved).toMatchObject({
      sourceKind: "remote-url",
      alias: "",
      sourceUrl: "https://calibre-ebook.com/downloads/demos/demo.docx",
      fileName: "demo.docx",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
  });

  it("keeps query strings on full document URLs", () => {
    const resolved = resolveUrlDocument("/url/https://example.com/files/report.csv", undefined, "?download=1");

    expect(resolved.sourceUrl).toBe("https://example.com/files/report.csv?download=1");
  });

  it("converts GitHub blob URLs to GitHub Contents API URLs", () => {
    expect(githubBlobUrlToContentsApiUrl("https://github.com/arfoux/simpenan/blob/main/folder/data mahasiswa.xlsx")).toBe(
      "https://api.github.com/repos/arfoux/simpenan/contents/folder/data%20mahasiswa.xlsx?ref=main",
    );
  });

  it("rejects parent directory paths", () => {
    expect(() => resolveUrlDocument("/url/fileku/../secret.docx")).toThrow("parent directory");
    expect(() => resolveUrlDocument("/url/fileku/%2E%2E/secret.docx")).toThrow("parent directory");
  });

  it("rejects unknown aliases", () => {
    expect(() => resolveUrlDocument("/url/arfoux/simpenan/blob/main/data.docx")).toThrow("Unknown document URL alias");
  });

  it("rejects remote URLs for blocked local hosts", () => {
    expect(() => resolveUrlDocument("/url/http://localhost/demo.docx")).toThrow("host is not allowed");
    expect(() => resolveUrlDocument("/url/http://127.0.0.1/demo.docx")).toThrow("host is not allowed");
    expect(() => resolveUrlDocument("/url/http://192.168.1.10/demo.docx")).toThrow("host is not allowed");
  });

  it("only allows GitHub blob token access for allowlisted alias repositories", () => {
    expect(isGitHubBlobSourceAllowed("https://github.com/arfoux/simpenan/blob/main/datamahasiswa.docx")).toBe(true);
    expect(isGitHubBlobSourceAllowed("https://github.com/other/private/blob/main/secret.docx")).toBe(false);
  });

  it("resolves supported content types", () => {
    expect(resolveDocumentContentType("file.pdf")).toBe("application/pdf");
    expect(resolveDocumentContentType("file.docx")).toBe("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    expect(resolveDocumentContentType("file.xlsx")).toBe("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    expect(resolveDocumentContentType("file.pptx")).toBe("application/vnd.openxmlformats-officedocument.presentationml.presentation");
    expect(resolveDocumentContentType("file.csv")).toBe("text/csv; charset=utf-8");
    expect(resolveDocumentContentType("file.bin")).toBe("application/octet-stream");
  });
});

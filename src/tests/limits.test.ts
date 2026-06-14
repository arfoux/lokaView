import { describe, expect, it } from "vitest";
import { FILE_LIMITS, validateFileSize } from "../documents/limits";
import { assertZipWithinLimits } from "../documents/zipInspection";

describe("limits", () => {
  it("warns for large files and blocks unreasonable files", () => {
    expect(validateFileSize(FILE_LIMITS.largeWarningBytes + 1).status).toBe("warn");
    expect(validateFileSize(FILE_LIMITS.hardBlockBytes + 1).status).toBe("block");
  });

  it("blocks ZIP packages that expand beyond local safety limits", () => {
    expect(() =>
      assertZipWithinLimits({
        isZip: true,
        entries: [
          {
            name: "xl/worksheets/sheet1.xml",
            compressedSize: 10,
            uncompressedSize: FILE_LIMITS.zip.maxSingleEntryBytes + 1,
          },
        ],
        totalUncompressedSize: FILE_LIMITS.zip.maxSingleEntryBytes + 1,
        warnings: [],
      }),
    ).toThrow(/larger than the local safety limit/);
  });
});

import { DOMParser } from "@xmldom/xmldom";
import { strToU8, zipSync } from "fflate";
import { afterEach, describe, expect, it, vi } from "vitest";
import { extractWorkbookStyles } from "../documents/xlsx/xlsxStyles";

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

function makeWorkbookWithMerges(): ArrayBuffer {
  return toArrayBuffer(zipSync({
    "xl/workbook.xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>Title</v></c></row>
    <row r="2"><c r="B2"><v>Block</v></c></row>
  </sheetData>
  <mergeCells count="2">
    <mergeCell ref="A1:D1"/>
    <mergeCell ref="B2:C4"/>
  </mergeCells>
</worksheet>`),
  }));
}

describe("xlsxStyles", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("extracts worksheet merge ranges", () => {
    vi.stubGlobal("DOMParser", DOMParser);

    const styles = extractWorkbookStyles(makeWorkbookWithMerges());

    expect(styles.Sheet1?.merges).toEqual([
      { startRow: 0, startColumn: 0, rowSpan: 1, columnSpan: 4 },
      { startRow: 1, startColumn: 1, rowSpan: 3, columnSpan: 2 },
    ]);
  });
});

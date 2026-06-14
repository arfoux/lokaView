function u16(value: number): number[] {
  return [value & 0xff, (value >> 8) & 0xff];
}

function u32(value: number): number[] {
  return [value & 0xff, (value >> 8) & 0xff, (value >> 16) & 0xff, (value >> 24) & 0xff];
}

function encodeText(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

function asArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

export function makeFile(name: string, bytes: BlobPart[], type = ""): File {
  return new File(bytes, name, { type });
}

export function makeOfficeZipFile(name: string, entries: string[], type: string): File {
  const encoder = new TextEncoder();
  const chunks: Uint8Array[] = [];
  const centralDirectory: Uint8Array[] = [];
  let offset = 0;

  for (const entryName of entries) {
    const nameBytes = encoder.encode(entryName);
    const localHeader = new Uint8Array([
      ...u32(0x04034b50),
      ...u16(20),
      ...u16(0),
      ...u16(0),
      ...u16(0),
      ...u16(0),
      ...u32(0),
      ...u32(0),
      ...u32(0),
      ...u16(nameBytes.length),
      ...u16(0),
    ]);
    const centralHeader = new Uint8Array([
      ...u32(0x02014b50),
      ...u16(20),
      ...u16(20),
      ...u16(0),
      ...u16(0),
      ...u16(0),
      ...u16(0),
      ...u32(0),
      ...u32(0),
      ...u32(0),
      ...u16(nameBytes.length),
      ...u16(0),
      ...u16(0),
      ...u16(0),
      ...u16(0),
      ...u32(0),
      ...u32(offset),
    ]);

    chunks.push(localHeader, nameBytes);
    centralDirectory.push(centralHeader, nameBytes);
    offset += localHeader.byteLength + nameBytes.byteLength;
  }

  const centralDirectoryOffset = offset;
  const centralDirectorySize = centralDirectory.reduce((sum, chunk) => sum + chunk.byteLength, 0);
  const eocd = new Uint8Array([
    ...u32(0x06054b50),
    ...u16(0),
    ...u16(0),
    ...u16(entries.length),
    ...u16(entries.length),
    ...u32(centralDirectorySize),
    ...u32(centralDirectoryOffset),
    ...u16(0),
  ]);

  return makeFile(name, [...chunks, ...centralDirectory, eocd].map(asArrayBuffer), type);
}

export function makePdfFile(name = "sample.pdf"): File {
  return makeFile(name, [asArrayBuffer(encodeText("%PDF-1.7\n1 0 obj\n<<>>\nendobj\n%%EOF"))], "application/pdf");
}

export function makeLegacyOfficeFile(name = "legacy.doc"): File {
  return makeFile(name, [
    asArrayBuffer(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])),
  ]);
}

export function makeMinimalXlsxFile(name = "summary.xlsx"): File {
  const files = {
    "[Content_Types].xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    "xl/workbook.xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`),
    "xl/styles.xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/></patternFill></fill>
  </fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1" applyFont="1"><alignment horizontal="center"/></xf>
  </cellXfs>
</styleSheet>`),
    "xl/worksheets/sheet1.xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr" s="1"><is><t>Name</t></is></c>
      <c r="B1" t="inlineStr" s="1"><is><t>Count</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Alpha</t></is></c>
      <c r="B2"><v>3</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Beta</t></is></c>
      <c r="B3"><v>7</v></c>
    </row>
  </sheetData>
</worksheet>`),
  };
  const zipped = zipSync(files);

  return makeFile(
    name,
    [asArrayBuffer(zipped)],
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
}
import { strToU8, zipSync } from "fflate";

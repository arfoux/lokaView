import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { afterEach, describe, expect, it, vi } from "vitest";
import { getDocxLayoutHints, prepareDocxForPreview } from "../documents/docx/docxLayoutFixes";

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

function makeDocxBuffer(sectionProperties: string): ArrayBuffer {
  const files = {
    "word/document.xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Example</w:t></w:r></w:p>
    <w:sectPr>${sectionProperties}</w:sectPr>
  </w:body>
</w:document>`),
  };

  return toArrayBuffer(zipSync(files));
}

function makePreviewDocxBuffer(sectionProperties: string, header = ""): ArrayBuffer {
  const files = {
    "[Content_Types].xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>`),
    "word/_rels/document.xml.rels": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>`),
    "word/document.xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Example</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:left="2268" w:right="1701" w:top="1440" w:bottom="1440"/>
      ${sectionProperties}
    </w:sectPr>
  </w:body>
</w:document>`),
    "word/header1.xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
  ${header}
</w:hdr>`),
  };

  return toArrayBuffer(zipSync(files));
}

function makePreviewDocxBufferWithBody(body: string): ArrayBuffer {
  const files = {
    "[Content_Types].xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`),
    "word/document.xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${body}
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="1440" w:right="1440" w:top="1440" w:bottom="1440"/></w:sectPr>
  </w:body>
</w:document>`),
  };

  return toArrayBuffer(zipSync(files));
}

function makeLetterheadDocxBuffer(part: "header" | "document"): ArrayBuffer {
  const textBox = `
    <w:txbxContent>
      <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>HIMPUNAN MAHASISWA TEKNOLOGI INFORMASI</w:t></w:r></w:p>
      <w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Sekretariat: Example Street</w:t></w:r></w:p>
    </w:txbxContent>`;
  const pictureDrawing = (id: string, posOffset: number) => `
    <w:drawing>
      <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:positionH relativeFrom="page"><wp:posOffset>${posOffset}</wp:posOffset></wp:positionH>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData>
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:blipFill><a:blip r:embed="${id}"/></pic:blipFill>
              <pic:spPr><a:xfrm><a:off x="0" y="0"/></a:xfrm></pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:anchor>
    </w:drawing>`;
  const textBoxDrawing = `
    <w:drawing>
      <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData>
            <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">${textBox}</wps:wsp>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>`;
  const drawing = `${pictureDrawing("rIdRight", 6000)}${pictureDrawing("rIdLeft", 0)}${textBoxDrawing}`;
  const headerReference = part === "header" ? `<w:headerReference w:type="default" r:id="rIdHeader"/>` : "";
  const headerPart: Record<string, Uint8Array> = part === "header"
    ? {
        "word/header1.xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${drawing}</w:hdr>`),
        "word/_rels/header1.xml.rels": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLeft" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/left.png"/>
  <Relationship Id="rIdRight" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/right.png"/>
</Relationships>`),
      }
    : {};
  const documentDrawing = part === "document" ? drawing : "";

  return toArrayBuffer(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>`),
    "word/_rels/document.xml.rels": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rIdLeft" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/left.png"/>
  <Relationship Id="rIdRight" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/right.png"/>
</Relationships>`),
    "word/document.xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r>${documentDrawing}<w:t>Body</w:t></w:r></w:p>
    <w:sectPr>${headerReference}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="1440" w:right="1440" w:top="1440" w:bottom="1440"/></w:sectPr>
  </w:body>
</w:document>`),
    "word/media/left.png": strToU8("left"),
    "word/media/right.png": strToU8("right"),
    ...headerPart,
  }));
}

describe("docxLayoutFixes", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("suppresses a title page default header when no first header exists", () => {
    vi.stubGlobal("DOMParser", DOMParser);

    const hints = getDocxLayoutHints(makeDocxBuffer(`
      <w:headerReference w:type="default" r:id="rId9"/>
      <w:titlePg/>
    `));

    expect(hints.sections[0]).toMatchObject({
      suppressFirstPageHeader: true,
      suppressFirstPageFooter: false,
    });
  });

  it("keeps a title page header when the document defines a first header", () => {
    vi.stubGlobal("DOMParser", DOMParser);

    const hints = getDocxLayoutHints(makeDocxBuffer(`
      <w:headerReference w:type="first" r:id="rId8"/>
      <w:headerReference w:type="default" r:id="rId9"/>
      <w:titlePg/>
    `));

    expect(hints.sections[0]).toMatchObject({
      suppressFirstPageHeader: false,
      suppressFirstPageFooter: false,
    });
  });

  it("returns empty hints for malformed DOCX data", () => {
    vi.stubGlobal("DOMParser", DOMParser);

    expect(getDocxLayoutHints(toArrayBuffer(strToU8("not a zip")))).toEqual({ sections: [] });
  });

  it("adds an empty first page header for title pages that only define a default header", () => {
    vi.stubGlobal("DOMParser", DOMParser);
    vi.stubGlobal("XMLSerializer", XMLSerializer);

    const prepared = prepareDocxForPreview(makePreviewDocxBuffer(`
      <w:headerReference w:type="default" r:id="rId9"/>
      <w:titlePg/>
    `));
    const files = unzipSync(new Uint8Array(prepared));
    const documentXml = strFromU8(files["word/document.xml"]!);

    expect(Object.keys(files).some((path) => path.startsWith("word/header-codex-first-"))).toBe(true);
    expect(documentXml).toContain("w:type=\"first\"");
    expect(documentXml).toContain("r:id=\"rIdCodexFirstHeader1\"");
  });

  it("normalizes right-margin anchored header drawings for docx-preview", () => {
    vi.stubGlobal("DOMParser", DOMParser);
    vi.stubGlobal("XMLSerializer", XMLSerializer);

    const prepared = prepareDocxForPreview(makePreviewDocxBuffer(
      `<w:headerReference w:type="default" r:id="rId9"/>`,
      `<w:p><w:r><w:drawing><wp:anchor>
        <wp:positionH relativeFrom="rightMargin"><wp:posOffset>-635</wp:posOffset></wp:positionH>
      </wp:anchor></w:drawing></w:r></w:p>`,
    ));
    const files = unzipSync(new Uint8Array(prepared));
    const headerXml = strFromU8(files["word/header1.xml"]!);

    expect(headerXml).toContain("relativeFrom=\"column\"");
    expect(headerXml).toContain("<wp:posOffset>5251450</wp:posOffset>");
  });

  it("normalizes left-margin anchored header drawings into the page margin", () => {
    vi.stubGlobal("DOMParser", DOMParser);
    vi.stubGlobal("XMLSerializer", XMLSerializer);

    const prepared = prepareDocxForPreview(makePreviewDocxBuffer(
      `<w:headerReference w:type="default" r:id="rId9"/>`,
      `<w:p><w:r><w:drawing><wp:anchor>
        <wp:positionH relativeFrom="leftMargin"><wp:posOffset>836930</wp:posOffset></wp:positionH>
      </wp:anchor></w:drawing></w:r></w:p>`,
    ));
    const files = unzipSync(new Uint8Array(prepared));
    const headerXml = strFromU8(files["word/header1.xml"]!);

    expect(headerXml).toContain("relativeFrom=\"column\"");
    expect(headerXml).toContain("<wp:posOffset>-603250</wp:posOffset>");
  });

  it("removes stale last rendered page breaks immediately after explicit page breaks", () => {
    vi.stubGlobal("DOMParser", DOMParser);
    vi.stubGlobal("XMLSerializer", XMLSerializer);

    const prepared = prepareDocxForPreview(makePreviewDocxBufferWithBody(`
      <w:p><w:r><w:t>Before</w:t></w:r></w:p>
      <w:p><w:r><w:br w:type="page"/></w:r></w:p>
      <w:p><w:r><w:lastRenderedPageBreak/><w:t>After</w:t></w:r></w:p>
      <w:p><w:r><w:lastRenderedPageBreak/><w:t>Keep later break</w:t></w:r></w:p>
    `));
    const files = unzipSync(new Uint8Array(prepared));
    const documentXml = strFromU8(files["word/document.xml"]!);

    expect(documentXml).not.toContain("<w:lastRenderedPageBreak/><w:t>After</w:t>");
    expect(documentXml).toContain("<w:lastRenderedPageBreak/><w:t>Keep later break</w:t>");
  });

  it("extracts repeatable letterhead fallbacks from header text boxes", () => {
    vi.stubGlobal("DOMParser", DOMParser);

    const hints = getDocxLayoutHints(makeLetterheadDocxBuffer("header"));

    expect(hints.letterhead?.repeatOnPages).toBe(true);
    expect(hints.letterhead?.paragraphs[0]?.runs[0]).toMatchObject({
      text: "HIMPUNAN MAHASISWA TEKNOLOGI INFORMASI",
      bold: true,
    });
    expect(hints.letterhead?.images).toHaveLength(2);
    expect(hints.letterhead?.images[0]?.src).toContain("bGVmdA==");
  });

  it("extracts first-page letterhead fallbacks from body text boxes", () => {
    vi.stubGlobal("DOMParser", DOMParser);

    const hints = getDocxLayoutHints(makeLetterheadDocxBuffer("document"));

    expect(hints.letterhead?.repeatOnPages).toBe(false);
    expect(hints.letterhead?.sourcePart).toBe("word/document.xml");
  });
});

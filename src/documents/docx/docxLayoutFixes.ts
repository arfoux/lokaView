import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

export interface DocxSectionLayoutHints {
  suppressFirstPageHeader: boolean;
  suppressFirstPageFooter: boolean;
}

export interface DocxLayoutHints {
  sections: DocxSectionLayoutHints[];
}

interface DocxPageMetrics {
  contentWidthEmu: number;
  leftMarginEmu: number;
  rightMarginEmu: number;
}

type ZipEntryBytes = Uint8Array<ArrayBufferLike>;
type ZipEntries = Record<string, ZipEntryBytes>;

const EMPTY_HINTS: DocxLayoutHints = {
  sections: [],
};

const DOCX_PAGE_SELECTOR = "section.docx-preview, section.docx";
const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const HEADER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
const HEADER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const EMU_PER_DXA = 635;

function isElement(node: Node): node is Element {
  return node.nodeType === 1;
}

function getLocalName(element: Element): string {
  return element.localName || element.nodeName.split(":").at(-1) || element.nodeName;
}

function getAttributeByLocalName(element: Element, name: string): string | null {
  for (let index = 0; index < element.attributes.length; index += 1) {
    const attribute = element.attributes.item(index);

    if (!attribute) {
      continue;
    }

    const attributeName = attribute.localName || attribute.nodeName.split(":").at(-1) || attribute.nodeName;

    if (attributeName === name) {
      return attribute.value;
    }
  }

  return null;
}

function getElementsByLocalName(root: Document | Element, name: string): Element[] {
  return Array.from(root.getElementsByTagName("*")).filter((element) => getLocalName(element) === name);
}

function getChildElementsByLocalName(parent: Element, name: string): Element[] {
  return Array.from(parent.childNodes)
    .filter(isElement)
    .filter((element) => getLocalName(element) === name);
}

function getDirectChild(parent: Element, tagName: string): HTMLElement | undefined {
  return Array.from(parent.children).find((child): child is HTMLElement => (
    child instanceof HTMLElement && child.tagName.toLowerCase() === tagName
  ));
}

function toArrayBuffer(bytes: ZipEntryBytes): ArrayBuffer {
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);

  return copy.buffer;
}

function parseXml(bytes: ZipEntryBytes): Document {
  return new DOMParser().parseFromString(strFromU8(bytes), "application/xml");
}

function serializeXml(document: Document): ZipEntryBytes {
  return strToU8(new XMLSerializer().serializeToString(document));
}

function getNumericAttribute(element: Element | undefined, name: string): number | undefined {
  if (!element) {
    return undefined;
  }

  const value = getAttributeByLocalName(element, name);
  const parsedValue = value ? Number(value) : Number.NaN;

  return Number.isFinite(parsedValue) ? parsedValue : undefined;
}

function getFirstSectionPageMetrics(sectionProperties: Element | undefined): DocxPageMetrics | undefined {
  if (!sectionProperties) {
    return undefined;
  }

  const pageSize = getChildElementsByLocalName(sectionProperties, "pgSz")[0];
  const pageMargins = getChildElementsByLocalName(sectionProperties, "pgMar")[0];
  const width = getNumericAttribute(pageSize, "w");
  const leftMargin = getNumericAttribute(pageMargins, "left") ?? 0;
  const rightMargin = getNumericAttribute(pageMargins, "right") ?? 0;

  if (!width) {
    return undefined;
  }

  return {
    contentWidthEmu: Math.max(0, width - leftMargin - rightMargin) * EMU_PER_DXA,
    leftMarginEmu: leftMargin * EMU_PER_DXA,
    rightMarginEmu: rightMargin * EMU_PER_DXA,
  };
}

function parseSectionLayoutHints(sectionProperties: Element): DocxSectionLayoutHints {
  const hasTitlePage = getChildElementsByLocalName(sectionProperties, "titlePg").length > 0;
  const headerReferences = getChildElementsByLocalName(sectionProperties, "headerReference");
  const footerReferences = getChildElementsByLocalName(sectionProperties, "footerReference");
  const hasFirstHeader = headerReferences.some((reference) => getAttributeByLocalName(reference, "type") === "first");
  const hasFirstFooter = footerReferences.some((reference) => getAttributeByLocalName(reference, "type") === "first");

  return {
    suppressFirstPageHeader: hasTitlePage && headerReferences.length > 0 && !hasFirstHeader,
    suppressFirstPageFooter: hasTitlePage && footerReferences.length > 0 && !hasFirstFooter,
  };
}

function getUniqueRelationshipId(relationships: Document): string {
  const existingIds = new Set(getElementsByLocalName(relationships, "Relationship")
    .map((relationship) => getAttributeByLocalName(relationship, "Id"))
    .filter((id): id is string => Boolean(id)));
  let index = 1;

  while (existingIds.has(`rIdCodexFirstHeader${index}`)) {
    index += 1;
  }

  return `rIdCodexFirstHeader${index}`;
}

function appendContentTypeOverride(contentTypes: Document, partName: string): void {
  const root = contentTypes.documentElement;
  const alreadyExists = getElementsByLocalName(contentTypes, "Override").some((override) => (
    getAttributeByLocalName(override, "PartName") === partName
  ));

  if (alreadyExists) {
    return;
  }

  const override = contentTypes.createElementNS(root.namespaceURI, "Override");
  override.setAttribute("PartName", partName);
  override.setAttribute("ContentType", HEADER_CONTENT_TYPE);
  root.appendChild(override);
}

function appendHeaderRelationship(relationships: Document, id: string, target: string): void {
  const root = relationships.documentElement;
  const relationship = relationships.createElementNS(root.namespaceURI || PACKAGE_REL_NS, "Relationship");
  relationship.setAttribute("Id", id);
  relationship.setAttribute("Type", HEADER_REL_TYPE);
  relationship.setAttribute("Target", target);
  root.appendChild(relationship);
}

function insertFirstHeaderReference(documentXml: Document, sectionProperties: Element, id: string): void {
  const headerReference = documentXml.createElementNS(WORD_NS, "w:headerReference");
  headerReference.setAttributeNS(WORD_NS, "w:type", "first");
  headerReference.setAttributeNS(REL_NS, "r:id", id);

  const firstHeaderReference = getChildElementsByLocalName(sectionProperties, "headerReference")[0];
  sectionProperties.insertBefore(headerReference, firstHeaderReference ?? sectionProperties.firstChild);
}

function addEmptyFirstPageHeader(files: ZipEntries, documentXml: Document, firstSection: Element): boolean {
  const relationshipsXml = files["word/_rels/document.xml.rels"];
  const contentTypesXml = files["[Content_Types].xml"];

  if (!relationshipsXml || !contentTypesXml) {
    return false;
  }

  const relationships = parseXml(relationshipsXml);
  const contentTypes = parseXml(contentTypesXml);
  const relationshipId = getUniqueRelationshipId(relationships);
  const headerFileName = `header-codex-first-${relationshipId.replace(/\D/g, "") || "1"}.xml`;
  const headerPath = `word/${headerFileName}`;

  appendHeaderRelationship(relationships, relationshipId, headerFileName);
  appendContentTypeOverride(contentTypes, `/${headerPath}`);
  insertFirstHeaderReference(documentXml, firstSection, relationshipId);

  files[headerPath] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="${WORD_NS}"/>`);
  files["word/_rels/document.xml.rels"] = serializeXml(relationships);
  files["[Content_Types].xml"] = serializeXml(contentTypes);

  return true;
}

function ensureEmptyFirstHeaderForTitlePage(files: ZipEntries, documentXml: Document, firstSection: Element | undefined): boolean {
  if (!firstSection) {
    return false;
  }

  const hints = parseSectionLayoutHints(firstSection);

  if (!hints.suppressFirstPageHeader) {
    return false;
  }

  return addEmptyFirstPageHeader(files, documentXml, firstSection);
}

function normalizeHorizontalAnchors(xml: Document, metrics: DocxPageMetrics | undefined): boolean {
  if (!metrics) {
    return false;
  }

  let changed = false;

  getElementsByLocalName(xml, "positionH").forEach((position) => {
    const relativeFrom = getAttributeByLocalName(position, "relativeFrom");

    const offsetElement = getChildElementsByLocalName(position, "posOffset")[0];
    const offsetValue = offsetElement?.textContent ? Number(offsetElement.textContent) : Number.NaN;

    if (!offsetElement || !Number.isFinite(offsetValue)) {
      return;
    }

    if (relativeFrom === "leftMargin") {
      position.setAttribute("relativeFrom", "column");
      offsetElement.textContent = String(Math.round(offsetValue - metrics.leftMarginEmu));
      changed = true;
    }

    if (relativeFrom === "rightMargin") {
      position.setAttribute("relativeFrom", "column");
      offsetElement.textContent = String(Math.round(metrics.contentWidthEmu + offsetValue));
      changed = true;
    }
  });

  return changed;
}

function hasExplicitPageBreak(paragraph: Element): boolean {
  return getElementsByLocalName(paragraph, "br").some((breakElement) => (
    getAttributeByLocalName(breakElement, "type") === "page"
  ));
}

function removeRedundantLastRenderedPageBreaks(xml: Document): boolean {
  let changed = false;
  let lastExplicitBreakParagraph = -1;

  getElementsByLocalName(xml, "p").forEach((paragraph, index) => {
    const lastRenderedPageBreaks = getElementsByLocalName(paragraph, "lastRenderedPageBreak");

    if (lastRenderedPageBreaks.length > 0 && lastExplicitBreakParagraph >= 0 && index - lastExplicitBreakParagraph <= 1) {
      lastRenderedPageBreaks.forEach((breakElement) => breakElement.parentNode?.removeChild(breakElement));
      changed = true;
    }

    if (hasExplicitPageBreak(paragraph)) {
      lastExplicitBreakParagraph = index;
    }
  });

  return changed;
}

function keepWideTablesInsidePages(root: HTMLElement): void {
  root.querySelectorAll<HTMLElement>(DOCX_PAGE_SELECTOR).forEach((page) => {
    const pageRect = page.getBoundingClientRect();

    page.querySelectorAll<HTMLElement>("article table").forEach((table) => {
      table.style.transform = "";
      table.style.transformOrigin = "top left";

      const tableRect = table.getBoundingClientRect();
      const overflowRight = tableRect.right - pageRect.right;
      const overflowLeft = pageRect.left - tableRect.left;
      let shift = overflowRight > 0 ? -overflowRight : 0;

      if (tableRect.left + shift < pageRect.left) {
        shift += pageRect.left - (tableRect.left + shift);
      }

      if (overflowLeft > 0 && tableRect.right + shift <= pageRect.right) {
        shift += overflowLeft;
      }

      if (Math.abs(shift) >= 1) {
        table.style.transform = `translateX(${Math.round(shift)}px)`;
      }
    });
  });
}

function centerWideHeaderSvg(header: HTMLElement, page: HTMLElement): void {
  const pageRect = page.getBoundingClientRect();
  const headerRect = header.getBoundingClientRect();

  header.querySelectorAll<SVGSVGElement>("svg").forEach((svg) => {
    const svgStyle = window.getComputedStyle(svg);
    const svgRect = svg.getBoundingClientRect();

    if (svgStyle.position !== "absolute" || svgRect.width < headerRect.width * 0.8) {
      return;
    }

    const currentLeft = Number.parseFloat(svgStyle.left);
    const targetLeft = pageRect.left + ((pageRect.width - svgRect.width) / 2);

    if (!Number.isFinite(currentLeft)) {
      return;
    }

    svg.style.left = `${currentLeft + targetLeft - svgRect.left}px`;
  });
}

function keepHeaderFloatingElementsInsidePages(root: HTMLElement): void {
  root.querySelectorAll<HTMLElement>(DOCX_PAGE_SELECTOR).forEach((page) => {
    const pageRect = page.getBoundingClientRect();

    page.querySelectorAll<HTMLElement>("header").forEach((header) => {
      centerWideHeaderSvg(header, page);

      header.querySelectorAll<HTMLElement>("img, svg").forEach((element) => {
        const style = window.getComputedStyle(element);

        if (style.position !== "absolute") {
          return;
        }

        const rect = element.getBoundingClientRect();
        const overflowRight = rect.right - pageRect.right;
        const overflowLeft = pageRect.left - rect.left;
        const currentLeft = Number.parseFloat(style.left);

        if (!Number.isFinite(currentLeft)) {
          return;
        }

        if (overflowRight > 1) {
          element.style.left = `${currentLeft - overflowRight}px`;
        }

        if (overflowLeft > 1) {
          element.style.left = `${currentLeft + overflowLeft}px`;
        }
      });
    });
  });
}

export function prepareDocxForPreview(arrayBuffer: ArrayBuffer): ArrayBuffer {
  if (typeof DOMParser === "undefined" || typeof XMLSerializer === "undefined") {
    return arrayBuffer;
  }

  try {
    const files: ZipEntries = unzipSync(new Uint8Array(arrayBuffer));
    const documentXmlBytes = files["word/document.xml"];

    if (!documentXmlBytes) {
      return arrayBuffer;
    }

    const documentXml = parseXml(documentXmlBytes);
    const firstSection = getElementsByLocalName(documentXml, "sectPr")[0];
    const pageMetrics = getFirstSectionPageMetrics(firstSection);
    let changed = ensureEmptyFirstHeaderForTitlePage(files, documentXml, firstSection);

    if (normalizeHorizontalAnchors(documentXml, pageMetrics)) {
      changed = true;
    }

    if (removeRedundantLastRenderedPageBreaks(documentXml)) {
      changed = true;
    }

    files["word/document.xml"] = serializeXml(documentXml);

    Object.keys(files).forEach((path) => {
      if (!/^word\/(?:header|footer)\d+\.xml$/i.test(path)) {
        return;
      }

      const fileBytes = files[path];

      if (!fileBytes) {
        return;
      }

      const xml = parseXml(fileBytes);

      if (normalizeHorizontalAnchors(xml, pageMetrics)) {
        files[path] = serializeXml(xml);
        changed = true;
      }
    });

    return changed ? toArrayBuffer(zipSync(files)) : arrayBuffer;
  } catch {
    return arrayBuffer;
  }
}

export function getDocxLayoutHints(arrayBuffer: ArrayBuffer): DocxLayoutHints {
  if (typeof DOMParser === "undefined") {
    return EMPTY_HINTS;
  }

  try {
    const files = unzipSync(new Uint8Array(arrayBuffer));
    const documentXml = files["word/document.xml"];

    if (!documentXml) {
      return EMPTY_HINTS;
    }

    const xml = parseXml(documentXml);
    const sectionProperties = getElementsByLocalName(xml, "sectPr");

    return {
      sections: sectionProperties.map(parseSectionLayoutHints),
    };
  } catch {
    return EMPTY_HINTS;
  }
}

export function applyDocxPostRenderFixes(root: HTMLElement, hints: DocxLayoutHints): void {
  const firstPage = root.querySelector<HTMLElement>(DOCX_PAGE_SELECTOR);
  const firstSectionHints = hints.sections[0];

  if (firstPage && firstSectionHints?.suppressFirstPageHeader) {
    getDirectChild(firstPage, "header")?.remove();
  }

  if (firstPage && firstSectionHints?.suppressFirstPageFooter) {
    getDirectChild(firstPage, "footer")?.remove();
  }

  root.querySelectorAll<HTMLElement>("table td, table th").forEach((cell) => {
    cell.style.textAlign = "left";
  });

  keepWideTablesInsidePages(root);

  root.querySelectorAll<HTMLElement>("header").forEach((header) => {
    header.style.position = "relative";

    header.querySelectorAll<HTMLElement>("p div").forEach((drawing) => {
      if (!drawing.querySelector("img")) {
        return;
      }

      const style = window.getComputedStyle(drawing);

      if (style.position !== "relative" || style.width !== "0px" || style.height !== "0px") {
        return;
      }

      const paragraph = drawing.closest("p") as HTMLElement | null;
      const left = Number.parseFloat(style.left);
      const top = Number.parseFloat(style.top);

      drawing.style.position = "absolute";
      drawing.style.left = `${Number.isFinite(left) ? left : 0}px`;
      drawing.style.top = `${(paragraph?.offsetTop ?? 0) + (Number.isFinite(top) ? top : 0)}px`;
      drawing.style.zIndex = "0";
    });
  });

  keepHeaderFloatingElementsInsidePages(root);
}

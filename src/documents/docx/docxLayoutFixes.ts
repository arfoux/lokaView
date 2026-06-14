import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

export interface DocxSectionLayoutHints {
  suppressFirstPageHeader: boolean;
  suppressFirstPageFooter: boolean;
}

export interface DocxLetterheadRun {
  text: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  color?: string;
}

export interface DocxLetterheadParagraph {
  runs: DocxLetterheadRun[];
}

export interface DocxLetterheadImage {
  src: string;
  x: number;
}

export interface DocxLetterheadHint {
  sourcePart: string;
  repeatOnPages: boolean;
  paragraphs: DocxLetterheadParagraph[];
  images: DocxLetterheadImage[];
}

export interface DocxLayoutHints {
  sections: DocxSectionLayoutHints[];
  letterhead?: DocxLetterheadHint;
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
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

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

function getPartRelationshipsPath(partPath: string): string {
  const segments = partPath.split("/");
  const fileName = segments.pop() ?? "";
  return [...segments, "_rels", `${fileName}.rels`].join("/");
}

function normalizeZipPath(path: string): string {
  const segments: string[] = [];

  path.split("/").forEach((segment) => {
    if (!segment || segment === ".") {
      return;
    }

    if (segment === "..") {
      segments.pop();
      return;
    }

    segments.push(segment);
  });

  return segments.join("/");
}

function resolveRelationshipTarget(partPath: string, target: string): string {
  if (target.startsWith("/")) {
    return normalizeZipPath(target.slice(1));
  }

  const basePath = partPath.split("/").slice(0, -1).join("/");
  return normalizeZipPath(`${basePath}/${target}`);
}

function getRelationshipTargets(files: ZipEntries, partPath: string): Map<string, string> {
  const relationshipsPath = getPartRelationshipsPath(partPath);
  const relationshipsBytes = files[relationshipsPath];
  const targets = new Map<string, string>();

  if (!relationshipsBytes) {
    return targets;
  }

  const relationships = parseXml(relationshipsBytes);

  getElementsByLocalName(relationships, "Relationship").forEach((relationship) => {
    const id = getAttributeByLocalName(relationship, "Id");
    const type = getAttributeByLocalName(relationship, "Type");
    const target = getAttributeByLocalName(relationship, "Target");

    if (id && target && type === IMAGE_REL_TYPE) {
      targets.set(id, resolveRelationshipTarget(partPath, target));
    }
  });

  return targets;
}

function getImageMimeType(path: string): string {
  const lowerPath = path.toLowerCase();

  if (lowerPath.endsWith(".png")) {
    return "image/png";
  }

  if (lowerPath.endsWith(".jpg") || lowerPath.endsWith(".jpeg")) {
    return "image/jpeg";
  }

  if (lowerPath.endsWith(".gif")) {
    return "image/gif";
  }

  if (lowerPath.endsWith(".webp")) {
    return "image/webp";
  }

  return "application/octet-stream";
}

function bytesToBase64(bytes: ZipEntryBytes): string {
  let binary = "";
  const chunkSize = 0x8000;

  for (let index = 0; index < bytes.length; index += chunkSize) {
    const chunk = bytes.subarray(index, index + chunkSize);
    binary += String.fromCharCode(...chunk);
  }

  return btoa(binary);
}

function bytesToDataUrl(bytes: ZipEntryBytes, path: string): string {
  return `data:${getImageMimeType(path)};base64,${bytesToBase64(bytes)}`;
}

function getTogglePropertyValue(runProperties: Element | undefined, propertyName: string): boolean | undefined {
  if (!runProperties) {
    return undefined;
  }

  const property = getChildElementsByLocalName(runProperties, propertyName)[0];

  if (!property) {
    return undefined;
  }

  const value = getAttributeByLocalName(property, "val");

  return value !== "0" && value !== "false" && value !== "none";
}

function getRunProperties(element: Element): Element | undefined {
  return getChildElementsByLocalName(element, "rPr")[0];
}

function getParagraphRunProperties(paragraph: Element): Element | undefined {
  const paragraphProperties = getChildElementsByLocalName(paragraph, "pPr")[0];
  return paragraphProperties ? getChildElementsByLocalName(paragraphProperties, "rPr")[0] : undefined;
}

function getRunText(run: Element): string {
  return getElementsByLocalName(run, "t").map((text) => text.textContent ?? "").join("");
}

function getTextBoxText(textBox: Element): string {
  return getElementsByLocalName(textBox, "t").map((text) => text.textContent ?? "").join("");
}

function getRunColor(runProperties: Element | undefined): string | undefined {
  const color = runProperties ? getChildElementsByLocalName(runProperties, "color")[0] : undefined;
  const value = color ? getAttributeByLocalName(color, "val") : undefined;

  if (!value || value.toLowerCase() === "auto") {
    return undefined;
  }

  return `#${value}`;
}

function extractLetterheadParagraph(paragraph: Element): DocxLetterheadParagraph | undefined {
  const paragraphRunProperties = getParagraphRunProperties(paragraph);
  const paragraphBold = getTogglePropertyValue(paragraphRunProperties, "b") ?? false;
  const paragraphItalic = getTogglePropertyValue(paragraphRunProperties, "i") ?? false;
  const paragraphUnderline = getTogglePropertyValue(paragraphRunProperties, "u") ?? false;
  const runs: DocxLetterheadRun[] = [];

  const appendRun = (run: Element) => {
    const text = getRunText(run);

    if (!text) {
      return;
    }

    const runProperties = getRunProperties(run);
    runs.push({
      text,
      bold: getTogglePropertyValue(runProperties, "b") ?? paragraphBold,
      italic: getTogglePropertyValue(runProperties, "i") ?? paragraphItalic,
      underline: getTogglePropertyValue(runProperties, "u") ?? paragraphUnderline,
      color: getRunColor(runProperties),
    });
  };

  Array.from(paragraph.childNodes).filter(isElement).forEach((child) => {
    if (getLocalName(child) === "r") {
      appendRun(child);
      return;
    }

    if (getLocalName(child) === "hyperlink") {
      getChildElementsByLocalName(child, "r").forEach(appendRun);
    }
  });

  return runs.some((run) => run.text.trim()) ? { runs } : undefined;
}

function extractLetterheadParagraphs(textBox: Element): DocxLetterheadParagraph[] {
  return getChildElementsByLocalName(textBox, "p")
    .map(extractLetterheadParagraph)
    .filter((paragraph): paragraph is DocxLetterheadParagraph => Boolean(paragraph));
}

function getClosestDrawingContainer(element: Element): Element | undefined {
  let current = element.parentElement;

  while (current) {
    const name = getLocalName(current);

    if (name === "anchor" || name === "inline") {
      return current;
    }

    current = current.parentElement;
  }

  return undefined;
}

function getDrawingAnchorXPosition(picture: Element): number | undefined {
  const container = getClosestDrawingContainer(picture);

  if (!container) {
    return undefined;
  }

  const positionH = getChildElementsByLocalName(container, "positionH")[0];
  const posOffset = positionH ? getChildElementsByLocalName(positionH, "posOffset")[0] : undefined;
  const offset = Number.parseInt(posOffset?.textContent ?? "", 10);

  if (Number.isFinite(offset)) {
    return offset;
  }

  const simplePosition = getChildElementsByLocalName(container, "simplePos")[0];
  const x = Number.parseInt(simplePosition ? getAttributeByLocalName(simplePosition, "x") ?? "" : "", 10);

  return Number.isFinite(x) ? x : undefined;
}

function getPictureXPosition(picture: Element): number {
  const anchorX = getDrawingAnchorXPosition(picture);

  if (anchorX !== undefined) {
    return anchorX;
  }

  const transform = getElementsByLocalName(picture, "xfrm")[0];
  const offset = transform ? getChildElementsByLocalName(transform, "off")[0] : undefined;
  return getNumericAttribute(offset, "x") ?? 0;
}

function extractLetterheadImages(files: ZipEntries, xml: Document, partPath: string): DocxLetterheadImage[] {
  const targets = getRelationshipTargets(files, partPath);
  const images: DocxLetterheadImage[] = [];
  const usedIds = new Set<string>();

  getElementsByLocalName(xml, "pic").forEach((picture) => {
    const blip = getElementsByLocalName(picture, "blip")[0];
    const embedId = blip ? getAttributeByLocalName(blip, "embed") : undefined;
    const target = embedId ? targets.get(embedId) : undefined;
    const bytes = target ? files[target] : undefined;

    if (!embedId || !target || !bytes || usedIds.has(embedId)) {
      return;
    }

    usedIds.add(embedId);
    images.push({
      src: bytesToDataUrl(bytes, target),
      x: getPictureXPosition(picture),
    });
  });

  return images.sort((left, right) => left.x - right.x).slice(0, 2);
}

function extractLetterheadFromPart(files: ZipEntries, partPath: string): DocxLetterheadHint | undefined {
  const bytes = files[partPath];

  if (!bytes) {
    return undefined;
  }

  const xml = parseXml(bytes);
  const textBoxes = getElementsByLocalName(xml, "txbxContent");
  const textBox = textBoxes.find((candidate) => {
    const text = getTextBoxText(candidate).toUpperCase();
    return text.includes("HIMPUNAN") && text.includes("SEKRETARIAT");
  });

  if (!textBox) {
    return undefined;
  }

  const paragraphs = extractLetterheadParagraphs(textBox);

  if (paragraphs.length === 0) {
    return undefined;
  }

  return {
    sourcePart: partPath,
    repeatOnPages: /^word\/header\d+\.xml$/i.test(partPath),
    paragraphs,
    images: extractLetterheadImages(files, xml, partPath),
  };
}

function extractLetterheadHint(files: ZipEntries): DocxLetterheadHint | undefined {
  const candidateParts = [
    ...Object.keys(files).filter((path) => /^word\/header\d+\.xml$/i.test(path)).sort(),
    "word/document.xml",
  ];

  for (const partPath of candidateParts) {
    const hint = extractLetterheadFromPart(files, partPath);

    if (hint) {
      return hint;
    }
  }

  return undefined;
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

function renderLetterheadRun(run: DocxLetterheadRun): HTMLSpanElement {
  const span = window.document.createElement("span");
  span.textContent = run.text;
  span.style.fontFamily = '"Times New Roman", Times, serif';
  span.style.fontWeight = run.bold ? "700" : "400";
  span.style.fontStyle = run.italic ? "italic" : "normal";

  if (run.underline) {
    span.style.textDecorationLine = "underline";
  }

  if (run.color) {
    span.style.color = run.color;
  }

  return span;
}

function renderLetterheadParagraph(paragraph: DocxLetterheadParagraph, index: number): HTMLParagraphElement {
  const element = window.document.createElement("p");
  element.className = index === 0 ? "docx-letterhead-title" : "docx-letterhead-meta";
  paragraph.runs.forEach((run) => element.appendChild(renderLetterheadRun(run)));
  return element;
}

function createLetterheadFallback(letterhead: DocxLetterheadHint): HTMLElement {
  const container = window.document.createElement("div");
  container.className = "docx-letterhead-fallback";
  container.setAttribute("contenteditable", "false");

  const leftImage = letterhead.images[0];
  const rightImage = letterhead.images.at(-1);

  if (leftImage) {
    const image = window.document.createElement("img");
    image.className = "docx-letterhead-logo docx-letterhead-logo-left";
    image.src = leftImage.src;
    image.alt = "";
    container.appendChild(image);
  }

  const text = window.document.createElement("div");
  text.className = "docx-letterhead-text";
  letterhead.paragraphs.forEach((paragraph, index) => text.appendChild(renderLetterheadParagraph(paragraph, index)));
  container.appendChild(text);

  if (rightImage && rightImage !== leftImage) {
    const image = window.document.createElement("img");
    image.className = "docx-letterhead-logo docx-letterhead-logo-right";
    image.src = rightImage.src;
    image.alt = "";
    container.appendChild(image);
  }

  const rule = window.document.createElement("div");
  rule.className = "docx-letterhead-rule";
  container.appendChild(rule);

  return container;
}

function isElementVisiblyRendered(element: Element): boolean {
  const style = window.getComputedStyle(element);

  return style.display !== "none" && style.visibility !== "hidden" && Number.parseFloat(style.opacity || "1") > 0;
}

function isRectInsidePageHeader(rect: DOMRect, pageRect: DOMRect): boolean {
  const headerBottom = pageRect.top + Math.max(170, pageRect.height * 0.16);

  return (
    rect.width > 2 &&
    rect.height > 2 &&
    rect.bottom >= pageRect.top &&
    rect.top <= headerBottom &&
    rect.right >= pageRect.left &&
    rect.left <= pageRect.right
  );
}

function isElementTopmostAtRect(element: Element, rect: DOMRect): boolean {
  const x = rect.left + (rect.width / 2);
  const y = rect.top + (rect.height / 2);
  const topElement = window.document.elementFromPoint(x, y);

  return Boolean(topElement && (topElement === element || element.contains(topElement)));
}

function pageHasVisibleLetterheadText(page: HTMLElement): boolean {
  const pageRect = page.getBoundingClientRect();
  const walker = window.document.createTreeWalker(page, window.NodeFilter.SHOW_TEXT);

  while (walker.nextNode()) {
    const node = walker.currentNode;
    const text = node.textContent?.toUpperCase() ?? "";

    if (!text.includes("HIMPUNAN") && !text.includes("SEKRETARIAT")) {
      continue;
    }

    const parent = node.parentElement;

    if (!parent || parent.closest(".docx-letterhead-fallback") || !isElementVisiblyRendered(parent)) {
      continue;
    }

    const range = window.document.createRange();
    range.selectNodeContents(node);
    const isVisible = Array.from(range.getClientRects()).some((rect) => (
      isRectInsidePageHeader(rect, pageRect) && isElementTopmostAtRect(parent, rect)
    ));
    range.detach();

    if (isVisible) {
      return true;
    }
  }

  return false;
}

function pageHasVisibleLetterheadImage(page: HTMLElement): boolean {
  const pageRect = page.getBoundingClientRect();

  return Array.from(page.querySelectorAll<HTMLElement>("header img, header svg, img, svg")).some((element) => {
    if (element.closest(".docx-letterhead-fallback") || !isElementVisiblyRendered(element)) {
      return false;
    }

    const rect = element.getBoundingClientRect();
    return rect.width > 20 && rect.height > 20 && isRectInsidePageHeader(rect, pageRect);
  });
}

function pageHasUsableNativeLetterhead(page: HTMLElement): boolean {
  return pageHasVisibleLetterheadText(page) && pageHasVisibleLetterheadImage(page);
}

function getFirstVisibleArticleTextTop(article: HTMLElement): number | undefined {
  const walker = window.document.createTreeWalker(article, window.NodeFilter.SHOW_TEXT);
  let firstTop: number | undefined;

  while (walker.nextNode()) {
    const node = walker.currentNode;

    if (!node.textContent?.trim()) {
      continue;
    }

    const parent = node.parentElement;

    if (!parent || !isElementVisiblyRendered(parent)) {
      continue;
    }

    const range = window.document.createRange();
    range.selectNodeContents(node);
    const top = Array.from(range.getClientRects())
      .filter((rect) => rect.width > 1 && rect.height > 1)
      .reduce<number | undefined>((lowestTop, rect) => (
        lowestTop === undefined ? rect.top : Math.min(lowestTop, rect.top)
      ), undefined);
    range.detach();

    if (top === undefined) {
      continue;
    }

    firstTop = firstTop === undefined ? top : Math.min(firstTop, top);
  }

  return firstTop;
}

function reserveLetterheadFallbackSpace(page: HTMLElement, fallback: HTMLElement): void {
  const article = getDirectChild(page, "article") ?? page.querySelector<HTMLElement>("article");
  const firstTextTop = article ? getFirstVisibleArticleTextTop(article) : undefined;

  if (!article || firstTextTop === undefined) {
    return;
  }

  const fallbackBottom = fallback.getBoundingClientRect().bottom;
  const overlap = fallbackBottom + 14 - firstTextTop;

  if (overlap <= 1) {
    return;
  }

  const currentPaddingTop = Number.parseFloat(window.getComputedStyle(article).paddingTop);
  article.style.boxSizing = "border-box";
  article.style.paddingTop = `${Math.ceil((Number.isFinite(currentPaddingTop) ? currentPaddingTop : 0) + overlap)}px`;
}

function hideNativeLetterhead(page: HTMLElement, letterhead: DocxLetterheadHint): void {
  if (!letterhead.repeatOnPages) {
    return;
  }

  const header = getDirectChild(page, "header");

  if (!header) {
    return;
  }

  header.setAttribute("aria-hidden", "true");
  header.style.visibility = "hidden";
  header.style.pointerEvents = "none";
}

function injectLetterheadFallback(root: HTMLElement, letterhead: DocxLetterheadHint | undefined): void {
  if (!letterhead) {
    return;
  }

  const pages = Array.from(root.querySelectorAll<HTMLElement>(DOCX_PAGE_SELECTOR));
  const targetPages = letterhead.repeatOnPages ? pages : pages.slice(0, 1);

  targetPages.forEach((page) => {
    if (page.querySelector(".docx-letterhead-fallback") || pageHasUsableNativeLetterhead(page)) {
      return;
    }

    page.style.position = "relative";
    page.classList.add("docx-has-letterhead-fallback");

    const fallback = createLetterheadFallback(letterhead);
    page.insertBefore(fallback, page.firstChild);
    hideNativeLetterhead(page, letterhead);
    reserveLetterheadFallbackSpace(page, fallback);
  });
}

function normalizeInlineFontInheritance(root: HTMLElement): void {
  root.querySelectorAll<HTMLElement>("span").forEach((span) => {
    if (span.style.fontFamily) {
      return;
    }

    const paragraph = span.closest<HTMLElement>("p");

    if (!paragraph) {
      return;
    }

    const spanStyle = window.getComputedStyle(span);
    const paragraphStyle = window.getComputedStyle(paragraph);

    if (spanStyle.fontFamily === paragraphStyle.fontFamily || !paragraphStyle.fontFamily) {
      return;
    }

    const hasInlineFormatting = Boolean(
      span.style.fontWeight ||
      span.style.fontStyle ||
      span.style.textDecoration ||
      span.style.textDecorationLine ||
      span.style.fontSize ||
      span.style.minHeight,
    );

    if (hasInlineFormatting && spanStyle.fontFamily.toLowerCase().includes("calibri")) {
      span.style.fontFamily = paragraphStyle.fontFamily;
    }
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
      letterhead: extractLetterheadHint(files),
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

  normalizeInlineFontInheritance(root);
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
  injectLetterheadFallback(root, hints.letterhead);
}

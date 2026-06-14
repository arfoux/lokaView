import { Bold, Italic, Maximize2, Minus, PencilLine, Plus, RotateCcw, Underline } from "lucide-react";
import type { ChangeEvent, MouseEvent } from "react";
import { useCallback, useEffect, useRef, useState } from "react";
import { renderAsync } from "docx-preview";
import { neutralizeUnsafeDocumentDom } from "../sanitizeDocumentDom";
import type { DocxOpenedDocument } from "../types";
import { applyDocxPostRenderFixes, getDocxLayoutHints, prepareDocxForPreview } from "./docxLayoutFixes";

const DOCX_PAGE_SELECTOR = "section.docx-preview, section.docx";
const DOCX_FONT_CHOICES = ["Times New Roman", "Calibri", "Arial", "Cambria", "Georgia"];
const DOCX_FONT_SIZES = ["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "32"];
const DOCX_LINE_SPACING = [
  { label: "1.0", value: "1" },
  { label: "1.15", value: "1.15" },
  { label: "1.5", value: "1.5" },
  { label: "2.0", value: "2" },
];
const DEFAULT_INLINE_FORMAT = { bold: false, italic: false, underline: false };

type InlineFormatState = typeof DEFAULT_INLINE_FORMAT;
type RichTextCommand = keyof InlineFormatState;

type SelectionFormatting = {
  fontFamily?: string;
  fontSize?: string;
  inline: InlineFormatState;
  lineSpacing?: string;
};

function getElementFromNode(node: Node): Element | null {
  return node.nodeType === Node.ELEMENT_NODE ? node as Element : node.parentNode instanceof Element ? node.parentNode : null;
}

function getTextNodesInRange(root: HTMLElement, range: Range) {
  const textNodes: Text[] = [];
  const walker = window.document.createTreeWalker(root, NodeFilter.SHOW_TEXT);

  while (walker.nextNode()) {
    const node = walker.currentNode as Text;

    if (node.data.length > 0 && range.intersectsNode(node)) {
      textNodes.push(node);
    }
  }

  return textNodes;
}

function getRangeStyleElement(body: HTMLElement, range: Range) {
  const selectedTextNode = getTextNodesInRange(body, range).find((node) => node.data.trim().length > 0);
  const selectedElement = selectedTextNode?.parentElement ?? getElementFromNode(range.startContainer);

  return selectedElement instanceof HTMLElement && body.contains(selectedElement) ? selectedElement : null;
}

function normalizeFontName(value: string) {
  return value.trim().replace(/^["']|["']$/g, "").toLowerCase();
}

function matchFontChoice(fontFamily: string) {
  const computedFamilies = fontFamily.split(",").map(normalizeFontName);

  return DOCX_FONT_CHOICES.find((font) => computedFamilies.includes(normalizeFontName(font)));
}

function parseCssLengthToPixels(value: string) {
  const numericValue = Number.parseFloat(value);

  if (!Number.isFinite(numericValue)) {
    return null;
  }

  if (value.endsWith("pt")) {
    return numericValue / 0.75;
  }

  if (value.endsWith("px") || /^[\d.]+$/.test(value)) {
    return numericValue;
  }

  return null;
}

function pickClosestChoice(value: number, choices: string[]) {
  return choices.reduce((closest, choice) => {
    const choiceValue = Number.parseFloat(choice);
    const closestValue = Number.parseFloat(closest);

    return Math.abs(choiceValue - value) < Math.abs(closestValue - value) ? choice : closest;
  }, choices[0] ?? "");
}

function matchFontSizeChoice(fontSize: string) {
  const pixels = parseCssLengthToPixels(fontSize);

  if (pixels === null) {
    return undefined;
  }

  return pickClosestChoice(pixels * 0.75, DOCX_FONT_SIZES);
}

function computedStyleIsBold(style: CSSStyleDeclaration) {
  const numericWeight = Number.parseInt(style.fontWeight, 10);

  return style.fontWeight === "bold" || style.fontWeight === "bolder" || (Number.isFinite(numericWeight) && numericWeight >= 600);
}

function computedStyleIsItalic(style: CSSStyleDeclaration) {
  return style.fontStyle === "italic" || style.fontStyle === "oblique";
}

function elementHasUnderline(element: HTMLElement, boundary: HTMLElement) {
  let current: HTMLElement | null = element;

  while (current && boundary.contains(current)) {
    const inlineDecoration = `${current.style.textDecorationLine} ${current.style.textDecoration}`.toLowerCase();

    if (inlineDecoration.includes("none")) {
      return false;
    }

    const style = window.getComputedStyle(current);
    const decoration = `${style.textDecorationLine} ${style.textDecoration}`.toLowerCase();

    if (decoration.includes("underline")) {
      return true;
    }

    current = current.parentElement;
  }

  return false;
}

function clearMatchingUnderlineAncestor(span: HTMLElement, boundary: HTMLElement) {
  let current = span.parentElement;

  while (current && boundary.contains(current)) {
    const decoration = `${current.style.textDecorationLine} ${current.style.textDecoration} ${window.getComputedStyle(current).textDecorationLine}`.toLowerCase();

    if (decoration.includes("underline") && current.textContent === span.textContent) {
      current.style.textDecoration = "none";
      current.style.textDecorationLine = "none";
      return;
    }

    current = current.parentElement;
  }
}

function matchLineSpacingChoice(style: CSSStyleDeclaration) {
  if (!style.lineHeight || style.lineHeight === "normal") {
    return undefined;
  }

  const lineHeightPixels = parseCssLengthToPixels(style.lineHeight);
  const fontSizePixels = parseCssLengthToPixels(style.fontSize);

  if (lineHeightPixels === null || fontSizePixels === null || fontSizePixels === 0) {
    return undefined;
  }

  return pickClosestChoice(lineHeightPixels / fontSizePixels, DOCX_LINE_SPACING.map((spacing) => spacing.value));
}

function readSelectionFormatting(body: HTMLElement, range: Range): SelectionFormatting | null {
  const styleElement = getRangeStyleElement(body, range);

  if (!styleElement) {
    return null;
  }

  const textElements = getTextNodesInRange(body, range)
    .filter((node) => node.data.trim().length > 0)
    .map((node) => node.parentElement)
    .filter((element): element is HTMLElement => element instanceof HTMLElement && body.contains(element));
  const styledElements = textElements.length > 0 ? textElements : [styleElement];
  const computedStyles = styledElements.map((element) => window.getComputedStyle(element));
  const firstStyle = computedStyles[0] ?? window.getComputedStyle(styleElement);
  const paragraphStyle = window.getComputedStyle(styleElement.closest<HTMLElement>("p") ?? styleElement);

  return {
    fontFamily: matchFontChoice(firstStyle.fontFamily),
    fontSize: matchFontSizeChoice(firstStyle.fontSize),
    inline: {
      bold: computedStyles.length > 0 && computedStyles.every(computedStyleIsBold),
      italic: computedStyles.length > 0 && computedStyles.every(computedStyleIsItalic),
      underline: styledElements.length > 0 && styledElements.every((element) => elementHasUnderline(element, body)),
    },
    lineSpacing: matchLineSpacingChoice(paragraphStyle),
  };
}

function getPreservedTextStyle(source: Element | null, incomingStyle: Partial<CSSStyleDeclaration>) {
  if (!(source instanceof HTMLElement)) {
    return {};
  }

  const computedStyle = window.getComputedStyle(source);
  const preservedStyle: Partial<CSSStyleDeclaration> = {};

  if (!incomingStyle.fontFamily && computedStyle.fontFamily) {
    preservedStyle.fontFamily = computedStyle.fontFamily;
  }

  if (!incomingStyle.fontSize && computedStyle.fontSize) {
    preservedStyle.fontSize = computedStyle.fontSize;
  }

  return preservedStyle;
}

function getRichTextCommandStyle(command: RichTextCommand, shouldEnable: boolean): Partial<CSSStyleDeclaration> {
  if (command === "bold") {
    return { fontWeight: shouldEnable ? "700" : "400" };
  }

  if (command === "italic") {
    return { fontStyle: shouldEnable ? "italic" : "normal" };
  }

  return shouldEnable
    ? { textDecorationLine: "underline" }
    : { textDecoration: "none", textDecorationLine: "none" };
}

export function DocxViewer({ document }: { document: DocxOpenedDocument }) {
  const bodyRef = useRef<HTMLDivElement>(null);
  const styleRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const zoomRef = useRef(1);
  const savedRangeRef = useRef<Range | null>(null);
  const [zoom, setZoom] = useState(1);
  const [fitWidth, setFitWidth] = useState(true);
  const [editMode, setEditMode] = useState(false);
  const [selectedFont, setSelectedFont] = useState(DOCX_FONT_CHOICES[0] ?? "Times New Roman");
  const [selectedFontSize, setSelectedFontSize] = useState("12");
  const [selectedInlineStyles, setSelectedInlineStyles] = useState<InlineFormatState>(DEFAULT_INLINE_FORMAT);
  const [selectedLineSpacing, setSelectedLineSpacing] = useState("1.15");
  const [status, setStatus] = useState<"loading" | "ready" | "error">("loading");
  const [errorMessage, setErrorMessage] = useState("");

  useEffect(() => {
    zoomRef.current = zoom;
  }, [zoom]);

  const fitDocumentWidth = useCallback(() => {
    const body = bodyRef.current;
    const scroll = scrollRef.current;
    const page = body?.querySelector<HTMLElement>(DOCX_PAGE_SELECTOR);

    if (!body || !scroll || !page) {
      return;
    }

    const visualWidth = page.getBoundingClientRect().width;
    const intrinsicWidth = visualWidth / zoomRef.current;
    const availableWidth = Math.max(scroll.clientWidth - 32, 320);

    if (intrinsicWidth > 0) {
      setZoom(Math.min(1, Math.max(0.45, availableWidth / intrinsicWidth)));
    }
  }, []);

  useEffect(() => {
    let canceled = false;
    const body = bodyRef.current;
    const styles = styleRef.current;

    if (!body || !styles) {
      return;
    }

    body.innerHTML = "";
    styles.innerHTML = "";
    setStatus("loading");
    setEditMode(false);
    setSelectedInlineStyles(DEFAULT_INLINE_FORMAT);

    const layoutHints = getDocxLayoutHints(document.arrayBuffer);
    const previewBuffer = prepareDocxForPreview(document.arrayBuffer);

    renderAsync(previewBuffer.slice(0), body, styles, {
      className: "docx-preview",
      inWrapper: true,
      breakPages: true,
      ignoreFonts: false,
      experimental: true,
      ignoreWidth: false,
      ignoreHeight: false,
      ignoreLastRenderedPageBreak: false,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
      renderComments: false,
      renderAltChunks: false,
      useBase64URL: false,
      debug: false,
    })
      .then(() => {
        if (canceled) {
          return;
        }
        neutralizeUnsafeDocumentDom(body);
        applyDocxPostRenderFixes(body, layoutHints);
        setStatus("ready");
        window.requestAnimationFrame(() => {
          if (!canceled) {
            fitDocumentWidth();
          }
        });
      })
      .catch((error: unknown) => {
        if (canceled) {
          return;
        }
        if (import.meta.env.DEV) {
          console.error("DOCX render failed", error);
        }
        setStatus("error");
        setErrorMessage("The Word preview could not be rendered. The file may be encrypted, corrupted, or contain unsupported content.");
      });

    return () => {
      canceled = true;
      body.innerHTML = "";
      styles.innerHTML = "";
    };
  }, [document, fitDocumentWidth]);

  const syncSelectionFormatting = useCallback((formatting: SelectionFormatting | null) => {
    if (!formatting) {
      return;
    }

    setSelectedInlineStyles(formatting.inline);

    if (formatting.fontFamily) {
      setSelectedFont(formatting.fontFamily);
    }

    if (formatting.fontSize) {
      setSelectedFontSize(formatting.fontSize);
    }

    if (formatting.lineSpacing) {
      setSelectedLineSpacing(formatting.lineSpacing);
    }
  }, []);

  useEffect(() => {
    if (!editMode) {
      savedRangeRef.current = null;
      return;
    }

    const saveSelection = () => {
      const body = bodyRef.current;
      const selection = window.getSelection();

      if (!body || !selection || selection.rangeCount === 0) {
        return;
      }

      const range = selection.getRangeAt(0);
      const container = getElementFromNode(range.commonAncestorContainer);

      if (container && body.contains(container)) {
        savedRangeRef.current = range.cloneRange();
        syncSelectionFormatting(readSelectionFormatting(body, range));
      }
    };

    window.document.addEventListener("selectionchange", saveSelection);

    return () => window.document.removeEventListener("selectionchange", saveSelection);
  }, [editMode, syncSelectionFormatting]);

  const getEditableSelection = useCallback(() => {
    const body = bodyRef.current;
    const selection = window.getSelection();

    if (!body || !selection) {
      return null;
    }

    const currentRange = selection.rangeCount > 0 ? selection.getRangeAt(0) : null;
    const currentContainer = currentRange ? getElementFromNode(currentRange.commonAncestorContainer) : null;

    if (currentRange && currentContainer && body.contains(currentContainer)) {
      return { body, range: currentRange, selection };
    }

    const savedRange = savedRangeRef.current;
    const savedContainer = savedRange ? getElementFromNode(savedRange.commonAncestorContainer) : null;

    if (!savedRange || !savedContainer || !body.contains(savedContainer)) {
      return null;
    }

    selection.removeAllRanges();
    selection.addRange(savedRange.cloneRange());

    return { body, range: selection.getRangeAt(0), selection };
  }, []);

  const restoreSelectionAround = useCallback((nodes: HTMLElement[]) => {
    if (nodes.length === 0) {
      return null;
    }

    const range = window.document.createRange();
    range.setStartBefore(nodes[0]!);
    range.setEndAfter(nodes[nodes.length - 1]!);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
    savedRangeRef.current = range.cloneRange();

    return range;
  }, []);

  const applyInlineStyle = useCallback((style: Partial<CSSStyleDeclaration>) => {
    if (!editMode) {
      return;
    }

    const editableSelection = getEditableSelection();

    if (!editableSelection || editableSelection.selection.isCollapsed) {
      return;
    }

    const { body, range } = editableSelection;
    const textNodes = getTextNodesInRange(body, range);
    const styledNodes: HTMLElement[] = [];

    textNodes.forEach((node) => {
      const startOffset = node === range.startContainer ? range.startOffset : 0;
      const endOffset = node === range.endContainer ? range.endOffset : node.data.length;

      if (startOffset >= endOffset) {
        return;
      }

      let selectedText = node;

      if (startOffset > 0) {
        selectedText = selectedText.splitText(startOffset);
      }

      if (endOffset - startOffset < selectedText.data.length) {
        selectedText.splitText(endOffset - startOffset);
      }

      const span = window.document.createElement("span");
      Object.assign(span.style, getPreservedTextStyle(selectedText.parentElement, style), style);
      selectedText.parentNode?.insertBefore(span, selectedText);
      span.appendChild(selectedText);

      if (style.textDecoration === "none" || style.textDecorationLine === "none") {
        clearMatchingUnderlineAncestor(span, body);
      }

      styledNodes.push(span);
    });

    const restoredRange = restoreSelectionAround(styledNodes);

    if (restoredRange) {
      syncSelectionFormatting(readSelectionFormatting(body, restoredRange));
    }
  }, [editMode, getEditableSelection, restoreSelectionAround, syncSelectionFormatting]);

  const applyRichTextCommand = useCallback((command: RichTextCommand) => {
    const editableSelection = getEditableSelection();

    if (!editableSelection || editableSelection.selection.isCollapsed) {
      return;
    }

    const formatting = readSelectionFormatting(editableSelection.body, editableSelection.range);
    const shouldEnable = !(formatting?.inline[command] ?? selectedInlineStyles[command]);

    applyInlineStyle(getRichTextCommandStyle(command, shouldEnable));
  }, [applyInlineStyle, getEditableSelection, selectedInlineStyles]);

  const applyLineSpacing = useCallback((value: string) => {
    if (!editMode) {
      return;
    }

    const editableSelection = getEditableSelection();

    if (!editableSelection) {
      return;
    }

    const paragraphs = Array.from(editableSelection.body.querySelectorAll<HTMLElement>("p"))
      .filter((paragraph) => editableSelection.range.intersectsNode(paragraph));
    const currentParagraph = getElementFromNode(editableSelection.range.startContainer)?.closest<HTMLElement>("p") ?? null;
    const targetParagraphs = paragraphs.length > 0
      ? paragraphs
      : currentParagraph ? [currentParagraph] : [];

    targetParagraphs.forEach((paragraph) => {
      paragraph.style.lineHeight = value;
    });
  }, [editMode, getEditableSelection]);

  const keepToolbarSelection = useCallback((event: MouseEvent<HTMLElement>) => {
    event.preventDefault();
  }, []);

  const handleFontChange = useCallback((event: ChangeEvent<HTMLSelectElement>) => {
    const value = event.target.value;
    setSelectedFont(value);
    applyInlineStyle({ fontFamily: value });
  }, [applyInlineStyle]);

  const handleFontSizeChange = useCallback((event: ChangeEvent<HTMLSelectElement>) => {
    const value = event.target.value;
    setSelectedFontSize(value);
    applyInlineStyle({ fontSize: `${value}pt`, minHeight: `${value}pt` });
  }, [applyInlineStyle]);

  const handleLineSpacingChange = useCallback((event: ChangeEvent<HTMLSelectElement>) => {
    const value = event.target.value;
    setSelectedLineSpacing(value);
    applyLineSpacing(value);
  }, [applyLineSpacing]);

  useEffect(() => {
    if (!fitWidth || status !== "ready") {
      return;
    }

    const scroll = scrollRef.current;
    if (!scroll || typeof ResizeObserver === "undefined") {
      return;
    }

    const observer = new ResizeObserver(() => fitDocumentWidth());
    observer.observe(scroll);

    return () => observer.disconnect();
  }, [fitDocumentWidth, fitWidth, status]);

  return (
    <div className="office-viewer docx-viewer">
      <div className="viewer-controls">
        <p className="viewer-note">Word preview layout may differ slightly from the original document.</p>
        <div className="segmented-controls" aria-label="DOCX zoom controls">
          <button type="button" className="icon-button" onClick={() => { setFitWidth(false); setZoom((value) => Math.max(0.5, value - 0.1)); }} title="Zoom out">
            <Minus aria-hidden="true" size={17} />
            <span className="sr-only">Zoom out</span>
          </button>
          <span>{Math.round(zoom * 100)}%</span>
          <button type="button" className="icon-button" onClick={() => { setFitWidth(false); setZoom((value) => Math.min(2, value + 0.1)); }} title="Zoom in">
            <Plus aria-hidden="true" size={17} />
            <span className="sr-only">Zoom in</span>
          </button>
          <button type="button" className="text-button" onClick={() => { setFitWidth(true); fitDocumentWidth(); }}>
            Fit width
          </button>
          <button type="button" className="icon-button" onClick={() => { setFitWidth(false); setZoom(1); }} title="Reset zoom">
            <RotateCcw aria-hidden="true" size={17} />
            <span className="sr-only">Reset zoom</span>
          </button>
          <button type="button" className="icon-button" onClick={() => { void scrollRef.current?.requestFullscreen?.(); }} title="Fullscreen">
            <Maximize2 aria-hidden="true" size={17} />
            <span className="sr-only">Fullscreen</span>
          </button>
        </div>
        <div className="segmented-controls docx-edit-controls" aria-label="DOCX edit controls">
          <button type="button" className={`text-button${editMode ? " is-active" : ""}`} onClick={() => setEditMode((value) => !value)} disabled={status !== "ready"}>
            <PencilLine aria-hidden="true" size={16} />
            Edit
          </button>
          <select aria-label="Font family" value={selectedFont} onChange={handleFontChange} disabled={!editMode || status !== "ready"}>
            {DOCX_FONT_CHOICES.map((font) => (
              <option key={font} value={font}>{font}</option>
            ))}
          </select>
          <select aria-label="Font size" value={selectedFontSize} onChange={handleFontSizeChange} disabled={!editMode || status !== "ready"}>
            {DOCX_FONT_SIZES.map((size) => (
              <option key={size} value={size}>{size} pt</option>
            ))}
          </select>
          <button
            type="button"
            className={`icon-button${selectedInlineStyles.bold ? " is-active" : ""}`}
            aria-pressed={selectedInlineStyles.bold}
            onMouseDown={keepToolbarSelection}
            onClick={() => applyRichTextCommand("bold")}
            disabled={!editMode || status !== "ready"}
            title="Bold"
          >
            <Bold aria-hidden="true" size={17} />
            <span className="sr-only">Bold</span>
          </button>
          <button
            type="button"
            className={`icon-button${selectedInlineStyles.italic ? " is-active" : ""}`}
            aria-pressed={selectedInlineStyles.italic}
            onMouseDown={keepToolbarSelection}
            onClick={() => applyRichTextCommand("italic")}
            disabled={!editMode || status !== "ready"}
            title="Italic"
          >
            <Italic aria-hidden="true" size={17} />
            <span className="sr-only">Italic</span>
          </button>
          <button
            type="button"
            className={`icon-button${selectedInlineStyles.underline ? " is-active" : ""}`}
            aria-pressed={selectedInlineStyles.underline}
            onMouseDown={keepToolbarSelection}
            onClick={() => applyRichTextCommand("underline")}
            disabled={!editMode || status !== "ready"}
            title="Underline"
          >
            <Underline aria-hidden="true" size={17} />
            <span className="sr-only">Underline</span>
          </button>
          <select aria-label="Line spacing" value={selectedLineSpacing} onChange={handleLineSpacingChange} disabled={!editMode || status !== "ready"}>
            {DOCX_LINE_SPACING.map((spacing) => (
              <option key={spacing.value} value={spacing.value}>{spacing.label}</option>
            ))}
          </select>
        </div>
      </div>
      <div ref={styleRef} />
      {status === "loading" && <p className="viewer-note">Rendering document locally...</p>}
      {status === "error" && <p className="viewer-error">{errorMessage}</p>}
      <div ref={scrollRef} className="docx-scroll">
        <div
          ref={bodyRef}
          className={`docx-render-surface${editMode ? " is-editing" : ""}`}
          contentEditable={editMode}
          suppressContentEditableWarning
          style={{ zoom }}
        />
      </div>
    </div>
  );
}

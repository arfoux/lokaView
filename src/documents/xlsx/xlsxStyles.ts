import { unzipSync } from "fflate";
import type { SpreadsheetCellStyle } from "../types";

interface WorkbookSheetInfo {
  name: string;
  path: string;
}

interface ExtractedSheetStyles {
  cellStyles: Record<string, SpreadsheetCellStyle>;
  columnStyles: Record<number, SpreadsheetCellStyle>;
  rowStyles: Record<number, SpreadsheetCellStyle>;
}

export type ExtractedWorkbookStyles = Record<string, ExtractedSheetStyles>;

interface ParsedStyleTable {
  fills: Array<Pick<SpreadsheetCellStyle, "backgroundColor">>;
  fonts: Array<Pick<SpreadsheetCellStyle, "color" | "fontStyle" | "fontWeight" | "textDecoration">>;
  cellXfs: SpreadsheetCellStyle[];
}

const decoder = new TextDecoder("utf-8", { fatal: false });

const INDEXED_COLORS: Record<number, string> = {
  0: "#000000",
  1: "#ffffff",
  2: "#ff0000",
  3: "#00ff00",
  4: "#0000ff",
  5: "#ffff00",
  6: "#ff00ff",
  7: "#00ffff",
  8: "#000000",
  9: "#ffffff",
  10: "#ff0000",
  11: "#00ff00",
  12: "#0000ff",
  13: "#ffff00",
  14: "#ff00ff",
  15: "#00ffff",
};

function getXml(files: Record<string, Uint8Array>, path: string): Document | undefined {
  const bytes = files[path];
  if (!bytes || typeof DOMParser === "undefined") {
    return undefined;
  }

  return new DOMParser().parseFromString(decoder.decode(bytes), "application/xml");
}

function childrenByLocalName(element: Element | Document, name: string): Element[] {
  return Array.from(element.getElementsByTagName("*")).filter((child) => child.localName === name);
}

function directChildrenByLocalName(element: Element, name: string): Element[] {
  return Array.from(element.children).filter((child) => child.localName === name);
}

function normalizePath(path: string): string {
  const parts: string[] = [];

  for (const part of path.replaceAll("\\", "/").split("/")) {
    if (!part || part === ".") {
      continue;
    }

    if (part === "..") {
      parts.pop();
    } else {
      parts.push(part);
    }
  }

  return parts.join("/");
}

function resolveRelationshipTarget(basePath: string, target: string): string {
  if (target.startsWith("/")) {
    return normalizePath(target.slice(1));
  }

  const baseDirectory = basePath.includes("/") ? basePath.slice(0, basePath.lastIndexOf("/")) : "";
  return normalizePath(`${baseDirectory}/${target}`);
}

function parseWorkbookSheets(files: Record<string, Uint8Array>): WorkbookSheetInfo[] {
  const workbook = getXml(files, "xl/workbook.xml");
  const relationships = getXml(files, "xl/_rels/workbook.xml.rels");

  if (!workbook || !relationships) {
    return [];
  }

  const relTargets = new Map<string, string>();

  for (const relationship of childrenByLocalName(relationships, "Relationship")) {
    const id = relationship.getAttribute("Id");
    const target = relationship.getAttribute("Target");
    const type = relationship.getAttribute("Type") ?? "";

    if (id && target && type.includes("/worksheet")) {
      relTargets.set(id, resolveRelationshipTarget("xl/workbook.xml", target));
    }
  }

  return childrenByLocalName(workbook, "sheet")
    .map((sheet) => {
      const name = sheet.getAttribute("name");
      const relationshipId = sheet.getAttribute("r:id") ?? sheet.getAttribute("id");
      const path = relationshipId ? relTargets.get(relationshipId) : undefined;

      return name && path ? { name, path } : undefined;
    })
    .filter((sheet): sheet is WorkbookSheetInfo => Boolean(sheet));
}

function parseThemeColors(files: Record<string, Uint8Array>): string[] {
  const theme = getXml(files, "xl/theme/theme1.xml");
  const scheme = theme ? childrenByLocalName(theme, "clrScheme")[0] : undefined;

  if (!scheme) {
    return [];
  }

  return Array.from(scheme.children).map((slot) => {
    const color = directChildrenByLocalName(slot, "srgbClr")[0]?.getAttribute("val")
      ?? directChildrenByLocalName(slot, "sysClr")[0]?.getAttribute("lastClr");
    return color ? `#${color.slice(-6).toLowerCase()}` : "";
  });
}

function clampChannel(value: number): number {
  return Math.max(0, Math.min(255, Math.round(value)));
}

function applyTint(hex: string, tint?: number): string {
  if (!tint) {
    return hex;
  }

  const clean = hex.replace("#", "");
  const channels = [0, 2, 4].map((offset) => Number.parseInt(clean.slice(offset, offset + 2), 16));
  const tinted = channels.map((channel) => {
    if (tint < 0) {
      return clampChannel(channel * (1 + tint));
    }

    return clampChannel(channel + (255 - channel) * tint);
  });

  return `#${tinted.map((channel) => channel.toString(16).padStart(2, "0")).join("")}`;
}

function parseColor(element: Element | undefined, themeColors: string[]): string | undefined {
  if (!element) {
    return undefined;
  }

  const rgb = element.getAttribute("rgb");
  const indexed = element.getAttribute("indexed");
  const theme = element.getAttribute("theme");
  const tint = element.getAttribute("tint");
  const tintValue = tint === null ? undefined : Number(tint);

  if (rgb) {
    return applyTint(`#${rgb.slice(-6).toLowerCase()}`, tintValue);
  }

  if (theme !== null) {
    const themeColor = themeColors[Number(theme)];
    return themeColor ? applyTint(themeColor, tintValue) : undefined;
  }

  if (indexed !== null) {
    return INDEXED_COLORS[Number(indexed)];
  }

  return undefined;
}

function mergeStyle(...styles: Array<SpreadsheetCellStyle | undefined>): SpreadsheetCellStyle {
  const merged: SpreadsheetCellStyle = {};

  for (const style of styles) {
    if (style) {
      Object.assign(merged, style);
    }
  }

  return merged;
}

function isEmptyStyle(style: SpreadsheetCellStyle): boolean {
  return Object.values(style).every((value) => value === undefined);
}

function parseStyleTable(files: Record<string, Uint8Array>): ParsedStyleTable {
  const styles = getXml(files, "xl/styles.xml");
  const themeColors = parseThemeColors(files);

  if (!styles) {
    return { fills: [], fonts: [], cellXfs: [] };
  }

  const fills = directChildrenByLocalName(childrenByLocalName(styles, "fills")[0] ?? styles.documentElement, "fill")
    .map((fill) => {
      const pattern = childrenByLocalName(fill, "patternFill")[0];
      const patternType = pattern?.getAttribute("patternType");
      const color = parseColor(
        childrenByLocalName(pattern ?? fill, "fgColor")[0] ?? childrenByLocalName(pattern ?? fill, "bgColor")[0],
        themeColors,
      );

      return patternType && patternType !== "none" && patternType !== "gray125" && color
        ? { backgroundColor: color }
        : {};
    });

  const fonts = directChildrenByLocalName(childrenByLocalName(styles, "fonts")[0] ?? styles.documentElement, "font")
    .map((font) => ({
      color: parseColor(childrenByLocalName(font, "color")[0], themeColors),
      fontWeight: childrenByLocalName(font, "b").length > 0 ? "bold" as const : undefined,
      fontStyle: childrenByLocalName(font, "i").length > 0 ? "italic" as const : undefined,
      textDecoration: childrenByLocalName(font, "u").length > 0 ? "underline" as const : undefined,
    }));

  const cellXfs = directChildrenByLocalName(childrenByLocalName(styles, "cellXfs")[0] ?? styles.documentElement, "xf")
    .map((xf) => {
      const fillId = Number(xf.getAttribute("fillId") ?? 0);
      const fontId = Number(xf.getAttribute("fontId") ?? 0);
      const alignment = childrenByLocalName(xf, "alignment")[0]?.getAttribute("horizontal");
      const horizontalAlign =
        alignment === "center" || alignment === "right" || alignment === "left" ? alignment : undefined;
      const style = mergeStyle(fonts[fontId], fills[fillId], { horizontalAlign });

      return isEmptyStyle(style) ? {} : style;
    });

  return { fills, fonts, cellXfs };
}

function decodeCellReference(reference: string): { row: number; column: number } | undefined {
  const match = /^([A-Z]+)(\d+)$/i.exec(reference);

  if (!match) {
    return undefined;
  }

  const letters = match[1]?.toUpperCase();

  if (!letters || !match[2]) {
    return undefined;
  }
  let column = 0;

  for (const letter of letters) {
    column = column * 26 + (letter.charCodeAt(0) - 64);
  }

  return { row: Number(match[2]) - 1, column: column - 1 };
}

function styleFromIndex(styleTable: ParsedStyleTable, index: string | null): SpreadsheetCellStyle | undefined {
  if (index === null) {
    return undefined;
  }

  return styleTable.cellXfs[Number(index)];
}

function applyStyle(
  target: Record<string, SpreadsheetCellStyle>,
  key: string,
  style: SpreadsheetCellStyle | undefined,
): void {
  if (!style || isEmptyStyle(style)) {
    return;
  }

  target[key] = mergeStyle(target[key], style);
}

function extractSheetStyles(sheetXml: Document, styleTable: ParsedStyleTable): ExtractedSheetStyles {
  const cellStyles: Record<string, SpreadsheetCellStyle> = {};
  const columnStyles: Record<number, SpreadsheetCellStyle> = {};
  const rowStyles: Record<number, SpreadsheetCellStyle> = {};

  for (const column of childrenByLocalName(sheetXml, "col")) {
    const style = styleFromIndex(styleTable, column.getAttribute("style"));
    const min = Number(column.getAttribute("min") ?? 0);
    const max = Number(column.getAttribute("max") ?? min);

    if (style && min > 0 && max >= min) {
      for (let index = min - 1; index <= max - 1; index += 1) {
        columnStyles[index] = mergeStyle(columnStyles[index], style);
      }
    }
  }

  for (const row of childrenByLocalName(sheetXml, "row")) {
    const rowIndex = Number(row.getAttribute("r") ?? 0) - 1;
    const rowStyle = styleFromIndex(styleTable, row.getAttribute("s"));

    if (rowIndex >= 0 && rowStyle) {
      rowStyles[rowIndex] = mergeStyle(rowStyles[rowIndex], rowStyle);
    }

    for (const cell of directChildrenByLocalName(row, "c")) {
      const reference = cell.getAttribute("r");
      const decoded = reference ? decodeCellReference(reference) : undefined;
      const cellStyle = styleFromIndex(styleTable, cell.getAttribute("s"));

      if (decoded && cellStyle) {
        applyStyle(cellStyles, `${decoded.row}:${decoded.column}`, cellStyle);
      }
    }
  }

  return { cellStyles, columnStyles, rowStyles };
}

export function extractWorkbookStyles(arrayBuffer: ArrayBuffer): ExtractedWorkbookStyles {
  if (typeof DOMParser === "undefined") {
    return {};
  }

  const files = unzipSync(new Uint8Array(arrayBuffer));
  const workbookSheets = parseWorkbookSheets(files);
  const styleTable = parseStyleTable(files);
  const result: ExtractedWorkbookStyles = {};

  for (const sheet of workbookSheets) {
    const sheetXml = getXml(files, sheet.path);

    if (sheetXml) {
      result[sheet.name] = extractSheetStyles(sheetXml, styleTable);
    }
  }

  return result;
}

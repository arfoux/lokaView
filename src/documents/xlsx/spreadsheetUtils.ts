import type { SpreadsheetMerge, SpreadsheetSheet } from "../types";

export function columnLabel(index: number): string {
  let value = index + 1;
  let label = "";

  while (value > 0) {
    const modulo = (value - 1) % 26;
    label = String.fromCharCode(65 + modulo) + label;
    value = Math.floor((value - modulo) / 26);
  }

  return label;
}

export function escapeCsvCell(value: string): string {
  if (/[",\r\n]/.test(value)) {
    return `"${value.replaceAll("\"", "\"\"")}"`;
  }

  return value;
}

export function sheetToCsv(sheet: SpreadsheetSheet): string {
  return sheet.rows.map((row) => row.map(escapeCsvCell).join(",")).join("\n");
}

export function sheetToJson(sheet: SpreadsheetSheet): string {
  const headers = sheet.rows[0]?.map((value, index) => value || columnLabel(index)) ?? [];
  const data = sheet.rows.slice(1).map((row) =>
    Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ""])),
  );

  return JSON.stringify(data, null, 2);
}

export function downloadTextFile(fileName: string, text: string, type: string): void {
  const blob = new Blob([text], { type });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

export function getMergeAt(
  merges: readonly SpreadsheetMerge[],
  row: number,
  column: number,
): SpreadsheetMerge | undefined {
  return merges.find((merge) => merge.startRow === row && merge.startColumn === column);
}

export function isCoveredByMerge(
  merges: readonly SpreadsheetMerge[],
  row: number,
  column: number,
): boolean {
  return merges.some(
    (merge) =>
      row >= merge.startRow &&
      row < merge.startRow + merge.rowSpan &&
      column >= merge.startColumn &&
      column < merge.startColumn + merge.columnSpan &&
      !(row === merge.startRow && column === merge.startColumn),
  );
}

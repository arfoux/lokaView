import readXlsxFile from "read-excel-file/browser";
import type { CellValue, Sheet } from "read-excel-file/browser";
import { FILE_LIMITS } from "../limits";
import { createOpenedDocumentBase } from "../openedDocument";
import type { DocumentAdapter, SpreadsheetMerge, SpreadsheetSheet, XlsxOpenedDocument } from "../types";
import { DocumentError } from "../types";
import { assertZipWithinLimits, inspectZip } from "../zipInspection";
import { extractWorkbookStyles, type ExtractedWorkbookStyles } from "./xlsxStyles";

function stringifyCell(value: CellValue | null | undefined): string {
  if (value === null || value === undefined) {
    return "";
  }

  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }

  return String(value);
}

function clampMergeToPreview(
  merge: SpreadsheetMerge,
  renderedRows: number,
  renderedColumns: number,
): SpreadsheetMerge | undefined {
  if (merge.startRow >= renderedRows || merge.startColumn >= renderedColumns) {
    return undefined;
  }

  const rowSpan = Math.min(merge.rowSpan, renderedRows - merge.startRow);
  const columnSpan = Math.min(merge.columnSpan, renderedColumns - merge.startColumn);

  return rowSpan > 1 || columnSpan > 1
    ? { ...merge, rowSpan, columnSpan }
    : undefined;
}

function toSheetPreview(sheet: Sheet, extractedStyles?: ExtractedWorkbookStyles[string]): SpreadsheetSheet {
  const rowCount = sheet.data.length;
  const columnCount = sheet.data.reduce((max, row) => Math.max(max, row.length), 0);
  const renderedRows = Math.min(rowCount, FILE_LIMITS.sheetPreviewRows);
  const renderedColumns = Math.min(columnCount, FILE_LIMITS.sheetPreviewColumns);
  const rows = sheet.data.slice(0, renderedRows).map((row) =>
    Array.from({ length: renderedColumns }, (_, index) => stringifyCell(row[index])),
  );
  const styles: SpreadsheetSheet["styles"] = {};
  const merges = (extractedStyles?.merges ?? [])
    .map((merge) => clampMergeToPreview(merge, renderedRows, renderedColumns))
    .filter((merge): merge is SpreadsheetMerge => Boolean(merge));

  for (let row = 0; row < renderedRows; row += 1) {
    for (let column = 0; column < renderedColumns; column += 1) {
      const style = {
        ...extractedStyles?.columnStyles[column],
        ...extractedStyles?.rowStyles[row],
        ...extractedStyles?.cellStyles[`${row}:${column}`],
      };

      if (Object.values(style).some((value) => value !== undefined)) {
        styles[`${row}:${column}`] = style;
      }
    }
  }

  return {
    name: sheet.sheet,
    rows,
    rowCount,
    columnCount,
    truncatedRows: rowCount > renderedRows,
    truncatedColumns: columnCount > renderedColumns,
    merges,
    styles,
  };
}

export const xlsxAdapter: DocumentAdapter<XlsxOpenedDocument> = {
  kind: "xlsx",
  label: "Excel",
  acceptedExtensions: [".xlsx"],
  acceptedMimeTypes: ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"],
  canHandle(_file, _signature, detection) {
    return detection?.kind === "xlsx";
  },
  async open({ file, arrayBuffer, detection }) {
    const zip = inspectZip(arrayBuffer);
    assertZipWithinLimits(zip);

    if (!zip.isZip) {
      throw new DocumentError({
        code: "parse-failed",
        title: "This is not a valid XLSX package",
        message: "XLSX files must be ZIP-based Office Open XML packages. This file could not be inspected as one.",
      });
    }

    try {
      const workbook = await readXlsxFile(arrayBuffer.slice(0));
      const workbookStyles = extractWorkbookStyles(arrayBuffer);
      const sheets = workbook.map((sheet) => toSheetPreview(sheet, workbookStyles[sheet.sheet]));

      if (sheets.length === 0) {
        throw new DocumentError({
          code: "parse-failed",
          title: "Workbook has no readable sheets",
          message: "This XLSX file opened, but no worksheet content could be read.",
        });
      }

      const workbookWarnings = sheets
        .filter((sheet) => sheet.truncatedRows || sheet.truncatedColumns)
        .map(
          (sheet) =>
            `${sheet.name} is previewed at ${sheet.rows.length} rows by ${
              sheet.rows[0]?.length ?? 0
            } columns for browser responsiveness.`,
        );

      const base = createOpenedDocumentBase({
        kind: "xlsx",
        file,
        detection,
        capabilities: [
          "download-original",
          "sheet-navigation",
          "export-csv",
          "export-json",
          "search",
        ],
        warnings: [...zip.warnings, ...workbookWarnings],
      });

      return {
        ...base,
        kind: "xlsx",
        sheets,
        workbookWarnings,
      };
    } catch (error) {
      if (error instanceof DocumentError) {
        throw error;
      }

      throw new DocumentError({
        code: "parse-failed",
        title: "Could not open this workbook",
        message: "The workbook may be corrupted, encrypted, or use unsupported spreadsheet content.",
        cause: error,
      });
    }
  },
};

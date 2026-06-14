import { createOpenedDocumentBase } from "../openedDocument";
import type { CsvOpenedDocument, DocumentAdapter } from "../types";
import { parseCsv } from "./csvParser";

const decoder = new TextDecoder("utf-8", { fatal: false });

export const csvAdapter: DocumentAdapter<CsvOpenedDocument> = {
  kind: "csv",
  label: "CSV",
  acceptedExtensions: [".csv"],
  acceptedMimeTypes: ["text/csv", "application/csv"],
  canHandle(_file, _signature, detection) {
    return detection?.kind === "csv";
  },
  open({ file, arrayBuffer, detection }) {
    const parsed = parseCsv(decoder.decode(arrayBuffer));
    const base = createOpenedDocumentBase({
      kind: "csv",
      file,
      detection,
      capabilities: ["download-original", "sheet-navigation", "export-csv", "export-json", "search"],
      warnings: [
        ...(parsed.confidence === "low"
          ? ["Delimiter detection was inconclusive. Comma was used for the initial preview."]
          : []),
      ],
    });

    return Promise.resolve({
      ...base,
      kind: "csv",
      delimiter: parsed.delimiter,
      delimiterConfidence: parsed.confidence,
      sheet: {
        name: "CSV",
        rows: parsed.rows,
        rowCount: parsed.rowCount,
        columnCount: parsed.columnCount,
        truncatedRows: parsed.truncatedRows,
        truncatedColumns: parsed.truncatedColumns,
        merges: [],
        styles: {},
      },
    });
  },
};

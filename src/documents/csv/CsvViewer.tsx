import { Download, FileJson } from "lucide-react";
import type { CsvOpenedDocument } from "../types";
import { SpreadsheetGrid } from "../xlsx/SpreadsheetGrid";
import { downloadTextFile, sheetToCsv, sheetToJson } from "../xlsx/spreadsheetUtils";

export function CsvViewer({ document }: { document: CsvOpenedDocument }) {
  const baseName = document.fileName.replace(/\.[^.]+$/, "");

  return (
    <div className="office-viewer">
      <div className="viewer-controls">
        <p className="viewer-note">
          Delimiter: {document.delimiter === "\t" ? "Tab" : document.delimiter} ({document.delimiterConfidence} confidence)
        </p>
        <div className="button-row">
          <button
            type="button"
            className="secondary-action"
            onClick={() => downloadTextFile(`${baseName}.csv`, sheetToCsv(document.sheet), "text/csv")}
          >
            <Download aria-hidden="true" size={17} />
            Export CSV
          </button>
          <button
            type="button"
            className="secondary-action"
            onClick={() =>
              downloadTextFile(`${baseName}.json`, sheetToJson(document.sheet), "application/json")
            }
          >
            <FileJson aria-hidden="true" size={17} />
            Export JSON
          </button>
        </div>
      </div>
      <SpreadsheetGrid sheet={document.sheet} />
    </div>
  );
}

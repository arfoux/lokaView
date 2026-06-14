import { Download, FileJson } from "lucide-react";
import { useState } from "react";
import type { XlsxOpenedDocument } from "../types";
import { SpreadsheetGrid } from "./SpreadsheetGrid";
import { downloadTextFile, sheetToCsv, sheetToJson } from "./spreadsheetUtils";

export function XlsxViewer({ document }: { document: XlsxOpenedDocument }) {
  const [activeSheetName, setActiveSheetName] = useState(document.sheets[0]?.name ?? "");
  const activeSheet = document.sheets.find((sheet) => sheet.name === activeSheetName) ?? document.sheets[0];

  if (!activeSheet) {
    return (
      <section className="viewer-card">
        <h2>No worksheet content found</h2>
        <p>This workbook opened, but there is no sheet data to preview.</p>
      </section>
    );
  }

  const baseName = document.fileName.replace(/\.[^.]+$/, "");

  return (
    <div className="office-viewer">
      <div className="viewer-controls">
        <div className="tab-list" role="tablist" aria-label="Workbook sheets">
          {document.sheets.map((sheet) => (
            <button
              key={sheet.name}
              type="button"
              role="tab"
              aria-selected={sheet.name === activeSheet.name}
              className={sheet.name === activeSheet.name ? "is-active" : undefined}
              onClick={() => setActiveSheetName(sheet.name)}
            >
              {sheet.name}
            </button>
          ))}
        </div>
        <div className="button-row">
          <button
            type="button"
            className="secondary-action"
            onClick={() =>
              downloadTextFile(`${baseName}-${activeSheet.name}.csv`, sheetToCsv(activeSheet), "text/csv")
            }
          >
            <Download aria-hidden="true" size={17} />
            Export CSV
          </button>
          <button
            type="button"
            className="secondary-action"
            onClick={() =>
              downloadTextFile(
                `${baseName}-${activeSheet.name}.json`,
                sheetToJson(activeSheet),
                "application/json",
              )
            }
          >
            <FileJson aria-hidden="true" size={17} />
            Export JSON
          </button>
        </div>
      </div>
      <SpreadsheetGrid sheet={activeSheet} />
    </div>
  );
}

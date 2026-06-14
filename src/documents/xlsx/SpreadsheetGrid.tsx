import { Search } from "lucide-react";
import { useMemo, useState } from "react";
import type { SpreadsheetSheet } from "../types";
import { columnLabel, getMergeAt, isCoveredByMerge } from "./spreadsheetUtils";

interface SpreadsheetGridProps {
  sheet: SpreadsheetSheet;
}

export function SpreadsheetGrid({ sheet }: SpreadsheetGridProps) {
  const [query, setQuery] = useState("");
  const normalizedQuery = query.trim().toLowerCase();

  const matches = useMemo(() => {
    if (!normalizedQuery) {
      return new Set<string>();
    }

    const found = new Set<string>();
    sheet.rows.forEach((row, rowIndex) => {
      row.forEach((value, columnIndex) => {
        if (value.toLowerCase().includes(normalizedQuery)) {
          found.add(`${rowIndex}:${columnIndex}`);
        }
      });
    });
    return found;
  }, [normalizedQuery, sheet]);

  const columnCount = Math.max(sheet.rows[0]?.length ?? 0, Math.min(sheet.columnCount, 80));

  return (
    <div className="sheet-viewer">
      <div className="sheet-tools">
        <label className="search-field">
          <Search aria-hidden="true" size={17} />
          <span className="sr-only">Search cells</span>
          <input
            type="search"
            value={query}
            onChange={(event) => setQuery(event.currentTarget.value)}
            placeholder="Search cells"
          />
        </label>
        <span>
          {sheet.rowCount.toLocaleString()} rows, {sheet.columnCount.toLocaleString()} columns
        </span>
      </div>
      {(sheet.truncatedRows || sheet.truncatedColumns) && (
        <p className="viewer-note">
          Preview capped for responsiveness. The original file remains available for download.
        </p>
      )}
      <div className="grid-scroll" role="region" aria-label={`${sheet.name} worksheet preview`}>
        <table className="spreadsheet-grid">
          <thead>
            <tr>
              <th className="corner-cell" aria-label="Rows and columns" />
              {Array.from({ length: columnCount }, (_, index) => (
                <th key={index} scope="col">
                  {columnLabel(index)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {sheet.rows.map((row, rowIndex) => (
              <tr key={rowIndex}>
                <th scope="row">{rowIndex + 1}</th>
                {Array.from({ length: columnCount }, (_, columnIndex) => {
                  if (isCoveredByMerge(sheet.merges, rowIndex, columnIndex)) {
                    return null;
                  }

                  const merge = getMergeAt(sheet.merges, rowIndex, columnIndex);
                  const value = row[columnIndex] ?? "";
                  const key = `${rowIndex}:${columnIndex}`;
                  const style = sheet.styles[key];

                  return (
                    <td
                      key={columnIndex}
                      colSpan={merge?.columnSpan}
                      rowSpan={merge?.rowSpan}
                      className={matches.has(key) ? "is-match" : undefined}
                      style={{
                        backgroundColor: style?.backgroundColor,
                        color: style?.color,
                        fontWeight: style?.fontWeight,
                        fontStyle: style?.fontStyle,
                        textDecoration: style?.textDecoration,
                        textAlign: style?.horizontalAlign,
                      }}
                    >
                      {value}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

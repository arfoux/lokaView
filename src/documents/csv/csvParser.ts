import { FILE_LIMITS } from "../limits";

export interface ParsedCsv {
  rows: string[][];
  delimiter: string;
  confidence: "high" | "medium" | "low";
  rowCount: number;
  columnCount: number;
  truncatedRows: boolean;
  truncatedColumns: boolean;
}

const DELIMITERS = [",", ";", "\t", "|"];

function scoreDelimiter(text: string, delimiter: string): number {
  const lines = text.split(/\r?\n/).filter((line) => line.trim().length > 0).slice(0, 20);
  if (lines.length === 0) {
    return 0;
  }

  const counts = lines.map((line) => line.split(delimiter).length);
  const average = counts.reduce((sum, count) => sum + count, 0) / counts.length;
  const variance =
    counts.reduce((sum, count) => sum + Math.abs(count - average), 0) / Math.max(counts.length, 1);

  return average > 1 ? average * 4 - variance : 0;
}

function detectDelimiter(text: string): { delimiter: string; confidence: ParsedCsv["confidence"] } {
  const scored = DELIMITERS.map((delimiter) => ({
    delimiter,
    score: scoreDelimiter(text, delimiter),
  })).sort((a, b) => b.score - a.score);

  const best = scored[0] ?? { delimiter: ",", score: 0 };
  const second = scored[1]?.score ?? 0;

  if (best.score <= 1) {
    return { delimiter: ",", confidence: "low" };
  }

  if (best.score - second < 2) {
    return { delimiter: best.delimiter, confidence: "medium" };
  }

  return { delimiter: best.delimiter, confidence: "high" };
}

export function parseCsv(text: string, forcedDelimiter?: string): ParsedCsv {
  const { delimiter, confidence } = forcedDelimiter
    ? { delimiter: forcedDelimiter, confidence: "high" as const }
    : detectDelimiter(text);

  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;
  let rowCount = 0;
  let columnCount = 0;
  let truncatedColumns = false;

  const pushField = () => {
    if (row.length < FILE_LIMITS.csvPreviewColumns) {
      row.push(field);
    } else {
      truncatedColumns = true;
    }
    field = "";
  };

  const pushRow = () => {
    pushField();
    columnCount = Math.max(columnCount, row.length);
    rowCount += 1;

    if (rows.length < FILE_LIMITS.csvPreviewRows) {
      rows.push(row);
    }

    row = [];
  };

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (char === "\"") {
      if (inQuotes && next === "\"") {
        field += "\"";
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (!inQuotes && char === delimiter) {
      pushField();
      continue;
    }

    if (!inQuotes && (char === "\n" || char === "\r")) {
      if (char === "\r" && next === "\n") {
        index += 1;
      }
      pushRow();
      continue;
    }

    field += char;
  }

  if (field.length > 0 || row.length > 0) {
    pushRow();
  }

  return {
    rows,
    delimiter,
    confidence,
    rowCount,
    columnCount,
    truncatedRows: rowCount > rows.length,
    truncatedColumns,
  };
}

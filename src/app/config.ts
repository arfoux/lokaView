import type { DocumentKind } from "../documents/types";

export const APP_NAME = "Local Office Viewer";

export const ACCEPTED_FILE_EXTENSIONS = [
  ".pdf",
  ".docx",
  ".xlsx",
  ".pptx",
  ".csv",
  ".doc",
  ".xls",
  ".ppt",
].join(",");

export const KIND_LABELS: Record<DocumentKind, string> = {
  pdf: "PDF",
  docx: "Word",
  xlsx: "Excel",
  pptx: "PowerPoint",
  csv: "CSV",
};

export const PRIVACY_PROMISE =
  "Your documents are processed locally in your browser. Files are not uploaded to a server.";

import { lazy, Suspense } from "react";
import { LoadingState } from "./LoadingState";
import type { OpenedDocument } from "../documents/types";

const PdfViewer = lazy(() => import("../documents/pdf/PdfViewer").then((module) => ({ default: module.PdfViewer })));
const DocxViewer = lazy(() => import("../documents/docx/DocxViewer").then((module) => ({ default: module.DocxViewer })));
const XlsxViewer = lazy(() => import("../documents/xlsx/XlsxViewer").then((module) => ({ default: module.XlsxViewer })));
const PptxViewer = lazy(() => import("../documents/pptx/PptxViewer").then((module) => ({ default: module.PptxViewer })));
const CsvViewer = lazy(() => import("../documents/csv/CsvViewer").then((module) => ({ default: module.CsvViewer })));

export function ViewerHost({ document }: { document: OpenedDocument }) {
  return (
    <Suspense
      fallback={<LoadingState title="Loading viewer" message="Preparing the local renderer..." />}
    >
      {document.kind === "pdf" && <PdfViewer document={document} />}
      {document.kind === "docx" && <DocxViewer document={document} />}
      {document.kind === "xlsx" && <XlsxViewer document={document} />}
      {document.kind === "pptx" && <PptxViewer document={document} />}
      {document.kind === "csv" && <CsvViewer document={document} />}
    </Suspense>
  );
}

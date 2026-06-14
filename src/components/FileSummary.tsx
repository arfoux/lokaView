import { KIND_LABELS } from "../app/config";
import { formatBytes } from "../documents/limits";
import type { OpenedDocument } from "../documents/types";

export function FileSummary({ document }: { document: OpenedDocument }) {
  return (
    <div className="file-summary">
      <strong>{document.fileName}</strong>
      <span>{KIND_LABELS[document.kind]}</span>
      <span>{formatBytes(document.fileSize)}</span>
      <span>Detected by {document.detectedBy}</span>
    </div>
  );
}

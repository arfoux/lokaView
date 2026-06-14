import { Download, FolderOpen, ShieldCheck, X } from "lucide-react";
import { FileSummary } from "./FileSummary";
import type { OpenedDocument } from "../documents/types";

interface ViewerToolbarProps {
  document: OpenedDocument;
  onOpenAnother: () => void;
  onClose: () => void;
}

export function ViewerToolbar({ document, onOpenAnother, onClose }: ViewerToolbarProps) {
  return (
    <div className="viewer-toolbar">
      <FileSummary document={document} />
      <div className="toolbar-actions">
        <span className="toolbar-status">
          <ShieldCheck aria-hidden="true" size={16} />
          Local only
        </span>
        <a className="icon-button" href={document.originalUrl} download={document.fileName} title="Download original file">
          <Download aria-hidden="true" size={18} />
          <span className="sr-only">Download original file</span>
        </a>
        <button type="button" className="icon-button" onClick={onOpenAnother} title="Open another file">
          <FolderOpen aria-hidden="true" size={18} />
          <span className="sr-only">Open another file</span>
        </button>
        <button type="button" className="icon-button" onClick={onClose} title="Close document">
          <X aria-hidden="true" size={18} />
          <span className="sr-only">Close document</span>
        </button>
      </div>
    </div>
  );
}

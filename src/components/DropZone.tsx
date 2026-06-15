import { FileUp, FolderOpen } from "lucide-react";
import { useState } from "react";
import { PRIVACY_PROMISE } from "../app/config";

interface DropZoneProps {
  onFileSelected: (file: File) => void;
  onChooseFile: () => void;
  compact?: boolean;
}

export function DropZone({ onFileSelected, onChooseFile, compact = false }: DropZoneProps) {
  const [isDragOver, setIsDragOver] = useState(false);

  const handleFiles = (files: FileList | null) => {
    const file = files?.[0];
    if (file) {
      onFileSelected(file);
    }
  };

  return (
    <section
      className={`drop-zone${isDragOver ? " is-drag-over" : ""}${compact ? " is-compact" : ""}`}
      tabIndex={0}
      role="button"
      aria-label="Choose a local document"
      onClick={onChooseFile}
      onKeyDown={(event) => {
        if (event.target !== event.currentTarget) {
          return;
        }

        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          onChooseFile();
        }
      }}
      onDragOver={(event) => {
        event.preventDefault();
        setIsDragOver(true);
      }}
      onDragLeave={() => setIsDragOver(false)}
      onDrop={(event) => {
        event.preventDefault();
        setIsDragOver(false);
        handleFiles(event.dataTransfer.files);
      }}
    >
      <div className="drop-zone-icon" aria-hidden="true">
        <FileUp size={compact ? 22 : 30} />
      </div>
      <div className="drop-zone-copy">
        <h1>{compact ? "Open another document" : "Open PDF, Word, Excel, and PowerPoint files directly in your browser."}</h1>
        {!compact && <p>{PRIVACY_PROMISE}</p>}
        <span className="supported-formats">PDF, DOCX, XLSX, PPTX, CSV</span>
      </div>
      <button
        type="button"
        className="primary-action"
        onClick={(event) => {
          event.stopPropagation();
          onChooseFile();
        }}
      >
        <FolderOpen aria-hidden="true" size={18} />
        Choose file
      </button>
    </section>
  );
}

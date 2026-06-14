import { AlertTriangle, FolderOpen } from "lucide-react";
import { DocumentError } from "../documents/types";

interface ErrorStateProps {
  error: unknown;
  onOpenAnother: () => void;
}

function getErrorCopy(error: unknown) {
  if (error instanceof DocumentError) {
    return {
      title: error.title,
      message: error.message,
    };
  }

  return {
    title: "Something went wrong",
    message: "The document could not be opened. Try another file or a simpler export.",
  };
}

export function ErrorState({ error, onOpenAnother }: ErrorStateProps) {
  const copy = getErrorCopy(error);

  return (
    <section className="state-panel error-panel" role="alert">
      <AlertTriangle aria-hidden="true" size={30} />
      <h2>{copy.title}</h2>
      <p>{copy.message}</p>
      <button type="button" className="secondary-action" onClick={onOpenAnother}>
        <FolderOpen aria-hidden="true" size={18} />
        Choose another file
      </button>
    </section>
  );
}

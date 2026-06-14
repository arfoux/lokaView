import { LoaderCircle } from "lucide-react";
import type { ReadProgress } from "../documents/types";

interface LoadingStateProps {
  title: string;
  message?: string;
  progress?: ReadProgress;
}

export function LoadingState({ title, message, progress }: LoadingStateProps) {
  const percent = progress ? Math.round(progress.ratio * 100) : undefined;

  return (
    <section className="state-panel" aria-live="polite">
      <LoaderCircle className="spin" aria-hidden="true" size={28} />
      <h2>{title}</h2>
      {message && <p>{message}</p>}
      {progress && (
        <div className="progress" aria-label={`Reading file ${percent ?? 0}% complete`}>
          <span style={{ width: `${percent ?? 0}%` }} />
        </div>
      )}
    </section>
  );
}

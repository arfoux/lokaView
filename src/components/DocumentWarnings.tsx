import { Info } from "lucide-react";

export function DocumentWarnings({ warnings }: { warnings: readonly string[] }) {
  if (warnings.length === 0) {
    return null;
  }

  return (
    <div className="document-warnings" role="status">
      <Info aria-hidden="true" size={17} />
      <div>
        {warnings.map((warning) => (
          <p key={warning}>{warning}</p>
        ))}
      </div>
    </div>
  );
}

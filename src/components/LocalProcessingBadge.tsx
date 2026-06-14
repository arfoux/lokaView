import { ShieldCheck } from "lucide-react";

export function LocalProcessingBadge() {
  return (
    <span className="local-badge" title="Selected document bytes stay in this browser session.">
      <ShieldCheck aria-hidden="true" size={16} />
      Local processing only
    </span>
  );
}

import { FileText } from "lucide-react";
import { APP_NAME } from "../app/config";
import { LocalProcessingBadge } from "./LocalProcessingBadge";

export function AppHeader() {
  return (
    <header className="app-header">
      <div className="brand">
        <span className="brand-mark" aria-hidden="true">
          <FileText size={22} />
        </span>
        <div>
          <p className="brand-name">{APP_NAME}</p>
          <p className="brand-subtitle">Browser-local document viewer</p>
        </div>
      </div>
      <LocalProcessingBadge />
    </header>
  );
}

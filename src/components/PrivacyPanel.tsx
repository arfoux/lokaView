import { LockKeyhole } from "lucide-react";
import { PRIVACY_PROMISE } from "../app/config";

export function PrivacyPanel() {
  return (
    <details className="privacy-panel">
      <summary>
        <LockKeyhole aria-hidden="true" size={18} />
        Privacy model
      </summary>
      <div>
        <p>{PRIVACY_PROMISE}</p>
        <p>
          The app reads selected files with browser File APIs and hands the bytes to local viewer
          adapters. There is no login, database, document upload endpoint, or remote conversion
          service in the core viewer.
        </p>
        <p>
          Internal /url/... links are optional and are fetched through the app proxy before being
          rendered locally in this browser.
        </p>
      </div>
    </details>
  );
}

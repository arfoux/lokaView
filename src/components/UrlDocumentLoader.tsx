import { Link2 } from "lucide-react";
import { useState } from "react";

interface UrlDocumentLoaderProps {
  onOpenUrl: (internalUrl: string) => void;
}

export function UrlDocumentLoader({ onOpenUrl }: UrlDocumentLoaderProps) {
  const [value, setValue] = useState("");
  const [error, setError] = useState("");

  return (
    <form
      className="url-loader"
      onSubmit={(event) => {
        event.preventDefault();
        const nextValue = value.trim();

        if (!nextValue) {
          setError("Enter an internal document URL.");
          return;
        }

        if (!nextValue.startsWith("/url/") && !/^https?:\/\//i.test(nextValue)) {
          setError("Use an internal /url/... path or an HTTP(S) document URL.");
          return;
        }

        setError("");
        onOpenUrl(nextValue);
      }}
    >
      <label htmlFor="url-document-input">
        <Link2 aria-hidden="true" size={18} />
        Open from internal URL
      </label>
      <div className="url-loader-row">
        <input
          id="url-document-input"
          type="text"
          inputMode="url"
          value={value}
          placeholder="https://calibre-ebook.com/downloads/demos/demo.docx"
          onChange={(event) => {
            setValue(event.currentTarget.value);
            setError("");
          }}
        />
        <button type="submit" className="secondary-action">Open URL</button>
      </div>
      {error && <p role="alert">{error}</p>}
    </form>
  );
}

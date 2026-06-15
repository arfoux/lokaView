import { FolderOpen, ShieldCheck } from "lucide-react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { ACCEPTED_FILE_EXTENSIONS, PRIVACY_PROMISE } from "./config";
import { AppHeader } from "../components/AppHeader";
import { DocumentWarnings } from "../components/DocumentWarnings";
import { DropZone } from "../components/DropZone";
import { ErrorBoundary } from "../components/ErrorBoundary";
import { ErrorState } from "../components/ErrorState";
import { LoadingState } from "../components/LoadingState";
import { PrivacyPanel } from "../components/PrivacyPanel";
import { UrlDocumentLoader } from "../components/UrlDocumentLoader";
import { ViewerHost } from "../components/ViewerHost";
import { ViewerToolbar } from "../components/ViewerToolbar";
import { validateFileSize } from "../documents/limits";
import { openLocalDocument } from "../documents/openLocalDocument";
import { DocumentSession } from "../documents/session";
import { DocumentError } from "../documents/types";
import type { OpenedDocument, ReadProgress } from "../documents/types";

type AppState =
  | { status: "idle" }
  | { status: "confirm-large"; file: File; title: string; message: string }
  | { status: "fetching-url"; url: string }
  | { status: "reading"; fileName: string; progress?: ReadProgress }
  | { status: "parsing"; fileName: string }
  | { status: "ready"; document: OpenedDocument }
  | { status: "error"; error: unknown };

function normalizeInternalDocumentUrl(value: string) {
  const source = value.trim();
  let url: URL;

  try {
    url = new URL(source, window.location.origin);
  } catch {
    throw new DocumentError({
      code: "read-failed",
      title: "Invalid document URL",
      message: "Use an internal /url/... document path or paste a full HTTP(S) document URL.",
    });
  }

  if (url.origin === window.location.origin && url.pathname.startsWith("/url/")) {
    return `${url.pathname}${url.search}`;
  }

  if (/^https?:\/\//i.test(source) && (url.protocol === "https:" || url.protocol === "http:")) {
    return `/url/${url.href}`;
  }

  throw new DocumentError({
    code: "read-failed",
    title: "Invalid document URL",
    message: "Only /url/... paths or HTTP(S) document URLs can be opened.",
  });
}

function getFileNameFromContentDisposition(header: string | null) {
  const filename = /filename="([^"]+)"/i.exec(header ?? "")?.[1] ?? /filename=([^;]+)/i.exec(header ?? "")?.[1];
  return filename?.trim();
}

function getFileNameFromUrl(internalUrl: string) {
  const pathname = new URL(internalUrl, window.location.origin).pathname;
  const lastSegment = pathname.split("/").filter(Boolean).at(-1);

  if (!lastSegment) {
    return "document.bin";
  }

  try {
    return decodeURIComponent(lastSegment);
  } catch {
    return lastSegment;
  }
}

export function App() {
  const [state, setState] = useState<AppState>({ status: "idle" });
  const session = useMemo(() => new DocumentSession(), []);
  const abortControllerRef = useRef<AbortController | undefined>(undefined);
  const autoOpenedUrlRef = useRef("");
  const hiddenInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    return () => {
      abortControllerRef.current?.abort();
      session.clear();
    };
  }, [session]);

  const openPicker = () => {
    if (!hiddenInputRef.current) {
      return;
    }

    hiddenInputRef.current.value = "";
    hiddenInputRef.current.click();
  };

  const closeDocument = () => {
    abortControllerRef.current?.abort();
    session.clear();
    setState({ status: "idle" });
  };

  const openFile = useCallback(async (file: File, allowLargeFile = false) => {
    abortControllerRef.current?.abort();
    abortControllerRef.current = undefined;
    session.clear();

    const sizeValidation = validateFileSize(file.size);

    if (sizeValidation.status === "block") {
      setState({
        status: "error",
        error: new DocumentError({
          code: "file-too-large",
          title: sizeValidation.title ?? "File is too large",
          message: sizeValidation.message ?? "This file is too large for safe local processing.",
        }),
      });
      return;
    }

    if (sizeValidation.status === "warn" && !allowLargeFile) {
      setState({
        status: "confirm-large",
        file,
        title: sizeValidation.title ?? "Large local file",
        message: sizeValidation.message ?? "This file is large enough to require confirmation.",
      });
      return;
    }

    const abortController = new AbortController();
    abortControllerRef.current = abortController;
    setState({ status: "reading", fileName: file.name });

    try {
      const document = await openLocalDocument(file, {
        allowLargeFile: true,
        signal: abortController.signal,
        onProgress: (progress) => setState({ status: "reading", fileName: file.name, progress }),
      });

      if (abortController.signal.aborted) {
        document.dispose();
        return;
      }

      setState({ status: "parsing", fileName: file.name });
      session.setActive(document);
      setState({ status: "ready", document });
    } catch (error) {
      if (abortController.signal.aborted) {
        return;
      }

      if (import.meta.env.DEV) {
        console.error("Document open failed", error);
      }

      setState({ status: "error", error });
    }
  }, [session]);

  const handleFileInput = (files: FileList | null) => {
    const file = files?.[0];
    if (file) {
      void openFile(file);
    }

    if (hiddenInputRef.current) {
      hiddenInputRef.current.value = "";
    }
  };

  const openUrlDocument = useCallback(async (value: string) => {
    let internalUrl: string;

    try {
      internalUrl = normalizeInternalDocumentUrl(value);
    } catch (error) {
      setState({ status: "error", error });
      return;
    }

    abortControllerRef.current?.abort();
    const abortController = new AbortController();
    abortControllerRef.current = abortController;
    session.clear();
    setState({ status: "fetching-url", url: internalUrl });

    try {
      const response = await fetch(internalUrl, {
        credentials: "same-origin",
        signal: abortController.signal,
      });

      if (!response.ok) {
        throw new DocumentError({
          code: "read-failed",
          title: "URL document could not be loaded",
          message: response.status === 404
            ? "The internal document URL was not found or uses an unknown alias."
            : "The document proxy could not fetch that file. Check the path and try again.",
        });
      }

      const blob = await response.blob();

      if (abortController.signal.aborted) {
        return;
      }

      if (blob.size === 0) {
        throw new DocumentError({
          code: "read-failed",
          title: "URL response was empty",
          message: "The internal URL responded successfully, but it did not contain a document file.",
        });
      }

      const fileName = getFileNameFromContentDisposition(response.headers.get("Content-Disposition")) ?? getFileNameFromUrl(internalUrl);
      const file = new File([blob], fileName, {
        type: response.headers.get("Content-Type") ?? blob.type,
        lastModified: Date.now(),
      });

      await openFile(file);
    } catch (error) {
      if (abortController.signal.aborted) {
        return;
      }

      if (import.meta.env.DEV) {
        console.error("URL document open failed", error);
      }

      setState({ status: "error", error });
    }
  }, [openFile, session]);

  useEffect(() => {
    if (!window.location.pathname.startsWith("/url/")) {
      return;
    }

    const internalUrl = `${window.location.pathname}${window.location.search}`;

    if (autoOpenedUrlRef.current === internalUrl) {
      return;
    }

    autoOpenedUrlRef.current = internalUrl;
    void openUrlDocument(internalUrl);
  }, [openUrlDocument]);

  return (
    <div className="app-shell">
      <input
        ref={hiddenInputRef}
        className="sr-only"
        type="file"
        accept={ACCEPTED_FILE_EXTENSIONS}
        onChange={(event) => handleFileInput(event.currentTarget.files)}
        aria-hidden="true"
        tabIndex={-1}
      />
      <AppHeader />
      <main className={state.status === "ready" ? "viewer-main" : "landing-main"}>
        {state.status === "idle" && (
          <>
            <DropZone onFileSelected={(file) => void openFile(file)} onChooseFile={openPicker} />
            <UrlDocumentLoader onOpenUrl={(url) => void openUrlDocument(url)} />
            <div className="landing-support">
              <div>
                <ShieldCheck aria-hidden="true" size={20} />
                <strong>Your file stays on your device.</strong>
                <p>{PRIVACY_PROMISE}</p>
              </div>
              <PrivacyPanel />
            </div>
          </>
        )}

        {state.status === "confirm-large" && (
          <section className="state-panel">
            <h2>{state.title}</h2>
            <p>{state.message}</p>
            <div className="button-row">
              <button
                type="button"
                className="primary-action"
                onClick={() => void openFile(state.file, true)}
              >
                <FolderOpen aria-hidden="true" size={18} />
                Open locally
              </button>
              <button type="button" className="secondary-action" onClick={openPicker}>
                Choose another file
              </button>
            </div>
          </section>
        )}

        {state.status === "reading" && (
          <LoadingState
            title="Reading local file"
            message={`Preparing ${state.fileName} without uploading it.`}
            progress={state.progress}
          />
        )}

        {state.status === "fetching-url" && (
          <LoadingState title="Loading URL document" message={`Fetching ${state.url} through the document proxy.`} />
        )}

        {state.status === "parsing" && (
          <LoadingState title="Opening document" message={`Building a local preview for ${state.fileName}.`} />
        )}

        {state.status === "error" && <ErrorState error={state.error} onOpenAnother={openPicker} />}

        {state.status === "ready" && (
          <section className="viewer-shell" aria-label="Document viewer">
            <ViewerToolbar document={state.document} onOpenAnother={openPicker} onClose={closeDocument} />
            <DocumentWarnings warnings={state.document.warnings} />
            <ErrorBoundary onOpenAnother={openPicker}>
              <ViewerHost document={state.document} />
            </ErrorBoundary>
          </section>
        )}
      </main>
    </div>
  );
}

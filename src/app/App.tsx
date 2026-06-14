import { FolderOpen, ShieldCheck } from "lucide-react";
import { useEffect, useMemo, useRef, useState } from "react";
import { ACCEPTED_FILE_EXTENSIONS, PRIVACY_PROMISE } from "./config";
import { AppHeader } from "../components/AppHeader";
import { DocumentWarnings } from "../components/DocumentWarnings";
import { DropZone } from "../components/DropZone";
import { ErrorBoundary } from "../components/ErrorBoundary";
import { ErrorState } from "../components/ErrorState";
import { LoadingState } from "../components/LoadingState";
import { PrivacyPanel } from "../components/PrivacyPanel";
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
  | { status: "reading"; fileName: string; progress?: ReadProgress }
  | { status: "parsing"; fileName: string }
  | { status: "ready"; document: OpenedDocument }
  | { status: "error"; error: unknown };

export function App() {
  const [state, setState] = useState<AppState>({ status: "idle" });
  const session = useMemo(() => new DocumentSession(), []);
  const abortControllerRef = useRef<AbortController | undefined>(undefined);
  const hiddenInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    return () => {
      abortControllerRef.current?.abort();
      session.clear();
    };
  }, [session]);

  const openPicker = () => hiddenInputRef.current?.click();

  const closeDocument = () => {
    abortControllerRef.current?.abort();
    session.clear();
    setState({ status: "idle" });
  };

  const openFile = async (file: File, allowLargeFile = false) => {
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

    abortControllerRef.current?.abort();
    const abortController = new AbortController();
    abortControllerRef.current = abortController;
    session.clear();
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
  };

  const handleFileInput = (files: FileList | null) => {
    const file = files?.[0];
    if (file) {
      void openFile(file);
    }

    if (hiddenInputRef.current) {
      hiddenInputRef.current.value = "";
    }
  };

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
            <DropZone onFileSelected={(file) => void openFile(file)} />
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

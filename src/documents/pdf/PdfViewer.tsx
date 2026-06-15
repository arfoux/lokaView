import { ChevronLeft, ChevronRight, Maximize2, Minus, Plus, RotateCw } from "lucide-react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import * as pdfjsLib from "pdfjs-dist";
import type { PDFDocumentProxy, PDFPageProxy, RenderTask } from "pdfjs-dist";
import type { PdfOpenedDocument } from "../types";

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
  "pdfjs-dist/build/pdf.worker.min.mjs",
  import.meta.url,
).toString();

const MIN_PDF_OUTPUT_SCALE = 2;
const MAX_PDF_OUTPUT_SCALE = 3;
const MAX_PDF_CANVAS_PIXELS = 16_777_216;

function getPdfOutputScale(viewport: { width: number; height: number }): number {
  const deviceScale = window.devicePixelRatio || 1;
  const preferredScale = Math.min(MAX_PDF_OUTPUT_SCALE, Math.max(MIN_PDF_OUTPUT_SCALE, deviceScale));
  const viewportPixels = viewport.width * viewport.height;

  if (viewportPixels <= 0) {
    return preferredScale;
  }

  const maxSafeScale = Math.sqrt(MAX_PDF_CANVAS_PIXELS / viewportPixels);
  return Math.max(1, Math.min(preferredScale, maxSafeScale));
}

interface PdfPageProps {
  pdf: PDFDocumentProxy;
  pageNumber: number;
  scale: number;
  rotation: number;
  renderPixelRatio: number;
  onVisible: (pageNumber: number) => void;
}

function PdfPage({ pdf, pageNumber, scale, rotation, renderPixelRatio, onVisible }: PdfPageProps) {
  const wrapperRef = useRef<HTMLDivElement>(null);
  const pageViewportRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const textLayerRef = useRef<HTMLDivElement>(null);
  const [isVisible, setIsVisible] = useState(pageNumber <= 2);
  const [error, setError] = useState("");

  useEffect(() => {
    const element = wrapperRef.current;
    if (!element) {
      return;
    }

    const observer = new IntersectionObserver(
      (entries) => {
        const entry = entries[0];
        if (entry?.isIntersecting) {
          setIsVisible(true);
          onVisible(pageNumber);
        }
      },
      { rootMargin: "600px 0px" },
    );

    observer.observe(element);
    return () => observer.disconnect();
  }, [onVisible, pageNumber]);

  useEffect(() => {
    if (!isVisible) {
      return;
    }

    let canceled = false;
    let renderTask: RenderTask | undefined;
    let textLayer: { cancel: () => void; render: () => Promise<void> } | undefined;
    let activeTextLayerContainer: HTMLDivElement | undefined;
    let loadedPage: PDFPageProxy | undefined;

    async function renderPage() {
      const pageViewport = pageViewportRef.current;
      const canvas = canvasRef.current;
      const textLayerContainer = textLayerRef.current;
      if (!pageViewport || !canvas || !textLayerContainer) {
        return;
      }

      activeTextLayerContainer = textLayerContainer;

      try {
        loadedPage = await pdf.getPage(pageNumber);
        if (canceled) {
          return;
        }

        const viewport = loadedPage.getViewport({ scale, rotation });
        const context = canvas.getContext("2d");
        if (!context) {
          return;
        }

        pageViewport.style.width = `${viewport.width}px`;
        pageViewport.style.height = `${viewport.height}px`;
        pageViewport.style.setProperty("--total-scale-factor", `${viewport.scale}`);

        const outputScale = getPdfOutputScale(viewport);
        canvas.width = Math.ceil(viewport.width * outputScale);
        canvas.height = Math.ceil(viewport.height * outputScale);
        canvas.style.width = `${viewport.width}px`;
        canvas.style.height = `${viewport.height}px`;
        canvas.dataset.outputScale = outputScale.toFixed(2);

        textLayerContainer.replaceChildren();
        textLayerContainer.style.width = `${viewport.width}px`;
        textLayerContainer.style.height = `${viewport.height}px`;

        textLayer = new pdfjsLib.TextLayer({
          textContentSource: loadedPage.streamTextContent({
            includeMarkedContent: true,
            disableNormalization: true,
          }),
          container: textLayerContainer,
          viewport,
        });

        renderTask = loadedPage.render({
          canvas,
          canvasContext: context,
          viewport,
          transform: outputScale !== 1 ? [outputScale, 0, 0, outputScale, 0, 0] : undefined,
        });

        await Promise.all([
          renderTask.promise,
          textLayer.render().catch((textLayerError: unknown) => {
            if (import.meta.env.DEV && !canceled) {
              console.warn("PDF text layer render failed", textLayerError);
            }
          }),
        ]);
      } catch (renderError) {
        if (!canceled && renderError instanceof Error && renderError.name !== "RenderingCancelledException") {
          setError("This page could not be rendered.");
        }
      }
    }

    void renderPage();

    return () => {
      canceled = true;
      renderTask?.cancel();
      textLayer?.cancel();
      activeTextLayerContainer?.replaceChildren();
      loadedPage?.cleanup();
    };
  }, [isVisible, pageNumber, pdf, renderPixelRatio, rotation, scale]);

  return (
    <div ref={wrapperRef} className="pdf-page-shell" data-page={pageNumber}>
      <span className="pdf-page-label">Page {pageNumber}</span>
      {error ? (
        <p className="viewer-error">{error}</p>
      ) : (
        <div ref={pageViewportRef} className="pdf-page-viewport">
          <canvas ref={canvasRef} />
          <div ref={textLayerRef} className="textLayer pdf-text-layer" />
        </div>
      )}
    </div>
  );
}

function getPdfErrorMessage(error: unknown): string {
  if (error instanceof Error && /password/i.test(error.name + error.message)) {
    return "This PDF is password protected. Encrypted PDFs are not previewed yet.";
  }

  return "This PDF could not be opened. It may be corrupted, encrypted, or use unsupported features.";
}

export function PdfViewer({ document }: { document: PdfOpenedDocument }) {
  const [pdf, setPdf] = useState<PDFDocumentProxy>();
  const [pageCount, setPageCount] = useState(0);
  const [currentPage, setCurrentPage] = useState(1);
  const [scale, setScale] = useState(1.1);
  const [fitMode, setFitMode] = useState(() => window.innerWidth < 760);
  const [rotation, setRotation] = useState(0);
  const [renderPixelRatio, setRenderPixelRatio] = useState(() => window.devicePixelRatio || 1);
  const [loadingPercent, setLoadingPercent] = useState(0);
  const [error, setError] = useState("");
  const shellRef = useRef<HTMLDivElement>(null);
  const pagesRef = useRef<Array<HTMLDivElement | null>>([]);

  useEffect(() => {
    let canceled = false;
    const loadingTask = pdfjsLib.getDocument({
      data: new Uint8Array(document.arrayBuffer.slice(0)),
      useSystemFonts: true,
      stopAtErrors: false,
    });

    loadingTask.onProgress = (progress: { loaded: number; total: number }) => {
      if (progress.total > 0) {
        setLoadingPercent(Math.round((progress.loaded / progress.total) * 100));
      }
    };

    loadingTask.promise
      .then((loadedPdf) => {
        if (canceled) {
          void loadedPdf.destroy();
          return;
        }
        setPdf(loadedPdf);
        setPageCount(loadedPdf.numPages);
      })
      .catch((loadError: unknown) => {
        if (!canceled) {
          if (import.meta.env.DEV) {
            console.error("PDF load failed", loadError);
          }
          setError(getPdfErrorMessage(loadError));
        }
      });

    return () => {
      canceled = true;
      void loadingTask.destroy();
    };
  }, [document]);

  useEffect(() => {
    let mediaQuery: MediaQueryList | undefined;

    const syncPixelRatio = () => {
      setRenderPixelRatio(Math.round((window.devicePixelRatio || 1) * 100) / 100);
    };

    const watchPixelRatio = () => {
      mediaQuery?.removeEventListener("change", watchPixelRatio);
      mediaQuery = window.matchMedia(`(resolution: ${window.devicePixelRatio || 1}dppx)`);
      mediaQuery.addEventListener("change", watchPixelRatio);
      syncPixelRatio();
    };

    watchPixelRatio();
    window.addEventListener("resize", syncPixelRatio);
    window.visualViewport?.addEventListener("resize", syncPixelRatio);

    return () => {
      mediaQuery?.removeEventListener("change", watchPixelRatio);
      window.removeEventListener("resize", syncPixelRatio);
      window.visualViewport?.removeEventListener("resize", syncPixelRatio);
    };
  }, []);

  const pageNumbers = useMemo(
    () => Array.from({ length: pageCount }, (_, index) => index + 1),
    [pageCount],
  );

  const goToPage = (pageNumber: number) => {
    const bounded = Math.min(Math.max(pageNumber, 1), pageCount || 1);
    setCurrentPage(bounded);
    pagesRef.current[bounded - 1]?.scrollIntoView({ block: "start", behavior: "smooth" });
  };

  const fitWidth = useCallback(async () => {
    if (!pdf || !shellRef.current) {
      return;
    }

    const page = await pdf.getPage(1);
    const viewport = page.getViewport({ scale: 1, rotation });
    const horizontalPadding = window.innerWidth < 640 ? 28 : 64;
    const availableWidth = Math.max(shellRef.current.clientWidth - horizontalPadding, 260);
    setScale(Math.min(Math.max(availableWidth / viewport.width, 0.35), 2.5));
    page.cleanup();
  }, [pdf, rotation]);

  useEffect(() => {
    if (!fitMode) {
      return;
    }

    void fitWidth();
  }, [fitMode, fitWidth, pageCount]);

  useEffect(() => {
    const shell = shellRef.current;

    if (!shell || !fitMode || typeof ResizeObserver === "undefined") {
      return;
    }

    const observer = new ResizeObserver(() => {
      void fitWidth();
    });

    observer.observe(shell);
    return () => observer.disconnect();
  }, [fitMode, fitWidth]);

  const requestFullscreen = () => {
    void shellRef.current?.requestFullscreen?.();
  };

  if (error) {
    return <p className="viewer-error">{error}</p>;
  }

  return (
    <div ref={shellRef} className="office-viewer pdf-viewer">
      <div className="viewer-controls">
        <div className="segmented-controls" aria-label="PDF page controls">
          <button type="button" className="icon-button" onClick={() => goToPage(currentPage - 1)} disabled={currentPage <= 1} title="Previous page">
            <ChevronLeft aria-hidden="true" size={17} />
            <span className="sr-only">Previous page</span>
          </button>
          <span>
            {currentPage} / {pageCount || "..."}
          </span>
          <button type="button" className="icon-button" onClick={() => goToPage(currentPage + 1)} disabled={currentPage >= pageCount} title="Next page">
            <ChevronRight aria-hidden="true" size={17} />
            <span className="sr-only">Next page</span>
          </button>
        </div>
        <div className="segmented-controls" aria-label="PDF zoom controls">
          <button type="button" className="icon-button" onClick={() => { setFitMode(false); setScale((value) => Math.max(0.35, value - 0.1)); }} title="Zoom out">
            <Minus aria-hidden="true" size={17} />
            <span className="sr-only">Zoom out</span>
          </button>
          <span>{Math.round(scale * 100)}%</span>
          <button type="button" className="icon-button" onClick={() => { setFitMode(false); setScale((value) => Math.min(3, value + 0.1)); }} title="Zoom in">
            <Plus aria-hidden="true" size={17} />
            <span className="sr-only">Zoom in</span>
          </button>
          <button type="button" className={`text-button${fitMode ? " is-active" : ""}`} onClick={() => { setFitMode(true); void fitWidth(); }}>
            Fit width
          </button>
          <button type="button" className="icon-button" onClick={() => setRotation((value) => (value + 90) % 360)} title="Rotate clockwise">
            <RotateCw aria-hidden="true" size={17} />
            <span className="sr-only">Rotate clockwise</span>
          </button>
          <button type="button" className="icon-button" onClick={requestFullscreen} title="Fullscreen">
            <Maximize2 aria-hidden="true" size={17} />
            <span className="sr-only">Fullscreen</span>
          </button>
        </div>
      </div>
      {!pdf && (
        <p className="viewer-note">
          Loading PDF locally{loadingPercent > 0 ? ` (${loadingPercent}%)` : ""}...
        </p>
      )}
      {pdf && (
        <div className="pdf-pages">
          {pageNumbers.map((pageNumber, index) => (
            <div key={pageNumber} ref={(element) => { pagesRef.current[index] = element; }}>
              <PdfPage
                pdf={pdf}
                pageNumber={pageNumber}
                scale={scale}
                rotation={rotation}
                renderPixelRatio={renderPixelRatio}
                onVisible={setCurrentPage}
              />
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

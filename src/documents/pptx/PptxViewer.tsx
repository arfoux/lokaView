import { ChevronLeft, ChevronRight, Maximize2, Minus, Plus, Search } from "lucide-react";
import { useCallback, useEffect, useRef, useState } from "react";
import {
  PptxViewer as BrowserPptxViewer,
  RECOMMENDED_ZIP_LIMITS,
  type SlideHandle,
} from "@aiden0z/pptx-renderer";
import { neutralizeUnsafeDocumentDom } from "../sanitizeDocumentDom";
import type { PptxOpenedDocument } from "../types";

export function PptxViewer({ document }: { document: PptxOpenedDocument }) {
  const containerRef = useRef<HTMLDivElement>(null);
  const shellRef = useRef<HTMLDivElement>(null);
  const viewerRef = useRef<BrowserPptxViewer | undefined>(undefined);
  const thumbnailRefs = useRef<Array<HTMLDivElement | null>>([]);
  const thumbnailHandles = useRef<SlideHandle[]>([]);
  const [status, setStatus] = useState<"loading" | "ready" | "error">("loading");
  const [errorMessage, setErrorMessage] = useState("");
  const [slideCount, setSlideCount] = useState(0);
  const [currentSlide, setCurrentSlide] = useState(0);
  const [zoom, setZoom] = useState(100);
  const [query, setQuery] = useState("");
  const [matchCount, setMatchCount] = useState(0);

  useEffect(() => {
    const container = containerRef.current;
    if (!container) {
      return;
    }

    const abortController = new AbortController();
    setStatus("loading");
    setErrorMessage("");
    container.innerHTML = "";

    BrowserPptxViewer.open(document.arrayBuffer.slice(0), container, {
      renderMode: "slide",
      fitMode: "contain",
      zoomPercent: 100,
      zipLimits: RECOMMENDED_ZIP_LIMITS,
      signal: abortController.signal,
      pdfjs: {
        moduleUrl: new URL("pdfjs-dist/build/pdf.min.mjs", import.meta.url).toString(),
        workerUrl: new URL("pdfjs-dist/build/pdf.worker.min.mjs", import.meta.url).toString(),
      },
      onSlideChange: (index) => setCurrentSlide(index),
      onSlideRendered: (_index, element) => neutralizeUnsafeDocumentDom(element),
      onSlideError: (_index, error) => {
        if (import.meta.env.DEV) {
          console.error("PPTX slide render failed", error);
        }
      },
      onNodeError: (_nodeId, error) => {
        if (import.meta.env.DEV) {
          console.warn("PPTX node render failed", error);
        }
      },
    })
      .then((viewer) => {
        if (abortController.signal.aborted) {
          viewer.destroy();
          return;
        }

        viewerRef.current = viewer;
        setSlideCount(viewer.slideCount);
        setCurrentSlide(viewer.currentSlideIndex);
        setStatus("ready");

        window.setTimeout(() => {
          thumbnailHandles.current.forEach((handle) => handle.dispose());
          thumbnailHandles.current = [];

          for (let index = 0; index < Math.min(viewer.slideCount, 60); index += 1) {
            const thumbContainer = thumbnailRefs.current[index];
            if (thumbContainer) {
              thumbContainer.innerHTML = "";
              const handle = viewer.renderThumbnailToContainer(index, thumbContainer, { width: 116 });
              if (handle) {
                thumbnailHandles.current.push(handle);
              }
            }
          }
        }, 0);
      })
      .catch((error: unknown) => {
        if (abortController.signal.aborted) {
          return;
        }
        if (import.meta.env.DEV) {
          console.error("PPTX render failed", error);
        }
        setStatus("error");
        setErrorMessage("This presentation could not be rendered. It may use unsupported PowerPoint content, encryption, or a malformed package.");
      });

    return () => {
      abortController.abort();
      thumbnailHandles.current.forEach((handle) => handle.dispose());
      thumbnailHandles.current = [];
      viewerRef.current?.destroy();
      viewerRef.current = undefined;
      container.innerHTML = "";
    };
  }, [document]);

  const goToSlide = useCallback(
    async (index: number) => {
      const viewer = viewerRef.current;
      if (!viewer) {
        return;
      }

      const bounded = Math.min(Math.max(index, 0), Math.max(slideCount - 1, 0));
      await viewer.goToSlide(bounded, { behavior: "smooth", block: "center" });
      setCurrentSlide(bounded);
    },
    [slideCount],
  );

  useEffect(() => {
    const viewer = viewerRef.current;
    if (viewer) {
      void viewer.setZoom(zoom);
    }
  }, [zoom]);

  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "ArrowRight") {
        event.preventDefault();
        void goToSlide(currentSlide + 1);
      }

      if (event.key === "ArrowLeft") {
        event.preventDefault();
        void goToSlide(currentSlide - 1);
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [currentSlide, goToSlide, slideCount]);

  useEffect(() => {
    const viewer = viewerRef.current;

    if (!viewer || query.trim().length === 0) {
      viewer?.clearSearchHighlights();
      setMatchCount(0);
      return;
    }

    const matches = viewer.searchText(query.trim());
    setMatchCount(matches.length);
    viewer.clearSearchHighlights();

    if (matches[0]) {
      void viewer.highlightSearchResult(matches[0], {
        borderColor: "#0f766e",
        backgroundColor: "rgba(20, 184, 166, 0.18)",
      });
    }
  }, [query]);

  const requestFullscreen = () => {
    void shellRef.current?.requestFullscreen?.();
  };

  if (status === "error") {
    return <p className="viewer-error">{errorMessage}</p>;
  }

  return (
    <div ref={shellRef} className="office-viewer pptx-viewer">
      <div className="viewer-controls">
        <div className="segmented-controls" aria-label="Presentation navigation">
          <button type="button" className="icon-button" onClick={() => void goToSlide(currentSlide - 1)} disabled={currentSlide <= 0} title="Previous slide">
            <ChevronLeft aria-hidden="true" size={17} />
            <span className="sr-only">Previous slide</span>
          </button>
          <span>
            {slideCount ? currentSlide + 1 : "..."} / {slideCount || "..."}
          </span>
          <button type="button" className="icon-button" onClick={() => void goToSlide(currentSlide + 1)} disabled={currentSlide >= slideCount - 1} title="Next slide">
            <ChevronRight aria-hidden="true" size={17} />
            <span className="sr-only">Next slide</span>
          </button>
        </div>
        <label className="search-field">
          <Search aria-hidden="true" size={17} />
          <span className="sr-only">Search slides</span>
          <input
            type="search"
            value={query}
            onChange={(event) => setQuery(event.currentTarget.value)}
            placeholder="Search slides"
          />
        </label>
        {query && <span className="viewer-note">{matchCount} matches</span>}
        <div className="segmented-controls" aria-label="Presentation zoom controls">
          <button type="button" className="icon-button" onClick={() => setZoom((value) => Math.max(40, value - 10))} title="Zoom out">
            <Minus aria-hidden="true" size={17} />
            <span className="sr-only">Zoom out</span>
          </button>
          <span>{zoom}%</span>
          <button type="button" className="icon-button" onClick={() => setZoom((value) => Math.min(220, value + 10))} title="Zoom in">
            <Plus aria-hidden="true" size={17} />
            <span className="sr-only">Zoom in</span>
          </button>
          <button type="button" className="icon-button" onClick={requestFullscreen} title="Fullscreen">
            <Maximize2 aria-hidden="true" size={17} />
            <span className="sr-only">Fullscreen</span>
          </button>
        </div>
      </div>
      {status === "loading" && <p className="viewer-note">Rendering presentation locally...</p>}
      <div className="presentation-layout">
        <aside className="thumbnail-rail" aria-label="Slide thumbnails">
          {Array.from({ length: slideCount || 1 }, (_, index) => (
            <button
              key={index}
              type="button"
              className={index === currentSlide ? "is-active" : undefined}
              onClick={() => void goToSlide(index)}
              aria-label={`Go to slide ${index + 1}`}
            >
              <span>Slide {index + 1}</span>
              <div ref={(element) => { thumbnailRefs.current[index] = element; }} />
            </button>
          ))}
        </aside>
        <div ref={containerRef} className="pptx-render-surface" />
      </div>
    </div>
  );
}

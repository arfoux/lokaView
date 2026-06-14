# Local Office Viewer

Local Office Viewer is a privacy-first, browser-local document viewer for common office files. It opens documents with the browser File API, renders them in the client, and serves as a clean foundation for future local tools.

Core promise:

> Your documents are processed locally in your browser. Files are not uploaded to a server.

## Supported Formats

Implemented:

- PDF (`.pdf`): PDF.js rendering, page navigation, progressive page rendering, zoom, fit width, rotate, fullscreen, encrypted-file failure message, original download.
- Word (`.docx`): docx-preview rendering, zoom, layout-difference notice, unsafe link/resource neutralization, original download.
- Excel (`.xlsx`): browser-local parsing with `read-excel-file`, sheet tabs, bounded grid preview, row/column headers, basic fill/font style preview, search, CSV export, JSON export, original download.
- PowerPoint (`.pptx`): `@aiden0z/pptx-renderer` integration, recommended ZIP limits, slide navigation, thumbnails, fullscreen, zoom, keyboard arrows, text search, local PDF.js fallback assets for embedded previews, original download.
- CSV (`.csv`): local delimiter detection, grid preview, search, CSV/JSON export, original download.

Intentionally unsupported:

- Legacy binary Office files (`.doc`, `.xls`, `.ppt`). The app detects them and asks users to save as DOCX, XLSX, or PPTX.

## Privacy Model

The core app has no login, database, analytics, remote conversion service, or Worker-side document parsing.

Selected files are read through browser File APIs and passed to local adapters inside the same browser session. The frontend may fetch its own static JavaScript, CSS, and worker assets from the app origin, but selected document bytes, text, names, thumbnails, and metadata are not sent to a remote endpoint.

The optional `/url/...` mode is different from local upload: it fetches a document through the Cloudflare Worker proxy, then renders the returned bytes client-side with the same viewer pipeline as a local file. This exists for internal alias URLs such as `/url/fileku/datamahasiswa.docx` and full HTTP(S) document URLs such as `/url/https://calibre-ebook.com/downloads/demos/demo.docx`; it is a proxy response, not a redirect.

GitHub-backed allowlisted URL sources are fetched by the Worker with a `GITHUB_TOKEN` Cloudflare secret. The token is never placed in frontend code, local storage, query strings, or responses, and it is not used for arbitrary GitHub owner/repo input.

Additional safeguards:

- CSP in `public/_headers` restricts network destinations.
- Document links and remote resource attributes are neutralized after DOCX/PPTX rendering where practical.
- Office ZIP packages are inspected before parsing.
- Large files require confirmation, and unreasonable files are blocked.
- Old active documents are disposed when switching files.

## Local Development

```bash
npm install
npm run dev
```

Useful scripts:

```bash
npm run typecheck
npm run lint
npm run test
npm run build
npm run preview
```

## Cloudflare Deployment

This project uses Vite for the frontend build and Wrangler Static Assets for deployment:

```bash
npm run build
npm run deploy
```

`wrangler.toml` points `assets.directory` at `./dist` and enables `single-page-application` fallback routing.

The Worker also handles `/url/*` before falling back to static assets. Configure the GitHub token as a Worker secret before using GitHub-backed URL aliases:

```bash
npx wrangler secret put GITHUB_TOKEN
```

Note: the current Cloudflare Vite plugin line requires Wrangler 4. The Wrangler-3-compatible plugin line has dev-server audit advisories, so this repo intentionally uses direct Wrangler Static Assets config instead. The config uses features available in Wrangler 3.114.x; the repo pins a patched Wrangler 3.114 release.

## Project Structure

- `src/app`: top-level React app, state lifecycle, app config.
- `src/components`: shared shell UI, toolbar, drop zone, states, error boundary.
- `src/documents`: adapter contracts, detection, limits, registry, local file reading, session cleanup.
- `src/documents/pdf`: PDF adapter and PDF.js viewer.
- `src/documents/docx`: DOCX adapter and docx-preview viewer.
- `src/documents/xlsx`: XLSX adapter and shared spreadsheet grid.
- `src/documents/pptx`: PPTX adapter and renderer integration.
- `src/documents/csv`: CSV parser, adapter, and grid viewer.
- `src/worker`: Cloudflare Worker entry and `/url/*` proxy helpers.
- `src/tests`: detection, limits, adapter smoke, cleanup, malformed/unsupported tests.
- `public/_headers`: deployment security headers.

## Dependency Overview

- React + Vite + TypeScript for the app shell.
- `pdfjs-dist` for PDF rendering.
- `docx-preview` for browser-local DOCX previews.
- `read-excel-file` plus `fflate` for XLSX parsing and lightweight OOXML style inspection, chosen over SheetJS because production audit reported unresolved `xlsx` advisories.
- `@aiden0z/pptx-renderer` for local PPTX rendering.
- `lucide-react` for UI icons.
- Vitest and ESLint for checks.

## Security Notes

Documents are treated as untrusted input. The app validates basic signatures, detects legacy Office formats, inspects ZIP package shape and size, caps spreadsheet/CSV previews, avoids server-side processing, and uses a restrictive CSP.

`npm audit --omit=dev` passes with zero production vulnerabilities. A full audit may still report dev-tooling advisories for Wrangler 3; resolving those requires Wrangler 4, which is outside the Wrangler 3 compatibility target for this build.

## Known Limitations

- DOCX rendering is not pixel-perfect Microsoft Word layout.
- PPTX rendering is best-effort; uncommon effects, embedded media, or unsupported content may degrade.
- XLSX previews prioritize visible values and responsive grids; basic fills/font colors are supported, while formulas, rich formatting, charts, pivot tables, and merged-cell fidelity are limited.
- Spreadsheet/CSV previews are capped for responsiveness.
- Password-protected/encrypted documents are not opened.
- No OCR, annotations, editing, cloud storage, sharing, or conversion pipeline is implemented yet.

## Browser Expectations

Modern evergreen browsers with File, Blob, module worker, Canvas, and ES2022 support are expected. Very old browsers are not targeted.

## Implemented Bonus Tools

- XLSX active sheet export to CSV.
- XLSX active sheet export to JSON.
- CSV export to JSON.
- Local file metadata display.
- Local-only privacy panel.

## Future Extension Ideas

- PWA/offline shell.
- Virtualized spreadsheet rows for larger workbooks.
- Text extraction/search for DOCX.
- More complete PPTX outline/notes export.
- Local PDF split/merge/reorder tools with `pdf-lib`.
- Annotation layer.
- Localization.
- Desktop packaging.

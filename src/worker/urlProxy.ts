const GITHUB_BLOB_HOST = "github.com";
const REMOTE_URL_PATTERN = /^https?:\/\//i;

export interface UrlAlias {
  readonly baseUrl: string;
}

export interface ResolvedUrlDocument {
  readonly sourceKind: "alias" | "remote-url";
  readonly alias: string;
  readonly relativePath: string;
  readonly sourceUrl: string;
  readonly fileName: string;
  readonly contentType: string;
}

export class UrlProxyError extends Error {
  readonly status: number;

  constructor(status: number, message: string) {
    super(message);
    this.name = "UrlProxyError";
    this.status = status;
  }
}

export const URL_DOCUMENT_ALIASES: Record<string, UrlAlias> = {
  fileku: {
    baseUrl: "https://github.com/arfoux/simpenan/blob/main",
  },
};

export function resolveDocumentContentType(fileName: string) {
  const lowerName = fileName.toLowerCase();

  if (lowerName.endsWith(".pdf")) {
    return "application/pdf";
  }

  if (lowerName.endsWith(".docx")) {
    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  }

  if (lowerName.endsWith(".xlsx")) {
    return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  }

  if (lowerName.endsWith(".pptx")) {
    return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  }

  if (lowerName.endsWith(".csv")) {
    return "text/csv; charset=utf-8";
  }

  return "application/octet-stream";
}

export function createInlineContentDisposition(fileName: string) {
  const safeFileName = fileName.replace(/[\\"]/g, "_");

  return `inline; filename="${safeFileName}"`;
}

function parseIpv4Address(hostname: string) {
  const parts = hostname.split(".");

  if (parts.length !== 4) {
    return null;
  }

  const octets = parts.map((part) => {
    if (!/^\d+$/.test(part)) {
      return Number.NaN;
    }

    return Number.parseInt(part, 10);
  });

  return octets.every((octet) => Number.isInteger(octet) && octet >= 0 && octet <= 255) ? octets : null;
}

function isBlockedIpv4Address(hostname: string) {
  const octets = parseIpv4Address(hostname);

  if (!octets) {
    return false;
  }

  const [first = 0, second = 0] = octets;

  return first === 0 ||
    first === 10 ||
    first === 127 ||
    (first === 169 && second === 254) ||
    (first === 172 && second >= 16 && second <= 31) ||
    (first === 192 && second === 168) ||
    (first === 100 && second >= 64 && second <= 127) ||
    (first === 192 && second === 0) ||
    (first === 198 && (second === 18 || second === 19)) ||
    first >= 224;
}

function isBlockedHostname(hostname: string) {
  const normalized = hostname.toLowerCase().replace(/^\[|\]$/g, "");

  return normalized === "localhost" ||
    normalized.endsWith(".localhost") ||
    normalized.endsWith(".local") ||
    normalized === "::1" ||
    (normalized.includes(":") && (normalized.startsWith("fc") || normalized.startsWith("fd") || normalized.startsWith("fe80"))) ||
    isBlockedIpv4Address(normalized);
}

function validateRemoteSourceUrl(url: URL) {
  if (url.protocol !== "https:" && url.protocol !== "http:") {
    throw new UrlProxyError(400, "Only HTTP and HTTPS document URLs are supported.");
  }

  if (url.username || url.password) {
    throw new UrlProxyError(400, "Document URLs cannot contain credentials.");
  }

  if (isBlockedHostname(url.hostname)) {
    throw new UrlProxyError(400, "That document host is not allowed.");
  }

  if (!url.pathname || url.pathname === "/") {
    throw new UrlProxyError(400, "Document URL must include a file path.");
  }
}

function decodePathSegment(segment: string) {
  try {
    return decodeURIComponent(segment);
  } catch {
    throw new UrlProxyError(400, "The document URL contains invalid encoding.");
  }
}

function encodePathSegments(path: string) {
  return path
    .split("/")
    .filter(Boolean)
    .map((segment) => encodeURIComponent(segment))
    .join("/");
}

function joinUrlPath(baseUrl: string, relativePath: string) {
  const base = new URL(baseUrl);
  const basePath = base.pathname.replace(/\/+$/, "");
  const encodedRelativePath = encodePathSegments(relativePath);

  base.pathname = `${basePath}/${encodedRelativePath}`;
  return base.toString();
}

function getFileNameFromPath(pathname: string) {
  const fileName = pathname.split("/").filter(Boolean).at(-1) ?? "";

  try {
    return decodeURIComponent(fileName);
  } catch {
    return fileName;
  }
}

function decodeRoutePath(routePath: string) {
  try {
    return decodeURIComponent(routePath);
  } catch {
    throw new UrlProxyError(400, "The document URL contains invalid encoding.");
  }
}

function resolveRemoteUrlDocument(routePath: string, requestSearch: string): ResolvedUrlDocument {
  const decodedRoutePath = decodeRoutePath(routePath);
  const sourceUrl = new URL(decodedRoutePath.includes("?") ? decodedRoutePath : `${decodedRoutePath}${requestSearch}`);

  validateRemoteSourceUrl(sourceUrl);

  const fileName = getFileNameFromPath(sourceUrl.pathname);

  if (!fileName) {
    throw new UrlProxyError(400, "Document URL must include a file name.");
  }

  return {
    sourceKind: "remote-url",
    alias: "",
    relativePath: sourceUrl.pathname,
    sourceUrl: sourceUrl.toString(),
    fileName,
    contentType: resolveDocumentContentType(fileName),
  };
}

export function resolveUrlDocument(
  pathname: string,
  aliases: Record<string, UrlAlias> = URL_DOCUMENT_ALIASES,
  requestSearch = "",
): ResolvedUrlDocument {
  const routePrefix = "/url/";

  if (!pathname.startsWith(routePrefix)) {
    throw new UrlProxyError(404, "Document URL route was not found.");
  }

  const routePath = pathname.slice(routePrefix.length);

  if (REMOTE_URL_PATTERN.test(routePath) || REMOTE_URL_PATTERN.test(decodeRoutePath(routePath))) {
    return resolveRemoteUrlDocument(routePath, requestSearch);
  }

  const rawSegments = routePath.split("/").filter(Boolean);
  const alias = rawSegments[0] ? decodePathSegment(rawSegments[0]) : "";

  if (!alias) {
    throw new UrlProxyError(400, "Choose a registered document URL alias.");
  }

  const config = aliases[alias];

  if (!config) {
    throw new UrlProxyError(404, "Unknown document URL alias.");
  }

  const decodedPathSegments = rawSegments.slice(1).map(decodePathSegment);
  const relativePath = decodedPathSegments.join("/");

  if (!relativePath) {
    throw new UrlProxyError(400, "Document path is required.");
  }

  if (relativePath.includes("..")) {
    throw new UrlProxyError(400, "Document path cannot contain parent directory segments.");
  }

  const fileName = decodedPathSegments.at(-1) ?? "";

  if (!fileName) {
    throw new UrlProxyError(400, "Document file name is required.");
  }

  return {
    sourceKind: "alias",
    alias,
    relativePath,
    sourceUrl: joinUrlPath(config.baseUrl, relativePath),
    fileName,
    contentType: resolveDocumentContentType(fileName),
  };
}

export function githubBlobUrlToContentsApiUrl(blobUrl: string) {
  const url = new URL(blobUrl);

  if (url.hostname !== GITHUB_BLOB_HOST) {
    return null;
  }

  const segments = url.pathname.split("/").filter(Boolean).map(decodePathSegment);
  const [owner, repo, marker, branch, ...pathSegments] = segments;

  if (!owner || !repo || marker !== "blob" || !branch || pathSegments.length === 0) {
    return null;
  }

  const encodedPath = pathSegments.map((segment) => encodeURIComponent(segment)).join("/");
  const apiUrl = new URL(`https://api.github.com/repos/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}/contents/${encodedPath}`);
  apiUrl.searchParams.set("ref", branch);

  return apiUrl.toString();
}

function getGitHubRepoRef(blobUrl: string) {
  const url = new URL(blobUrl);

  if (url.hostname !== GITHUB_BLOB_HOST) {
    return null;
  }

  const [owner, repo, marker, branch] = url.pathname.split("/").filter(Boolean).map(decodePathSegment);

  if (!owner || !repo || marker !== "blob" || !branch) {
    return null;
  }

  return `${owner}/${repo}@${branch}`.toLowerCase();
}

export function isGitHubBlobSourceAllowed(
  blobUrl: string,
  aliases: Record<string, UrlAlias> = URL_DOCUMENT_ALIASES,
) {
  const requestedRepoRef = getGitHubRepoRef(blobUrl);

  if (!requestedRepoRef) {
    return false;
  }

  return Object.values(aliases).some((alias) => getGitHubRepoRef(alias.baseUrl) === requestedRepoRef);
}

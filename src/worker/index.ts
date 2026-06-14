import {
  UrlProxyError,
  createInlineContentDisposition,
  githubBlobUrlToContentsApiUrl,
  isGitHubBlobSourceAllowed,
  resolveUrlDocument,
} from "./urlProxy";

interface AssetsBinding {
  fetch(request: Request): Promise<Response>;
}

interface WorkerEnv {
  readonly ASSETS: AssetsBinding;
  readonly GITHUB_TOKEN?: string;
}

function textResponse(message: string, status: number) {
  return new Response(message, {
    status,
    headers: {
      "Content-Type": "text/plain; charset=utf-8",
      "X-Content-Type-Options": "nosniff",
    },
  });
}

async function handleUrlDocumentRequest(request: Request, env: WorkerEnv) {
  if (request.method !== "GET" && request.method !== "HEAD") {
    return textResponse("Method not allowed.", 405);
  }

  let document;

  try {
    const url = new URL(request.url);
    document = resolveUrlDocument(url.pathname, undefined, url.search);
  } catch (error) {
    if (error instanceof UrlProxyError) {
      return textResponse(error.message, error.status);
    }

    throw error;
  }

  const githubApiUrl = githubBlobUrlToContentsApiUrl(document.sourceUrl);
  let upstreamResponse: Response;

  if (githubApiUrl) {
    if (!isGitHubBlobSourceAllowed(document.sourceUrl)) {
      return textResponse("That GitHub document source is not allowed.", 403);
    }

    if (!env.GITHUB_TOKEN) {
      return textResponse("The document URL proxy is not configured.", 500);
    }

    upstreamResponse = await fetch(githubApiUrl, {
      headers: {
        Authorization: `Bearer ${env.GITHUB_TOKEN}`,
        Accept: "application/vnd.github.raw",
        "User-Agent": "lokaView-url-proxy",
      },
      redirect: "manual",
    });
  } else {
    upstreamResponse = await fetch(document.sourceUrl, {
      headers: {
        "User-Agent": "lokaView-url-proxy",
      },
      redirect: "manual",
    });
  }

  if (!upstreamResponse.ok) {
    const status = upstreamResponse.status === 404 ? 404 : 502;
    return textResponse("The requested document could not be loaded.", status);
  }

  return new Response(request.method === "HEAD" ? null : upstreamResponse.body, {
    status: 200,
    headers: {
      "Content-Type": document.contentType,
      "Content-Disposition": createInlineContentDisposition(document.fileName),
      "Cache-Control": "private, max-age=60",
      "X-Content-Type-Options": "nosniff",
    },
  });
}

export default {
  async fetch(request: Request, env: WorkerEnv) {
    const url = new URL(request.url);

    if (url.pathname === "/url" || url.pathname.startsWith("/url/")) {
      return handleUrlDocumentRequest(request, env);
    }

    return env.ASSETS.fetch(request);
  },
};

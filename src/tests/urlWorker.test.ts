import { describe, expect, it } from "vitest";
import { shouldServeAppShellForUrlRequest } from "../worker";

describe("url worker routing", () => {
  it("serves the app shell for direct browser navigation to /url paths", () => {
    const request = new Request("https://example.com/url/fileku/document.docx", {
      headers: {
        Accept: "text/html,application/xhtml+xml",
        "Sec-Fetch-Dest": "document",
      },
    });

    expect(shouldServeAppShellForUrlRequest(request)).toBe(true);
  });

  it("keeps programmatic URL fetches on the file proxy path", () => {
    const request = new Request("https://example.com/url/fileku/document.docx", {
      headers: {
        Accept: "*/*",
        "Sec-Fetch-Dest": "empty",
      },
    });

    expect(shouldServeAppShellForUrlRequest(request)).toBe(false);
  });
});

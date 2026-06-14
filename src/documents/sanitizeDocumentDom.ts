const REMOTE_URL_PATTERN = /^(https?:|\/\/)/i;

function isRemoteUrl(value: string | null): boolean {
  return Boolean(value && REMOTE_URL_PATTERN.test(value.trim()));
}

export function neutralizeUnsafeDocumentDom(root: HTMLElement): void {
  root.querySelectorAll<HTMLElement>("[onclick], [onload], [onerror]").forEach((element) => {
    for (const attribute of Array.from(element.attributes)) {
      if (attribute.name.toLowerCase().startsWith("on")) {
        element.removeAttribute(attribute.name);
      }
    }
  });

  root.querySelectorAll<HTMLAnchorElement>("a[href]").forEach((anchor) => {
    const href = anchor.getAttribute("href");

    if (isRemoteUrl(href)) {
      anchor.dataset.blockedHref = href ?? "";
      anchor.removeAttribute("href");
      anchor.setAttribute("aria-disabled", "true");
      anchor.title = "External links from documents are disabled by default.";
    } else {
      anchor.rel = "noopener noreferrer";
      anchor.target = "_blank";
    }
  });

  root.querySelectorAll<HTMLImageElement | HTMLIFrameElement | HTMLSourceElement>("[src]").forEach(
    (element) => {
      const src = element.getAttribute("src");

      if (isRemoteUrl(src)) {
        element.dataset.blockedSrc = src ?? "";
        element.removeAttribute("src");
        element.setAttribute("aria-label", "External document resource blocked");
      }
    },
  );

  root.querySelectorAll<HTMLElement>("[style]").forEach((element) => {
    const style = element.getAttribute("style");

    if (style && /url\(\s*['"]?(https?:|\/\/)/i.test(style)) {
      element.removeAttribute("style");
      element.dataset.blockedInlineStyle = "remote-url";
    }
  });
}

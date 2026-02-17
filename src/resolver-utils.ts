export function isGraphNotFound(error: unknown): boolean {
  const message = error instanceof Error ? error.message : String(error);
  return message.includes("Graph request failed (404)") || message.includes("itemNotFound");
}

export function isGraphAccessDenied(error: unknown): boolean {
  const message = error instanceof Error ? error.message : String(error);
  return message.includes("Graph request failed (403)") || message.includes("accessDenied");
}

export function isGraphCurrentVersionContentUnsupported(error: unknown): boolean {
  const message = error instanceof Error ? error.message : String(error);
  return message.includes("Graph content request failed (400)")
    && message.includes("invalidRequest")
    && message.includes("current version");
}

export function buildRemotePathCandidates(remotePath: string): string[] {
  const normalized = remotePath.startsWith("/") ? remotePath : `/${remotePath}`;
  const segments = normalized.split("/").filter((s) => s.length > 0);
  if (!segments.length) {
    return ["/"];
  }

  const candidates: string[] = [];
  for (let start = 0; start < segments.length; start++) {
    candidates.push(`/${segments.slice(start).join("/")}`);
  }

  return [...new Set(candidates)];
}

export function normalizeShareBaseUrl(input: string): string {
  const trimmed = input.trim();
  if (!trimmed) {
    return trimmed;
  }
  return trimmed.replace(/^https:\/(?!\/)/i, "https://").replace(/^http:\/(?!\/)/i, "http://");
}

export function appendPathSegmentsToUrl(baseUrl: string, segments: string[]): string {
  const url = new URL(baseUrl);
  const basePath = url.pathname.replace(/\/+$/, "");
  const extraPath = segments.map((segment) => encodeURIComponent(segment)).join("/");
  url.pathname = extraPath ? `${basePath}/${extraPath}` : basePath || "/";
  return url.toString();
}

export function toGraphShareId(webUrl: string): string {
  const base64 = Buffer.from(webUrl, "utf8").toString("base64");
  const base64Url = base64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
  return `u!${base64Url}`;
}

export function getRelativePathByUrlPrefix(targetUrl: string, baseUrl: string): string | undefined {
  try {
    const target = new URL(targetUrl);
    const base = new URL(baseUrl);

    if (target.origin.toLowerCase() !== base.origin.toLowerCase()) {
      return undefined;
    }

    const targetPath = target.pathname.replace(/\/+$/, "");
    const basePath = base.pathname.replace(/\/+$/, "");
    if (!targetPath.toLowerCase().startsWith(basePath.toLowerCase())) {
      return undefined;
    }

    const remaining = targetPath.slice(basePath.length).replace(/^\/+/, "");
    return decodeURIComponent(remaining);
  } catch {
    return undefined;
  }
}

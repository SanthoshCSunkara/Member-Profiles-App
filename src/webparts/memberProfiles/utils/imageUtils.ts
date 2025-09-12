export function buildRendition(rawUrl: string, w: number, h = w): string {
  try {
    const u = new URL(rawUrl, location.origin);
    u.searchParams.set('width', String(w));
    u.searchParams.set('height', String(h));
    u.searchParams.set('mode', 'crop');
    return u.toString();
  } catch {
    const sep = rawUrl.indexOf('?') ? '&' : '?';
    return `${rawUrl}${sep}width=${w}&height=${h}&mode=crop`;
  }
}

export function buildPreview(rawUrl: string, w: number, h = w): string {
  return `/_layouts/15/getpreview.ashx?path=${encodeURIComponent(rawUrl)}&width=${w}&height=${h}&mode=crop`;
}

/** DPR-aware sources for crisp avatars in grids */
export function avatarSources(baseUrl: string, cssPx: number, dprCap = 2) {
  const dpr = Math.min(Math.ceil((window as any).devicePixelRatio || 1), dprCap);
  const px1 = cssPx, px2 = cssPx * dpr;
  return {
    src: buildRendition(baseUrl, px1, px1),
    srcSet: `${buildRendition(baseUrl, px1, px1)} 1x, ${buildRendition(baseUrl, px2, px2)} ${dpr}x`,
    px2
  };
}

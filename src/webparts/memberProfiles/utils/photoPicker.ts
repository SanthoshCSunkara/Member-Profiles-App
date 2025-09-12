//import type { MSGraphClientFactory } from '@microsoft/sp-http';
export type PhotoSource = 'm365' | 'library' | 'none';


function getBase(siteUrl?: string): string {
  // window.location.origin may be undefined in some tooling contexts
  const loc = (typeof window !== 'undefined' && (window as any).location && (window as any).location.origin)
    ? (window as any).location.origin
    : '';
  const base = (siteUrl && siteUrl.length ? siteUrl : loc) || '';
  return base.replace(/\/$/, '');
}

/** Normalize to absolute URL using the current site (for workbench/local too). */
export function toAbsolute(url?: string, siteUrl?: string): string | undefined {
  if (!url) return undefined;
  if (/^https?:\/\//i.test(url)) return url;
  const base = getBase(siteUrl);
  if (!base) return url;
  // Avoid startsWith; use charAt
  return (url.length > 0 && url.charAt(0) === '/')
    ? (base + url)
    : (base + '/' + url);
}

/** SPO userphoto handler – always return something (photo or silhouette). */
export const userPhotoUrl = (upn: string, size: 'S'|'M'|'L' = 'L', siteBase?: string) => {
  const base = (siteBase || '').replace(/\/$/, '');
  return base + '/_layouts/15/userphoto.aspx?size=' + size + '&accountname=' + encodeURIComponent(upn);
};


/** Extract a usable UPN/email from any likely field on the item. */
export function extractUpn(item: unknown): string | undefined {
  const it = item as any;
  let upn: string | undefined =
    it?.email ?? it?.mail ?? it?.userPrincipalName ?? it?.upn ??
    it?.workEmail ?? it?.UserPrincipalName ?? it?.AccountName ?? it?.LoginName;

  if (typeof upn === 'string' && upn.indexOf('|') !== -1) {
    const parts = upn.split('|');
    upn = parts[parts.length - 1];
  }
  return (typeof upn === 'string' && upn.length) ? upn : undefined;
}

/**
 * Prefer M365 photo URL unconditionally if we have a UPN (no Graph required).
 * If no UPN, or you explicitly want to try the library, use the library URL.
 * Graph is optional: if provided/approved you could check existence, but we
 * intentionally *do not block* on it for workbench reliability.
 */
export async function pickBestPhotoUrl(
  upn: string | undefined,
  libraryPhotoUrl: string | undefined,
  _graphFactory?: any,
  siteUrl?: string
): Promise<{ url?: string; source: PhotoSource }> {
  if (upn) return { url: userPhotoUrl(upn, 'L', siteUrl), source: 'm365' };
  const absLib = toAbsolute(libraryPhotoUrl, siteUrl);
  if (absLib) return { url: absLib, source: 'library' };
  return { url: undefined, source: 'none' };
}

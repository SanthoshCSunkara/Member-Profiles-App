import { spfi, SPFx } from '@pnp/sp';
import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { IProfileItem, IProfileItemRaw } from '../models';

export class SpService {
  private _sp: SPFI;
  public constructor(private readonly _context: WebPartContext) {
    this._sp = spfi().using(SPFx(this._context));
  }

  /** Data lists (Custom List) */
  public async getLists(): Promise<Array<{ key: string; text: string }>> {
    const all = await this._sp.web.lists.select('Id','Title','BaseTemplate','Hidden')();
    const out: Array<{ key: string; text: string }> = [];
    for (let i = 0; i < all.length; i++) {
      const l = all[i];
      if (l.BaseTemplate === 100 && !l.Hidden) out.push({ key: l.Id, text: l.Title });
    }
    return out;
  }

  /** Image libraries (Document=101, Picture=109) */
  public async getImageLibraries(): Promise<Array<{ key: string; text: string }>> {
    const all = await this._sp.web.lists.select('Id','Title','BaseTemplate','Hidden')();
    const out: Array<{ key: string; text: string }> = [];
    for (let i = 0; i < all.length; i++) {
      const l = all[i];
      // 101: Document Library, 109: Picture Library
      if ((l.BaseTemplate === 101 || l.BaseTemplate === 109) && !l.Hidden) {
        out.push({ key: l.Id, text: l.Title });
      }
    }
    return out;
  }

  /** Profiles from the selected data list (no images here on purpose) */
  public async getProfiles(listId: string): Promise<IProfileItem[]> {
    if (!listId) return [];
    const rows: IProfileItemRaw[] = await this._sp.web.lists
      .getById(listId)
      .items
      .select(
        'Id','Title','Role','Hire_x0020_Date','Birthday',
        'CompanyProfile','LinkedIn','About','Modified','Created'
      )
      .top(5000)();

    const formatDate = (d?: string | Date): string | undefined => {
      if (!d) return undefined;
      const date = d instanceof Date ? d : new Date(d as string);
      if (isNaN(date.getTime())) return undefined;
      return date.toLocaleDateString();
    };

    const mapUrl = (v: any): string | undefined => {
      if (!v) return undefined;
      if (typeof v === 'string') return v;
      if (v.Url) return v.Url;
      if (v.url) return v.url;
      return undefined;
    };

    const items: IProfileItem[] = [];
    for (let i = 0; i < rows.length; i++) {
      const r: any = rows[i];
      items.push({
        id: r.Id,
        name: r.Title,
        role: r.Role || '',
        hireDate: formatDate(r.Hire_x0020_Date),
        birthday: r.Birthday || undefined,
        companyUrl: mapUrl(r.CompanyProfile),
        linkedInUrl: mapUrl(r.LinkedIn),
        photoUrl: undefined,          // will be filled by image map
        detailsHtml: r.About || undefined
      });
    }
    return items;
  }

  /** Build a name->absoluteUrl map from the selected image library */
  public async getImageMap(imageListId?: string): Promise<{ [key: string]: string }> {
    const map: { [key: string]: string } = {};
    if (!imageListId) return map;

    const origin = new URL(this._context.pageContext.web.absoluteUrl).origin;

    // FileRef = server relative path; Title may be filled in library
    const rows: Array<{ Title?: string; FileRef: string }> = await this._sp.web.lists
      .getById(imageListId)
      .items
      .select('Id','Title','FileRef')
      .top(5000)();

    const norm = (s?: string): string =>
      (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');

    const fileBase = (ref: string): string => {
      // get filename without extension
      const slash = ref.lastIndexOf('/');
      const name = slash >= 0 ? ref.substring(slash + 1) : ref;
      const dot = name.lastIndexOf('.');
      const base = dot > 0 ? name.substring(0, dot) : name;
      return norm(base);
    };

    for (let i = 0; i < rows.length; i++) {
      const it = rows[i];
      const absolute = origin + it.FileRef;
      const k1 = norm(it.Title);
      const k2 = fileBase(it.FileRef);
      if (k1) map[k1] = absolute;
      if (k2) map[k2] = absolute;
    }
    return map;
  }
}

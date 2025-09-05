// src/webparts/memberProfiles/services/SpService.ts
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

  public async getLists(): Promise<Array<{ key: string; text: string }>> {
    const lists = await this._sp.web.lists
      .select('Id','Title','BaseTemplate','Hidden')()
      .then(ls => ls.filter(l => l.BaseTemplate === 100 && !l.Hidden));
    return lists.map(l => ({ key: l.Id, text: l.Title }));
  }

  public async getProfiles(listId: string): Promise<IProfileItem[]> {
    if (!listId) return [];

    // IMPORTANT: only select Image0 (your list doesn’t have "Image")
    const select = [
      'Id','Title','Role','Hire_x0020_Date','Birthday',
      'CompanyProfile','LinkedIn','Image0','About',
      'Modified','Created'
    ].join(',');

    const rows: IProfileItemRaw[] = await this._sp.web.lists
      .getById(listId)
      .items.select(select)
      .top(5000)();

    const origin = new URL(this._context.pageContext.web.absoluteUrl).origin;

    const mapUrl = (v: any): string | undefined => {
      if (!v) return undefined;
      if (typeof v === 'string') return v;
      if (v.Url) return v.Url;
      if (v.url) return v.url;
      if (v.serverUrl && v.serverRelativeUrl) return v.serverUrl + v.serverRelativeUrl;
      if (v.ServerUrl && v.ServerRelativeUrl) return v.ServerUrl + v.ServerRelativeUrl;
      return undefined;
    };

    // Robust resolver for Image (Thumbnail) column; ES5-safe
    const mapImage = (img: any): string | undefined => {
      if (!img) return undefined;
      const v = Array.isArray(img) ? img[0] : img;

      const shapeToUrl = (o: any): string | undefined => {
        if (!o) return undefined;
        const su = o.serverUrl || o.ServerUrl;
        const sr = o.serverRelativeUrl || o.ServerRelativeUrl;
        const u  = o.Url || o.url;
        if (su && sr) return su + sr;
        if (u) return u;
        if (sr) return origin + sr;
        if (typeof o.path === 'string') {
          return o.path && o.path.charAt(0) === '/' ? origin + o.path : o.path;
        }
        return undefined;
      };

      if (typeof v === 'string') {
        try {
          const j = JSON.parse(v);
          const u = shapeToUrl(j);
          if (u) return u;
        } catch {
          if (/^https?:/i.test(v)) return v;
          if (v && v.charAt(0) === '/') return origin + v;
        }
        return undefined;
      }

      return shapeToUrl(v);
    };

    const formatDate = (d?: string | Date): string | undefined => {
      if (!d) return undefined;
      const date = d instanceof Date ? d : new Date(d as string);
      if (isNaN(date.getTime())) return undefined;
      return date.toLocaleDateString();
    };

    return rows.map((i: any) => {
      const imgSource = (i as any).Image0; // <- use Image0 only

      return {
        id: i.Id,
        name: i.Title,
        role: i.Role || '',
        hireDate: formatDate(i.Hire_x0020_Date),
        birthday: i.Birthday || undefined,
        companyUrl: mapUrl(i.CompanyProfile),
        linkedInUrl: mapUrl(i.LinkedIn),
        photoUrl: mapImage(imgSource),
        detailsHtml: i.About || undefined
      } as IProfileItem;
    });
  }
}

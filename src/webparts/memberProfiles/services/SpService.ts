import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFI, SPFx } from '@pnp/sp';

// Add webs/lists/items/fields to SPFI via side effects
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';

import type { IProfileItem } from '../models';

// ES5-safe key normalizer (used for matching library images)
const normKey = (s?: string) =>
  (s || '')
    .toString()
    .trim()
    .toLowerCase()
    .replace(/\.[^.]+$/, '')      // drop extension
    .replace(/[^a-z0-9]/g, '');   // compact

export class SpService {
  private _sp: SPFI;

  constructor(private _context: WebPartContext) {
    this._sp = spfi().using(SPFx(this._context));
  }

  /** Return the internal name of the first Image-type field if present (ES5-safe). */
  private async _getImageFieldInternalName(listId: string): Promise<string | undefined> {
    const fields: any[] = await this._sp.web.lists.getById(listId)
      .fields.select('InternalName', 'TypeAsString')();

    for (let i = 0; i < fields.length; i++) {
      const f = fields[i];
      if (f && f.TypeAsString === 'Image') return f.InternalName as string;
    }
    return undefined;
  }

  /** Normalize SP hyperlink / picture payloads to absolute URLs */
  private _mapUrl(v: any): string | undefined {
    // Classic hyperlink/picture
    if (v && typeof v === 'object' && (v.Url || v.url)) return v.Url || v.url;
    // Modern Image column (serverUrl + serverRelativeUrl)
    if (v && typeof v === 'object' && v.serverUrl && v.serverRelativeUrl) {
      return String(v.serverUrl).replace(/\/$/, '') + String(v.serverRelativeUrl);
    }
    if (typeof v === 'string' && v) return v;
    return undefined;
  }

  /** Main list read. Adds transient `upn` for userphoto.aspx */
  public async getProfiles(listId: string): Promise<IProfileItem[]> {
    const imageField = await this._getImageFieldInternalName(listId);

    const select: string[] = [
      'Id', 'Title', 'Role', 'Hire_x0020_Date', 'Birthday',
      'CompanyProfile', 'LinkedIn', 'About',
      'UserName/EMail', 'UserName/Title'
    ];
    if (imageField) select.push(imageField);

    const rows: any[] = await this._sp.web.lists
      .getById(listId)
      .items
      .select(...select)     // keep context; do NOT use .apply(null, …)
      .expand('UserName')();

    const items: IProfileItem[] = [];
    for (let i = 0; i < rows.length; i++) {
      const r: any = rows[i];
      const it: any = {
        id: r.Id,
        name: r.Title,
        role: r.Role,
        hireDate: r.Hire_x0020_Date,
        birthday: r.Birthday,
        companyUrl: this._mapUrl(r.CompanyProfile),
        linkedInUrl: this._mapUrl(r.LinkedIn),
        photoUrl: imageField ? (this._mapUrl(r[imageField]) || this._mapUrl(r.CompanyProfile))
                             : this._mapUrl(r.CompanyProfile),
        detailsHtml: r.About
      };

      const email: string | undefined = r && r.UserName ? r.UserName.EMail : undefined;
      if (email) it.upn = email; // transient only, used for M365 photo
      items.push(it as IProfileItem);
    }
    return items;
  }

  /**
   * Build a dictionary from the image library:
   *   key = normalized Title or FileLeafRef (no extension)
   *   value = absolute URL to image
   */
  public async getLibraryPhotoMap(libraryId: string): Promise<{ [key: string]: string }> {
    const rows: any[] = await this._sp.web.lists.getById(libraryId)
      .items.select('FileRef', 'FileLeafRef', 'Title')
      .top(5000)();

    // tenant root = absoluteUrl minus serverRelativeUrl
    const abs = this._context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    const rel = (this._context.pageContext.web.serverRelativeUrl || '').replace(/\/$/, '');
    const tenantRoot = abs.substring(0, abs.length - rel.length); // e.g. https://tenant.sharepoint.com

    const map: { [key: string]: string } = {};

    for (let i = 0; i < rows.length; i++) {
      const r: any = rows[i];
      const fileRef = String(r.FileRef || ''); // server-relative (/sites/.../file.jpg)
      const isAbs = fileRef && fileRef.toLowerCase().indexOf('http') === 0;
      const url = isAbs ? fileRef : (tenantRoot + fileRef);

      const k1 = normKey(r.Title);
      const k2 = normKey(r.FileLeafRef);

      if (k1 && !map[k1]) map[k1] = url;
      if (k2 && !map[k2]) map[k2] = url;
    }
    return map;
  }
}

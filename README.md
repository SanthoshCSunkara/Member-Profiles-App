# member-profiles

## Summary

Purpose: A modern SharePoint Online web part to browse team member profiles, search by name/role, and view details in a centered modal. Images are mapped from a separate image library without exposing organizational identifiers.

Feature Summary
•	Header & subtitle with optional accent color.
•	Dual search (name + role/title) with rounded inputs.
•	Responsive card grid with avatar, name, role, and “View more details…” hint.
•	Centered modal (details): photo, name, role, hire date, birthday, About (rich text), and action buttons (Company Profile, LinkedIn).
•	Image strategy: Photos are stored in a separate library and auto-matched to people by Title/filename. No per-item linking required in the profiles list.
•	Robust image pipeline: crisp renditions + safe fallback to SharePoint preview handler; final fallback is a neutral color block, not initials.
•	Performance & scale: supports thousands of items, client-side search, optional item cap (0 = all).
•	Security: DOMPurify sanitization for rich text.
•	Accessibility: keyboard/focus states and clear contrast.
________________________________________
2) High-level Architecture
flowchart LR
  A[Profiles List (Custom List)] -- Title, Role, Dates, About --> B(Web Part)
  C[Image Library (Doc/Picture)] -- Files + Title/Name --> B
  subgraph B[Member Profiles Web Part]
    D[SpService]
    E[MemberProfiles Component]
    F[MemberCard]
    G[DetailsPanel]
    H[SCSS Module]
  end
  D -->|getProfiles| A
  D -->|getImageMap| C
  D -->|merge by normalized name| E
  E -->|render grid + search| F
  F -->|open item| G
  H -->|theme + layout + responsiveness| E
Key idea: The web part merges two sources at runtime: 1) Profiles data (Title, Role, Hire Date, Birthday, About, etc.) 2) Image map from an Image Library (maps person-name → photo URL)
No hard-coded field names; the app matches on normalized person name to image Title or file name.
________________________________________
3) Data Model (normalized)
•	Profiles List (Custom List)
o	Title (single line) → person name
o	Role (single line)
o	Hire_x0020_Date (date only)
o	Birthday (single line or date)
o	About (rich text)
o	CompanyProfile (Hyperlink)
o	LinkedIn (Hyperlink)
•	Image Library (Document/Picture Library)
o	Files (jpg/png)
o	Optional Title (set to the person’s name). If empty, the file name is used for matching.
Matching logic (pseudo):
key = lowerCase(removeNonAlnum(person.Title))
imageMap[key] = absoluteUrl
________________________________________
4) User Flow
1)	Property pane: select Profiles List and (optionally) the Image Library.
2)	On load, the web part fetches all profiles + the image map, then merges by normalized name.
3)	The grid displays cards; search filters by name/role.
4)	Clicking a card opens a centered modal with image, details, and links.
________________________________________
5) Key Components
•	SpService
o	getProfiles(listId) → returns typed profile items
o	getImageLibraries() / getLists() for property pane
o	getImageMap(imageListId) → builds name→url map from files
•	MemberProfiles (container)
o	orchestrates data load, search, paging cap, and selection state
•	MemberCard (presentational)
o	renders avatar + name + role
o	thumbnail logic: primary rendition (?width&height&mode=crop) → fallback (/_layouts/15/getpreview.ashx?path=...)
•	DetailsPanel (modal)
o	big image (same crisp + fallback logic), About (sanitized), and action buttons
•	SCSS module
o	responsive sizing via clamp(), neutral modal styles, orange hover/active highlight for cards
________________________________________
6) High-level Code Snippets
Note: These are illustrative excerpts (redacted, no identifiers). Full code lives in the project files.
6.1 Image Map (SpService)
// Build a name->absoluteUrl map from an image library
const rows = await sp.web.lists.getById(imageListId)
  .items.select('Id','Title','FileRef').top(5000)();

const norm = (s?: string) => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
const fileBase = (ref: string) => {
  const idx = ref.lastIndexOf('/');
  const name = idx >= 0 ? ref.substring(idx + 1) : ref;
  const dot = name.lastIndexOf('.');
  return norm(dot > 0 ? name.substring(0, dot) : name);
};

const map: Record<string, string> = {};
for (const it of rows) {
  const absolute = origin + it.FileRef;
  const k1 = norm(it.Title);
  const k2 = fileBase(it.FileRef);
  if (k1) map[k1] = absolute;
  if (k2) map[k2] = absolute;
}
6.2 Merge Profiles + Photos
const profiles = await service.getProfiles(listId);
const imageMap = await service.getImageMap(imageListId);

const norm = (s?: string) => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
const items = profiles.map(p => ({
  ...p,
  photoUrl: imageMap[norm(p.name)] || p.photoUrl
}));
6.3 Card Thumbnail (Primary + Fallback)
const buildPrimary = (raw: string, w: number, h: number) => {
  try { const u = new URL(raw, location.origin);
        u.searchParams.set('width', String(w));
        u.searchParams.set('height', String(h));
        u.searchParams.set('mode', 'crop');
        return u.toString();
  } catch {
    const sep = raw.indexOf('?') > -1 ? '&' : '?';
    return `${raw}${sep}width=${w}&height=${h}&mode=crop`;
  }
};
const buildFallback = (raw: string, w: number, h: number) =>
  `/_layouts/15/getpreview.ashx?path=${encodeURIComponent(raw)}&width=${w}&height=${h}`;
6.4 Modal Image (Crisp + Fallback)
const TARGET = 720;
const base = item?.photoUrl || '';
const [src, setSrc] = useState(base ? buildPrimary(base, TARGET, TARGET) : undefined);
const [usedFallback, setUsedFallback] = useState(false);

useEffect(() => { setSrc(base ? buildPrimary(base, TARGET, TARGET) : undefined); setUsedFallback(false); }, [base]);

const onError = () => {
  if (!base) return;
  if (!usedFallback) { setSrc(buildFallback(base, TARGET, TARGET)); setUsedFallback(true); }
  else { setSrc(base); }
};
6.5 Sanitization (About field)
const safeHtml = useMemo(
  () => ({ __html: DOMPurify.sanitize(item?.detailsHtml || '') }),
  [item]
);
<div className={styles.detailsHtml} dangerouslySetInnerHTML={safeHtml} />
________________________________________
7) Styling & UX Decisions
•	Typography: system font stack; headings/labels use clamp() for responsive sizes; card names at 700 weight for crispness.
•	Card highlight: soft orange hover (#fff7ed) and active outline (#f07f17) for clear selection.
•	Neutral modal: no accent color; respects rich-text defaults; improved spacing (line-height: 1.65).
•	Avatar: circular crop (object-fit: cover), crisp image via rendition; solid color fallback when missing.
•	Search inputs: rounded, accessible, consistent size.
________________________________________
8) Security & Compliance
•	DOMPurify on all rich HTML (About) to mitigate XSS.
•	No organization names or identifiers embedded in code, content, or assets in this document.
•	Uses standard SharePoint APIs and PnPjs; no external network calls beyond the tenant.
________________________________________
9) Performance Notes
•	Single fetch for profiles + single fetch for image items (both up to 5000 via paging on SPO side).
•	Client-side filtering (name/role) with O(1) normalized lookups for image mapping.
•	Optional render cap: itemsPerPage = 0 → render all; otherwise slice client-side.
________________________________________
10) Configuration & Setup (End-user)
1.	Create/choose a Profiles List. Add fields: Role (text), Hire Date (date), Birthday (date or text), About (rich text), CompanyProfile (link), LinkedIn (link).
2.	Create an Image Library (Document/Picture). Upload photos; optionally fill Title per photo to match person names exactly. If Title is empty, filename (without extension) is used.
3.	Add the Web Part to a modern page.
4.	Open Properties: select Profiles List + Image Library. Set Max items (0 = all) and (optional) accent color.
5.	Publish.
________________________________________
11) Troubleshooting
•	No photo for a person: ensure a file in the image library matches the person’s name via Title or filename (normalized: letters/numbers only). Verify audience has read permissions to the library.
•	Blurry images: large originals are preferred; rendition clamps size but avoids upscaling. The modal caps height to avoid soft stretching.
•	Rich-text not rendering as expected: confirm the field is Enhanced rich text or Rich text in the list.
________________________________________
12) Future Enhancements
•	SessionStorage caching of the image map to skip one network call after first load.
•	Optional compact layout (denser cards) toggle.
•	Export to CSV/Excel for quick roster downloads.
•	Paging or virtualized grid for extremely large datasets.
________________________________________
13) Sample Screens (Redacted)
Replace these placeholders with your own neutral screenshots.
Cards grid (desktop):
Cards Grid Placeholder
Cards Grid Placeholder
Centered modal (details):
Details Modal Placeholder
Details Modal Placeholder
Image library (Tiles):
Image Library Placeholder
Image Library Placeholder
________________________________________
14) Appendix – Minimal Interfaces
export interface IMemberProfilesProps {
  listId: string;
  imageListId?: string;
  itemsPerPage?: number;      // 0 = all
  accentColor: string;        // header/cards only; modal is neutral
}

export interface IProfileItem {
  id: number;
  name: string;
  role: string;
  hireDate?: string;
  birthday?: string;
  companyUrl?: string;
  linkedInUrl?: string;
  photoUrl?: string;          // merged from library
  detailsHtml?: string;       // sanitized in modal
}

------------------------------------------------------------------------------------------------------------------------------------------------------------
## Prerequisites

> Any special pre-requisites?

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| folder name | Author details (name, company, twitter alias with link) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.1     | March 10, 2021   | Update comment  |
| 1.0     | January 29, 2021 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Include any additional steps as needed.

## Features

Description of the extension that expands upon high-level summary above.

This extension illustrates the following concepts:

- topic 1
- topic 2
- topic 3

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import styles from './MemberProfiles.module.scss';
import type { IMemberProfilesProps, IProfileItem } from '../models';
import { MemberCard } from './MemberCard';
import { DetailsPanel } from './DetailsPanel';
import { SpService } from '../services/SpService';

/** ES5-safe normalizers for search and library matching */
const toNorm = (s?: string) =>
  (s || '').toString().trim().toLowerCase();

const keyify = (s?: string) =>
  (s || '').toString().toLowerCase().replace(/\.[^.]+$/, '').replace(/[^a-z0-9]/g, '');

export const MemberProfiles: React.FC<IMemberProfilesProps> = (props) => {
  const { listId, itemsPerPage, accentColor, pageTitle, subTitle, imageLibraryId, context } = props;

  const [items, setItems] = React.useState<IProfileItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | null>(null);

  const [qPeople, setQPeople] = React.useState<string>('');
  const [qDetails, setQDetails] = React.useState<string>('');
  const [selected, setSelected] = React.useState<IProfileItem | undefined>(undefined);

  /** Lock background scroll while modal is open */
  React.useEffect(() => {
    if (selected) {
      const prev = document.body.style.overflow;
      document.body.style.overflow = 'hidden';
      return () => { document.body.style.overflow = prev; };
    }
  }, [selected]);

  /** Load data + library fallback */
  React.useEffect(() => {
    let dead = false;
    (async () => {
      try {
        setLoading(true);
        setError(null);

        const svc = new SpService(context);
        const base = await svc.getProfiles(listId);

        // Optional photo library map (keyed by Title and filename)
        let libMap: { [key: string]: string } | null = null;
        if (imageLibraryId) {
          try {
            libMap = await svc.getLibraryPhotoMap(imageLibraryId);
          } catch (e) { /* silent */ }
        }

        // Enrich with library photo when M365/primary is missing
        const enriched: IProfileItem[] = [];
        for (let i = 0; i < base.length; i++) {
          const it: any = base[i];
          if ((!it.photoUrl || it.photoUrl === '') && libMap) {
            const k1 = keyify(it.name);
            const k2 = keyify((it.upn || '').split('@')[0]);
            const candidate = (libMap as any)[k1] || (libMap as any)[k2];
            if (candidate) it.photoUrl = candidate;
          }
          enriched.push(it as IProfileItem);
        }

        if (!dead) setItems(enriched);
      } catch (e: any) {
        if (!dead) setError(e && e.message ? e.message : 'Failed to load profiles');
      } finally {
        if (!dead) setLoading(false);
      }
    })();
    return () => { dead = true; };
  }, [listId, imageLibraryId, context]);

  /** Filtering (ES5-safe; no String.includes) */
  const filtered = React.useMemo(() => {
    const a = toNorm(qPeople);
    const b = toNorm(qDetails);
    if (!a && !b) return items;

    const out: IProfileItem[] = [];
    for (let i = 0; i < items.length; i++) {
      const it: any = items[i];
      const name = toNorm(it.name);
      const role = toNorm(it.role);
      const about = toNorm(it.detailsHtml || '');
      const links = toNorm((it.linkedInUrl || '') + ' ' + (it.companyUrl || '') + ' ' + (it.photoUrl || ''));

      const matchA = !a || name.indexOf(a) > -1 || role.indexOf(a) > -1;
      const matchB = !b || about.indexOf(b) > -1 || links.indexOf(b) > -1;

      if (matchA && matchB) out.push(items[i]);
    }
    return out;
  }, [items, qPeople, qDetails]);

  /** Page size */
  const page = React.useMemo(() => {
    const n = Number(itemsPerPage) || 0;
    if (!n || n < 0) return filtered;
    const out: IProfileItem[] = [];
    for (let i = 0; i < filtered.length && i < n; i++) out.push(filtered[i]);
    return out;
  }, [filtered, itemsPerPage]);

  return (
    <div
      className={styles.wrapper}
      style={accentColor ? ({ ['--accent' as any]: accentColor } as React.CSSProperties) : undefined}
    >
      <div className={styles.container}>
        <div className={styles.headingWrap}>
          <h1 className={styles.heading}>{pageTitle || 'Employee Profiles'}</h1>
          {subTitle ? <p className={styles.subtitle}>{subTitle}</p> : null}
        </div>

        <div className={styles.searchRow}>
          <input
            type="text"
            aria-label="Search by name or role"
            placeholder="Search by name or role..."
            value={qPeople}
            onChange={(e) => setQPeople(e.target.value)}
            style={{ padding: '10px 12px', borderRadius: 8, border: '1px solid #dfe7ef', fontSize: 14 }}
          />
          <input
            type="text"
            aria-label="Search details"
            placeholder="Search about, links..."
            value={qDetails}
            onChange={(e) => setQDetails(e.target.value)}
            style={{ padding: '10px 12px', borderRadius: 8, border: '1px solid #dfe7ef', fontSize: 14 }}
          />
        </div>

        {loading && <div className={styles.loading}>Loading profiles…</div>}
        {error   && <div className={styles.loading}>Error: {error}</div>}

        {!loading && !error && (
          <div className={styles.grid}>
            {page.map((it) => (
              <MemberCard
                key={(it as any).id}
                item={it}
                active={selected ? ((selected as any).id === (it as any).id) : false}
                accentColor={accentColor}
                onClick={setSelected}
              />
            ))}
          </div>
        )}
      </div>

      {/* True modal via portal to avoid being clipped by the web part container */}
      {selected && ReactDOM.createPortal(
        <div className={styles.modalOverlay} onClick={() => setSelected(undefined)}>
          {/* stopPropagation to allow clicks inside the panel */}
          <div onClick={(e) => e.stopPropagation()}>
            <DetailsPanel item={selected} onDismiss={() => setSelected(undefined)} />
          </div>
        </div>,
        document.body
      )}
    </div>
  );
};

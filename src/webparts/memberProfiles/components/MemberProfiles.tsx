// src/webparts/memberProfiles/components/MemberProfiles.tsx
import * as React from 'react';
import styles from './MemberProfiles.module.scss';
import { SearchBox, ISearchBoxStyles } from '@fluentui/react/lib/SearchBox';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { MemberCard } from './MemberCard';
import { DetailsPanel } from './DetailsPanel';
import type { IProfileItem } from '../models';
import { SpService } from '../services/SpService';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { IMemberProfilesProps } from './IMemberProfilesProps';

interface IComponentProps extends IMemberProfilesProps {
  context: WebPartContext;
}

const norm = (s?: string) => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');

export const MemberProfiles: React.FC<IComponentProps> = (props) => {
  const { listId, imageListId, itemsPerPage, accentColor, context } = props;

  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>();
  const [items, setItems] = React.useState<IProfileItem[]>([]);
  const [searchName, setSearchName] = React.useState('');
  const [searchRole, setSearchRole] = React.useState('');
  const [active, setActive] = React.useState<IProfileItem | undefined>();

  const service = React.useMemo(() => new SpService(context), [context]);

  React.useEffect(() => {
    let mounted = true;
    setLoading(true);
    setError(undefined);

    Promise.all([
      service.getProfiles(listId),
      service.getImageMap(imageListId)
    ])
      .then(([profiles, imageMap]) => {
        if (!mounted) return;
        const merged: IProfileItem[] = [];
        for (let i = 0; i < profiles.length; i++) {
          const p = profiles[i];
          const k = norm(p.name);
          const photo = imageMap[k];
          merged.push({ ...p, photoUrl: photo || p.photoUrl });
        }
        setItems(merged);
        setLoading(false);
      })
      .catch((e) => { if (!mounted) return; setError(e?.message || 'Load failed'); setLoading(false); });

    return () => { mounted = false; };
  }, [listId, imageListId, service]);

  const normalized = React.useMemo(() => items.map((i) => ({ ...i, key: i.id })), [items]);

  const filtered = React.useMemo(() => {
    const n = searchName.trim().toLowerCase();
    const r = searchRole.trim().toLowerCase();
    const out: IProfileItem[] = [];
    for (let i = 0; i < normalized.length; i++) {
      const it = normalized[i];
      const byName = !n || ((it.name || '').toLowerCase().indexOf(n) > -1);
      const byRole = !r || ((it.role || '').toLowerCase().indexOf(r) > -1);
      if (byName && byRole) out.push(it);
    }
    return out;
    
  }, [normalized, searchName, searchRole]);

  // 0 or undefined => show ALL
  const page = React.useMemo(
    () => {
      const cap = (itemsPerPage && itemsPerPage > 0) ? itemsPerPage : filtered.length;
      return filtered.slice(0, cap);
    },
    [filtered, itemsPerPage]
  );

  const searchStyles: Partial<ISearchBoxStyles> = {
    root: { borderRadius: 24, border: '1px solid #e5e7eb', height: 40, overflow: 'hidden' },
    field: { fontSize: 14 },
    icon: { fontSize: 16 }
  };

  return (
    <div className={styles.wrapper} style={{ ['--accent' as any]: accentColor }}>
      <div className={styles.container}>
        <div className={styles.headingWrap}>
          <h2 className={styles.heading}>Team Member Profiles</h2>
          <div className={styles.subtitle}>Get to know more about our team!</div>
        </div>

        <div className={styles.searchRow}>
          <SearchBox placeholder="Search by name" styles={searchStyles} onChange={(_, v) => setSearchName(v || '')} />
          <SearchBox placeholder="Search by role/title" styles={searchStyles} onChange={(_, v) => setSearchRole(v || '')} />
        </div>
      </div>

      {error && (<MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>)}
      {loading && (<Spinner size={SpinnerSize.large} label="Loading profiles..." />)}

      {!loading && (
        <div className={styles.grid}>
          {page.map((p) => (
            <MemberCard
              key={p.id}
              item={p}
              onClick={setActive}
              accentColor={accentColor}
              active={active?.id === p.id}
            />
          ))}
        </div>
      )}

      <DetailsPanel
        item={active}
        isOpen={!!active}
        onDismiss={() => setActive(undefined)}
        accentColor={accentColor}
      />
    </div>
  );
};

export default MemberProfiles;

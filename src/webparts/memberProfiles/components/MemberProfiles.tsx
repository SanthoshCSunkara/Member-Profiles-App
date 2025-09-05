import * as React from 'react';
import styles from './MemberProfiles.module.scss';
import { SearchBox, ISearchBoxStyles } from '@fluentui/react/lib/SearchBox';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { MemberCard } from './MemberCard';
import { DetailsPanel } from './DetailsPanel';
import type { IProfileItem, IMemberProfilesProps } from '../models';
import { SpService } from '../services/SpService';
import type { WebPartContext } from '@microsoft/sp-webpart-base';

interface IComponentProps extends IMemberProfilesProps { context: WebPartContext; }

export const MemberProfiles: React.FC<IComponentProps> = (props) => {
  const { listId, itemsPerPage, accentColor, context } = props;

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
    service.getProfiles(listId)
      .then((data) => { if (!mounted) return; setItems(data); setLoading(false); })
      .catch((e) => { if (!mounted) return; setError(e?.message || 'Load failed'); setLoading(false); });
    return () => { mounted = false; };
  }, [listId, service]);

  const normalized = React.useMemo(() => items.map((i) => ({ ...i, key: i.id })), [items]);

  const filtered = React.useMemo(() => {
    const n = searchName.trim().toLowerCase();
    const r = searchRole.trim().toLowerCase();
    return normalized.filter((i) => {
      const byName = !n || ((i.name || '').toLowerCase().indexOf(n) > -1);
      const byRole = !r || ((i.role || '').toLowerCase().indexOf(r) > -1);
      return byName && byRole;
    });
  }, [normalized, searchName, searchRole]);

  const page = React.useMemo(
    () => filtered.slice(0, itemsPerPage || 9999),
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

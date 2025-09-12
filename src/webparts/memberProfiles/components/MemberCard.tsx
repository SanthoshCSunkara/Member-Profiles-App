import * as React from 'react';
import styles from './MemberProfiles.module.scss';
import type { IProfileItem } from '../models';

const AVATAR_PX = 96;

const userPhoto = (upn?: string) => {
  if (!upn) return undefined;
  const origin = window.location.origin.replace(/\/$/, '');
  return origin + '/_layouts/15/userphoto.aspx?size=L&accountname=' + encodeURIComponent(upn);
};

export const MemberCard: React.FC<{
  item: IProfileItem;
  active?: boolean;
  accentColor: string;
  onClick: (i: IProfileItem) => void;
}> = ({ item, active, accentColor, onClick }) => {
  const upn = (item as any).upn as string | undefined;
  const m365 = userPhoto(upn);     // prefer M365 (real photo or default silhouette)
  const lib  = item.photoUrl;      // fallback

  const [src, setSrc] = React.useState<string | undefined>(m365 || lib);
  React.useEffect(() => { setSrc(m365 || lib); }, [m365, lib, (item as any).id]);

  const onError = () => {
    if (src === m365 && lib) { setSrc(lib); return; }
    setSrc(undefined);
  };

  return (
    <button
      type="button"
      className={`${styles.card} ${active ? styles.cardActive : ''}`}
      onClick={() => onClick(item)}
      aria-label={`Open details for ${(item as any).name}`}
      title={`Open details for ${(item as any).name}`}
      style={{ '--accent': accentColor } as React.CSSProperties}
    >
      <div className={styles.row}>
        <div className={styles.avatar}>
          {src ? (
            <img
              src={src}
              alt=""
              width={AVATAR_PX}
              height={AVATAR_PX}
              loading="lazy"
              decoding="async"
              onError={onError}
              style={{ width: '100%', height: '100%', objectFit: 'cover', display: 'block', borderRadius: '50%' }}
            />
          ) : <span aria-hidden="true" />}
        </div>

        <div className={styles.meta}>
          <div className={styles.name}>{(item as any).name}</div>
          {item.role && <div className={styles.role}>{item.role}</div>}
        </div>
      </div>
    </button>
  );
};

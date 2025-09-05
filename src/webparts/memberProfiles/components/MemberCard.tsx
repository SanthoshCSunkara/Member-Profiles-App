import * as React from 'react';
import styles from './MemberProfiles.module.scss';
import type { IProfileItem } from '../models';

interface Props {
  item: IProfileItem;
  active?: boolean;
  accentColor: string;
  onClick: (item: IProfileItem) => void;
}

const AVATAR = 112; // retina-crisp (your CSS still shows 56x56)

/** Safely append width/height/mode to any URL */
const buildPrimary = (raw: string, w: number, h: number) => {
  try {
    const u = new URL(raw, window.location.origin);
    u.searchParams.set('width', String(w));
    u.searchParams.set('height', String(h));
    u.searchParams.set('mode', 'crop');
    return u.toString();
  } catch {
    // if URL() fails, fall back to your original heuristic
    const sep = raw.indexOf('?') > -1 ? '&' : '?';
    return `${raw}${sep}width=${w}&height=${h}&mode=crop`;
  }
};

/** SP preview handler (closest to list-formatting getThumbnailImage) */
const buildFallback = (raw: string, w: number, h: number) =>
  `/_layouts/15/getpreview.ashx?path=${encodeURIComponent(raw)}&width=${w}&height=${h}`;

export const MemberCard: React.FC<Props> = ({ item, active, onClick }) => {
  const baseUrl = item.photoUrl || '';

  const [src, setSrc] = React.useState<string | undefined>(() =>
    baseUrl ? buildPrimary(baseUrl, AVATAR, AVATAR) : undefined
  );
  const [usedFallback, setUsedFallback] = React.useState(false);

  React.useEffect(() => {
    setSrc(baseUrl ? buildPrimary(baseUrl, AVATAR, AVATAR) : undefined);
    setUsedFallback(false);
  }, [baseUrl, item.id]);

  const handleError = () => {
    if (!baseUrl) return;
    if (!usedFallback) {
      setSrc(buildFallback(baseUrl, AVATAR, AVATAR));
      setUsedFallback(true);
    } else {
      setSrc(undefined); // show solid block; no further retries
    }
  };

  return (
    <button
      type="button"
      className={`${styles.card} ${active ? styles.cardActive : ''}`}
      onClick={() => onClick(item)}
      aria-label={`Open details for ${item.name}`}
      title={`Open details for ${item.name}`}
    >
      <div className={styles.row}>
        <div className={styles.avatar}>
          {src ? (
            <img
              src={src}
              alt=""
              loading="lazy"
              decoding="async"
              onError={handleError}
            />
          ) : (
            <span aria-hidden="true" />
          )}
        </div>

        <div className={styles.meta}>
          <div className={styles.name}>{item.name}</div>
          {item.role && <div className={styles.role}>{item.role}</div>}
          
        </div>
      </div>
    </button>
  );
};

// src/webparts/memberProfiles/components/MemberCard.tsx
import * as React from 'react';
import styles from './MemberProfiles.module.scss';
import type { IProfileItem } from '../models';

interface Props {
  item: IProfileItem;
  active?: boolean;
  accentColor: string;
  onClick: (item: IProfileItem) => void;
}

/** Safe query-appender for width/height/mode=crop */
const buildPrimary = (raw: string, w: number, h: number) => {
  try {
    const u = new URL(raw, window.location.origin);
    u.searchParams.set('width', String(w));
    u.searchParams.set('height', String(h));
    u.searchParams.set('mode', 'crop');
    return u.toString();
  } catch {
    const sep = raw.indexOf('?') > -1 ? '&' : '?';
    return `${raw}${sep}width=${w}&height=${h}&mode=crop`;
  }
};

/** SP preview handler – last resort */
const buildPreview = (raw: string, w: number, h: number) =>
  `/_layouts/15/getpreview.ashx?path=${encodeURIComponent(raw)}&width=${w}&height=${h}`;

export const MemberCard: React.FC<Props> = ({ item, active, onClick }) => {
  const baseUrl = item.photoUrl || '';
  const CSS_SIZE = 56;                 // visual size in CSS
  const DPR = Math.min(2, Math.ceil(window.devicePixelRatio || 1));
  const REND_1X = CSS_SIZE;            // 56
  const REND_2X = CSS_SIZE * DPR;      // 112 on HiDPI

  const primary = baseUrl ? buildPrimary(baseUrl, REND_2X, REND_2X) : undefined;
  const oneX = baseUrl ? buildPrimary(baseUrl, REND_1X, REND_1X) : undefined;

  // 0 = primary, 1 = original, 2 = preview; then solid color
  const [src, setSrc] = React.useState<string | undefined>(primary);
  const [phase, setPhase] = React.useState<0 | 1 | 2>(0);

  React.useEffect(() => {
    const p = baseUrl ? buildPrimary(baseUrl, REND_2X, REND_2X) : undefined;
    setSrc(p);
    setPhase(0);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [item.id, baseUrl]);

  const handleError = () => {
    if (!baseUrl) return;
    if (phase === 0) { setSrc(baseUrl); setPhase(1); return; }
    if (phase === 1) { setSrc(buildPreview(baseUrl, REND_2X, REND_2X)); setPhase(2); return; }
    setSrc(undefined);
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
              srcSet={oneX && primary ? `${oneX} 1x, ${primary} 2x` : undefined}
              sizes={`${CSS_SIZE}px`}
              width={CSS_SIZE}
              height={CSS_SIZE}
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

// src/webparts/memberProfiles/components/MemberCard.tsx
import * as React from 'react';
import styles from './MemberProfiles.module.scss';
import type { IProfileItem } from '../models';

interface IProps {
  item: IProfileItem;
  onClick: (item: IProfileItem) => void;
  active?: boolean;
  accentColor?: String
}

/** CSS size in the layout (kept as 96px so your UI does not change). */
const AVATAR_CSS_PX = 96;

/**
 * Some endpoints (e.g., Graph /photo/$value) ignore resize query params.
 * We only add ?width=&height=&quality= when it’s a library URL.
 */
const canResizeViaQuery = (url?: string) =>
  !!url && !/\/photo\/\$value/i.test(url || '');

/** Build a preview URL with target pixel size. */
const buildPreview = (url: string, px: number) => {
  if (!url || !canResizeViaQuery(url)) return url;
  const sep = url.indexOf('?') >= 0 ? '&' : '?';
  return `${url}${sep}width=${px}&height=${px}&quality=90`;
};

/** Safe initials fallback */
const initials = (name?: string) => {
  const t = (name || '').trim();
  if (!t) return '';
  const parts = t.split(/\s+/).slice(0, 2);
  return parts.map(p => p[0]?.toUpperCase() || '').join('');
};

export const MemberCard: React.FC<IProps> = ({ item, onClick, active }) => {
  const base = item.photoUrl || '';

  // Device pixel ratio aware fetch sizes: always give the browser a 2x version,
  // and a 1.5x for mid-DPR screens. This closely mirrors your portrait branch.
  const dpr = typeof window !== 'undefined' ? Math.min(3, window.devicePixelRatio || 1) : 1;
  const fetch1x = Math.round(AVATAR_CSS_PX * Math.max(1, dpr >= 1.5 ? 1.5 : 1)); // 96 or 144
  const fetch2x = Math.round(AVATAR_CSS_PX * Math.min(2, Math.ceil(dpr)));        // 192

  const src1x = buildPreview(base, fetch1x);
  const src2x = buildPreview(base, fetch2x);

  const hasPhoto = !!base;

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
          {hasPhoto ? (
            <img
              // important: give the browser larger pixels than we render
              src={src1x}
              srcSet={
                src2x && src2x !== src1x ? `${src1x} 1x, ${src2x} 2x` : undefined
              }
              sizes={`${AVATAR_CSS_PX}px`}
              // lock raster to the pixel grid (prevents soft resampling)
              width={AVATAR_CSS_PX}
              height={AVATAR_CSS_PX}
              alt={`${item.name} photo`}
              loading="lazy"
              decoding="async"
              style={{
                backfaceVisibility: 'hidden',
                WebkitBackfaceVisibility: 'hidden',
                transform: 'translateZ(0)' // create a stable layer, no animation
              }}
            />
          ) : (
            <div
              style={{
                width: AVATAR_CSS_PX,
                height: AVATAR_CSS_PX,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                color: '#E8F0F7',
                background: '#2d3748',
                borderRadius: '50%',
                fontWeight: 800
              }}
            >
              {initials(item.name)}
            </div>
          )}
        </div>

        <div className={styles.meta}>
          <div className={styles.name}>{item.name}</div>
          <div className={styles.role}>{item.role}</div>
        </div>
      </div>
    </button>
  );
};

export default MemberCard;

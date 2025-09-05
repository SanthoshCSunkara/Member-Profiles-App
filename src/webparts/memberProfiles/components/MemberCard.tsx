import * as React from 'react';
import styles from './MemberProfiles.module.scss';
import type { IProfileItem } from '../models';

interface Props {
  item: IProfileItem;
  active?: boolean;
  accentColor: string;
  onClick: (item: IProfileItem) => void;
}

/** Primary preview: try modern width/height query */
const buildPrimary = (url: string, w: number, h: number) => {
  const sep = url.indexOf('?') > -1 ? '&' : '?';
  return `${url}${sep}width=${w}&height=${h}&mode=crop`;
};

/** Fallback preview: SharePoint preview handler (closest to list-formatting getThumbnailImage) */
const buildFallback = (url: string, w: number, h: number) =>
  `/_layouts/15/getpreview.ashx?path=${encodeURIComponent(url)}&width=${w}&height=${h}`;

export const MemberCard: React.FC<Props> = ({ item, active, onClick }) => {
  const baseUrl = item.photoUrl || '';
  const [src, setSrc] = React.useState<string | undefined>(() =>
    baseUrl ? buildPrimary(baseUrl, 96, 96) : undefined
  );
  const [triedFallback, setTriedFallback] = React.useState(false);

  React.useEffect(() => {
    setSrc(baseUrl ? buildPrimary(baseUrl, 96, 96) : undefined);
    setTriedFallback(false);
  }, [baseUrl, item.id]);

  const handleError = () => {
    if (!baseUrl) return;
    if (!triedFallback) {
      setSrc(buildFallback(baseUrl, 96, 96)); // try SP preview pipeline
      setTriedFallback(true);
    } else {
      setSrc(undefined); // show color block
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
            <img src={src} alt="" loading="lazy" onError={handleError} />
          ) : (
            // Solid color block fallback (no letter), to match your JSON
            <span aria-hidden="true" />
          )}
        </div>

        <div className={styles.meta}>
          <div className={styles.name}>{item.name}</div>
          {item.role && <div className={styles.role}>{item.role}</div>}
          <div className={styles.viewMoreInline}>View more details…</div>
        </div>
      </div>
    </button>
  );
};

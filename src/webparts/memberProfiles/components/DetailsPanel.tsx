import * as React from 'react';
import styles from './MemberProfiles.module.scss';
import type { IProfileItem } from '../models';

const userPhoto = (upn?: string) =>
  upn ? `${window.location.origin}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(upn)}` : undefined;

export const DetailsPanel: React.FC<{
  item?: IProfileItem;
  onDismiss: () => void;
}> = ({ item, onDismiss }) => {
  if (!item) return null;

  const upn = (item as any).upn as string | undefined;
  const m365 = userPhoto(upn);
  const lib  = item.photoUrl;

  const [src, setSrc] = React.useState<string | undefined>(m365 || lib);
  React.useEffect(() => setSrc(m365 || lib), [m365, lib, (item as any).id]);

  const onError = () => { if (src === m365 && lib) setSrc(lib); };

  return (
    <div className={styles.modalContainer} role="dialog" aria-modal="true" aria-label="Profile details">
      <div className={styles.modalHeader}>
        <div className={styles.modalTitle}>{(item as any).name}</div>
        <button onClick={onDismiss} aria-label="Close">✕</button>
      </div>

      <div className={styles.modalBody}>
        <div className={styles.modalLeft}>
          <div className={styles.modalCard}>
            <div className={styles.modalName}>{(item as any).name}</div>
            {item.role && <div className={styles.modalRole}>{item.role}</div>}
            {src && <img className={styles.modalImage} src={src} alt="" loading="eager" decoding="async" onError={onError} />}
          </div>
        </div>

        <div className={styles.modalRight}>
          <div className={styles.detailsHtml} dangerouslySetInnerHTML={{ __html: item.detailsHtml || '' }} />
        </div>
      </div>
    </div>
  );
};

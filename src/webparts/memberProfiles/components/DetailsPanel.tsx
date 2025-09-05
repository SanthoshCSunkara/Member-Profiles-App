import * as React from 'react';
import * as DOMPurifyNS from 'dompurify';
const createDOMPurify: any = (DOMPurifyNS as any).default || (DOMPurifyNS as any);
const DOMPurify = createDOMPurify(window as any);

import { Modal } from '@fluentui/react/lib/Modal';
import { IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import type { IProfileItem } from '../models';
import styles from './MemberProfiles.module.scss';

const buildPrimary = (url: string, w: number, h: number) => {
  const sep = url.indexOf('?') > -1 ? '&' : '?';
  return `${url}${sep}width=${w}&height=${h}&mode=crop`;
};
const buildFallback = (url: string, w: number, h: number) =>
  `/_layouts/15/getpreview.ashx?path=${encodeURIComponent(url)}&width=${w}&height=${h}`;

interface IDetailsPanelProps {
  item?: IProfileItem;
  isOpen: boolean;
  onDismiss: () => void;
  accentColor: string;
}

export const DetailsPanel: React.FC<IDetailsPanelProps> = ({ item, isOpen, onDismiss, accentColor }) => {
  const baseUrl = item?.photoUrl || '';
  const [src, setSrc] = React.useState<string | undefined>(() =>
    baseUrl ? buildPrimary(baseUrl, 600, 600) : undefined
  );
  const [triedFallback, setTriedFallback] = React.useState(false);

  React.useEffect(() => {
    setSrc(baseUrl ? buildPrimary(baseUrl, 600, 600) : undefined);
    setTriedFallback(false);
  }, [baseUrl, item?.id]);

  const handleError = () => {
    if (!baseUrl) return;
    if (!triedFallback) {
      setSrc(buildFallback(baseUrl, 600, 600));
      setTriedFallback(true);
    } else {
      setSrc(undefined);
    }
  };

  const safeHtml = React.useMemo(
    () => ({ __html: DOMPurify.sanitize(item?.detailsHtml || '') }),
    [item]
  );

  return (
    <Modal
      isOpen={isOpen}
      onDismiss={onDismiss}
      isBlocking={false}
      containerClassName={styles.modalContainer}
    >
      <div className={styles.modalHeader}>
        <div className={styles.modalTitle}>{item?.name || ''}</div>
        <IconButton ariaLabel="Close" iconProps={{ iconName: 'Cancel' }} onClick={onDismiss} />
      </div>

      {item && (
        <div className={styles.modalBody} style={{ ['--accent' as any]: accentColor }}>
          <div className={styles.modalLeft}>
            <div className={styles.modalCard}>
              {/* Photo first, like your JSON details card */}
              {src ? (
                <img className={styles.modalImage} src={src} alt={item.name} onError={handleError} />
              ) : null}

              <div className={styles.modalName}>{item.name}</div>
              {item.role && <div className={styles.modalRole}>{item.role}</div>}

              <div className={styles.modalMeta}>
                {item.hireDate && <div><Icon iconName="Calendar" /> Hire Date: {item.hireDate}</div>}
                {item.birthday && <div><Icon iconName="BirthdayCake" /> {item.birthday}</div>}
              </div>
            </div>
          </div>

          <div className={styles.modalRight}>
            <div className={styles.detailsHtml} dangerouslySetInnerHTML={safeHtml} />
          </div>
        </div>
      )}
    </Modal>
  );
};

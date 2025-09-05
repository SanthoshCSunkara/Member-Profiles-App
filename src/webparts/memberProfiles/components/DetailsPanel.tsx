import * as React from 'react';
import * as DOMPurifyNS from 'dompurify';
const createDOMPurify: any = (DOMPurifyNS as any).default || (DOMPurifyNS as any);
const DOMPurify = createDOMPurify(window as any);

import { Modal } from '@fluentui/react/lib/Modal';
import { IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import type { IProfileItem } from '../models';
import styles from './MemberProfiles.module.scss';

interface IDetailsPanelProps {
  item?: IProfileItem;
  isOpen: boolean;
  onDismiss: () => void;
  accentColor: string;
}

/** Safely append width/height/mode to any URL for a crisp rendition */
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

/** SharePoint preview handler fallback (closest to list-formatting helper) */
const buildFallback = (raw: string, w: number, h: number) =>
  `/_layouts/15/getpreview.ashx?path=${encodeURIComponent(raw)}&width=${w}&height=${h}`;

export const DetailsPanel: React.FC<IDetailsPanelProps> = ({ item, isOpen, onDismiss, accentColor }) => {
  const safeHtml = React.useMemo(
    () => ({ __html: DOMPurify.sanitize(item?.detailsHtml || '') }),
    [item]
  );

  // --- CRISP / FALLBACK IMAGE LOGIC (minimal, self-contained) ---
  const base = item?.photoUrl || '';
  // use a generous square so it stays sharp; CSS will cap height
  const TARGET = 720;

  const [panelSrc, setPanelSrc] = React.useState<string | undefined>(() =>
    base ? buildPrimary(base, TARGET, TARGET) : undefined
  );
  const [usedFallback, setUsedFallback] = React.useState(false);

  React.useEffect(() => {
    const b = item?.photoUrl || '';
    setPanelSrc(b ? buildPrimary(b, TARGET, TARGET) : undefined);
    setUsedFallback(false);
  }, [item?.id, item?.photoUrl]);

  const onImgError = () => {
    if (!base) return;
    if (!usedFallback) {
      setPanelSrc(buildFallback(base, TARGET, TARGET));
      setUsedFallback(true);
    } else {
      setPanelSrc(base); // last resort: original
    }
  };
  // --------------------------------------------------------------

  return (
    <Modal
      isOpen={isOpen}
      onDismiss={onDismiss}
      isBlocking={false}
      containerClassName={styles.modalContainer}
    >
      <div className={styles.modalHeader} style={{ ['--accent' as any]: accentColor }}>
        <div className={styles.modalTitle}>{item?.name || ''}</div>
        <IconButton ariaLabel="Close" iconProps={{ iconName: 'Cancel' }} onClick={onDismiss} />
      </div>

      {item && (
        <div className={styles.modalBody} style={{ ['--accent' as any]: accentColor }}>
          {/* Left summary card */}
          <div className={styles.modalLeft}>
            <div className={styles.modalCard}>
              <div className={styles.modalName}>{item.name}</div>
              {item.role && <div className={styles.modalRole}>{item.role}</div>}

              <div className={styles.modalMeta}>
                {item.hireDate && <div><Icon iconName="Calendar" /> Hire Date: {item.hireDate}</div>}
                {item.birthday && <div><Icon iconName="BirthdayCake" /> {item.birthday}</div>}
              </div>

              {(item.companyUrl || item.linkedInUrl) && (
                <div className={styles.btnRow}>
                  {item.companyUrl && (
                    <a
                      className={styles.btnGhost}
                      href={item.companyUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                    >
                      <Icon iconName="OpenFile" /> Company Profile
                    </a>
                  )}
                  {item.linkedInUrl && (
                    <a
                      className={styles.btnGhost}
                      href={item.linkedInUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                    >
                      <Icon iconName="LinkedInLogo" /> LinkedIn
                    </a>
                  )}
                </div>
              )}

              {panelSrc && (
                <img
                  className={styles.modalImage}
                  src={panelSrc}
                  alt={item.name}
                  loading="lazy"
                  onError={onImgError}
                />
              )}
            </div>
          </div>

          {/* Right rich text content */}
          <div className={styles.modalRight}>
            <div className={styles.detailsHtml} dangerouslySetInnerHTML={safeHtml} />
          </div>
        </div>
      )}
    </Modal>
  );
};

import * as React from 'react';
import styles from './PortraitCard.module.scss';

/** SharePoint Preview (crisp, same pipeline as modal) */
const preview = (raw: string, w: number, h: number) =>
  `/_layouts/15/getpreview.ashx?path=${encodeURIComponent(raw)}&width=${w}&height=${h}`;

const buildSrcSet = (raw: string, w: number, h: number) =>
  `${preview(raw, w, h)} 1x, ${preview(raw, w * 2, h * 2)} 2x`;

export const PortraitCard: React.FC<{
  name: string;
  role?: string;
  photoUrl?: string;   // server-relative FileRef
  onClick?: () => void;
}> = ({ name, role, photoUrl, onClick }) => {
  // Fetch generous previews so the final 220px tall crop is razor sharp
  const W = 720, H = 720;

  const base = photoUrl || '';
  const [src, setSrc] = React.useState<string | undefined>(base ? preview(base, W, H) : undefined);
  const [srcset, setSrcset] = React.useState<string | undefined>(base ? buildSrcSet(base, W, H) : undefined);

  React.useEffect(() => {
    if (!base) { setSrc(undefined); setSrcset(undefined); return; }
    setSrc(preview(base, W, H));
    setSrcset(buildSrcSet(base, W, H));
  }, [base]);

  const onError = () => {
    if (!base) return;
    setSrc(base);        // last-resort: raw image
    setSrcset(undefined);
  };

  const initials = name.split(/\s+/).map(p => p[0]).slice(0, 2).join('').toUpperCase();

  return (
    <button
      className={styles.card}
      onClick={onClick}
      aria-label={name}
    >
      {src ? (
        <img
          className={styles.photo}
          src={src}
          srcSet={srcset}
          sizes="(min-width:1200px) 360px, (min-width:720px) 45vw, 92vw"
          alt={name}
          onError={onError}
          loading="lazy"
          decoding="async"
        />
      ) : (
        <div
          className={styles.photo}
          style={{
            background:'#e9eef3',
            display:'flex', alignItems:'center', justifyContent:'center',
            color:'#334155', fontWeight:700, fontSize:26
          }}
        >
          {initials}
        </div>
      )}

      <div className={styles.body}>
        <h3 className={styles.name} title={name}>{name}</h3>
        {role ? <p className={styles.role} title={role}>{role}</p> : null}
      </div>
    </button>
  );
};

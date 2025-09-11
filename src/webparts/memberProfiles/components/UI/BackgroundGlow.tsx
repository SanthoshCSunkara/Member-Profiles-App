import * as React from 'react';
import styles from './BackgroundGlow.module.scss';

export const BackgroundGlow: React.FC<{
  children: React.ReactNode;
  className?: string;            // added to .inner
  containerClassName?: string;   // added to .container
}> = ({ children, className, containerClassName }) => (
  <div className={`${styles.container} ${containerClassName || ''}`}>
    <div className={styles.glow} />
    <div className={`${styles.inner} ${className || ''}`}>{children}</div>
  </div>
);

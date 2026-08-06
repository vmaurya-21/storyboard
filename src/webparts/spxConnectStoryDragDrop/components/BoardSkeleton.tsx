import * as React from 'react';
import styles from './BoardSkeleton.module.scss';

/**
 * Placeholder shown while stories load from SharePoint.
 *
 * Mirrors the loaded page — toolbar, board grid, then the scheduled and
 * available columns — so the layout does not shift when data arrives. Exposed
 * as `role="status"` with `aria-busy`, which is also how the tests detect the
 * loading state.
 */
const BoardSkeleton: React.FC = () => {
  const rows = [0, 1, 2, 3, 4];
  const dots = [0, 1, 2, 3, 4];

  return (
    <div className={styles.skeleton} role="status" aria-busy="true">
      <span className={styles.srOnly}>Loading stories…</span>

      <div className={styles.header}>
        <div className={styles.headerGroup}>
          <div className={`${styles.block} ${styles.btn}`} />
          <div className={`${styles.block} ${styles.btn}`} />
        </div>
        <div className={styles.headerGroup}>
          <div className={`${styles.block} ${styles.btn}`} />
          <div className={`${styles.block} ${styles.btn}`} />
        </div>
      </div>

      <div className={styles.carouselBoard}>
        <div className={styles.carouselViewport}>
          <div className={`${styles.block} ${styles.carouselSlide}`} />
          <div className={styles.carouselDots}>
            {dots.map((dot) => (
              <div key={dot} className={`${styles.block} ${styles.carouselDot}`} />
            ))}
          </div>
        </div>
        <div className={styles.carouselNav}>
          <div className={`${styles.block} ${styles.navBtn}`} />
          <div className={`${styles.block} ${styles.navBtn}`} />
        </div>
      </div>

      <div className={styles.bottom}>
        <div className={styles.scheduled}>
          <div className={`${styles.block} ${styles.sectionTitle}`} />
          <div className={`${styles.block} ${styles.groupCard}`} />
          <div className={`${styles.block} ${styles.groupCard}`} />
        </div>

        <div className={styles.available}>
          <div className={`${styles.block} ${styles.sectionTitle}`} />
          <div className={`${styles.block} ${styles.search}`} />
          {rows.map(i => (
            <div key={i} className={styles.row}>
              <div className={`${styles.block} ${styles.rowThumb}`} />
              <div className={styles.rowLines}>
                <div className={`${styles.block} ${styles.lineWide}`} />
                <div className={`${styles.block} ${styles.lineNarrow}`} />
              </div>
              <div className={`${styles.block} ${styles.rowAction}`} />
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

export default BoardSkeleton;

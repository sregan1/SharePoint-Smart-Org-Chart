import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import { GraphService, IGraphUser, IOrgNode, PresenceAvailability } from '../../../../services/GraphService';
import { getInitials } from '../personUtils';
import { OrgChartTheme, getThemeTokens } from './orgTheme';
import { OrgNodeCard, PresenceDot } from './OrgNodeCard';
import styles from './OrgChart.module.scss';

/* ── Drill-down view ─────────────────────── */

export interface IDrillViewProps {
  drillPath: IGraphUser[];
  /** Direct reports of the current person, already filtered for visibility */
  visibleReports: IGraphUser[];
  allReports: IGraphUser[];
  drillLoadingId: string | null;
  photos: { [id: string]: string | null };
  presenceMap: Map<string, PresenceAvailability>;
  drillReportCounts: Map<string, number>;
  graphService: GraphService | null;
  theme: OrgChartTheme;
  accentColor?: string;
  showDepartment: boolean;
  showOffice: boolean;
  compactCards: boolean;
  /** Re-rooting in progress — dims the view (the data stays on screen) */
  isBusy: boolean;
  headerRef: React.RefObject<HTMLDivElement>;
  onReturnToRoot: () => void;
  onNavigate: (index: number) => void;
  onDrillInto: (user: IGraphUser) => void;
  onDrillToggle: (node: IOrgNode) => void;
  onShowProfile: (user: IGraphUser) => void;
}

const fadeInPhoto = (e: React.SyntheticEvent<HTMLImageElement>): void => {
  e.currentTarget.style.opacity = '1';
};

export const DrillView: React.FC<IDrillViewProps> = ({
  drillPath, visibleReports, allReports, drillLoadingId, photos, presenceMap, drillReportCounts,
  graphService, theme, accentColor, showDepartment, showOffice, compactCards, isBusy, headerRef,
  onReturnToRoot, onNavigate, onDrillInto, onDrillToggle, onShowProfile,
}) => {
  const t = getThemeTokens(theme, accentColor);
  const currentUser = drillPath.length > 0 ? drillPath[drillPath.length - 1] : null;

  // Cards are memoized, so their synthetic nodes must keep a stable identity
  // between renders (a new object each render would defeat React.memo).
  const nodeCache = React.useRef(new Map<string, IOrgNode>());
  const getNode = (report: IGraphUser, childrenLoaded: boolean): IOrgNode => {
    const cached = nodeCache.current.get(report.id);
    if (cached && cached.user === report && cached.childrenLoaded === childrenLoaded) return cached;
    const node: IOrgNode = {
      user: report, directReports: [], isExpanded: false, childrenLoaded, level: 1,
      totalReportCount: graphService?.getTotalReportCount(report.id) ?? 0,
    };
    nodeCache.current.set(report.id, node);
    return node;
  };

  const initialsStyle = { background: t.accent, color: t.onAccent };

  return (
    <div className={`${styles.drillView} ${isBusy ? styles.refocusing : ''}`} aria-busy={isBusy}>

      {/* Breadcrumb nav — only shown when drilled deeper than root */}
      {drillPath.length > 1 && (
        <nav className={styles.drillNav} aria-label="Reporting line">
          <button
            className={styles.drillNavHomeBtn}
            onClick={onReturnToRoot}
            title="Back to top"
            aria-label="Back to top of the org chart"
          >
            <Icon iconName="Home" />
          </button>
          <Icon iconName="ChevronRight" className={styles.drillNavChevron} />
          {drillPath.slice(0, -1).map((person, i) => (
            <React.Fragment key={`${person.id}-${i}`}>
              <button
                className={styles.drillNavItem}
                onClick={() => onNavigate(i)}
                title={`Go back to ${person.displayName}`}
              >
                <span className={styles.drillNavInitials} style={initialsStyle}>
                  {getInitials(person.displayName)}
                </span>
                <span className={styles.drillNavName}>{person.displayName.split(' ')[0]}</span>
              </button>
              <Icon iconName="ChevronRight" className={styles.drillNavChevron} />
            </React.Fragment>
          ))}
          {currentUser && (
            <span className={styles.drillNavCurrent} aria-current="page">
              <span className={styles.drillNavInitials} style={initialsStyle}>
                {getInitials(currentUser.displayName)}
              </span>
              <span className={styles.drillNavName}>{currentUser.displayName}</span>
            </span>
          )}
        </nav>
      )}

      {/* Current person header — receives focus after a Focus action */}
      {currentUser && (
        <div
          ref={headerRef}
          tabIndex={-1}
          className={styles.drillCurrentHeader}
          aria-label={`${currentUser.displayName}${currentUser.jobTitle ? `, ${currentUser.jobTitle}` : ''}`}
          role="group"
          style={{
            background: t.headerBg,
            borderBottom: `3px solid ${t.accent}`,
            borderTop: `1px solid ${t.border}`,
          }}
        >
          <div className={styles.drillCurrentAvatar} style={{ position: 'relative' }}>
            {photos[currentUser.id]
              ? <img
                  src={photos[currentUser.id] as string}
                  alt=""
                  className={styles.drillCurrentAvatarImg}
                  style={{ opacity: 0, transition: 'opacity 0.35s ease' }}
                  onLoad={fadeInPhoto}
                />
              : <div className={styles.drillCurrentAvatarInitials} style={initialsStyle} aria-hidden="true">
                  {getInitials(currentUser.displayName)}
                </div>
            }
            <PresenceDot status={presenceMap.get(currentUser.id)} />
          </div>
          <div className={styles.drillCurrentInfo}>
            <div className={styles.drillCurrentName} style={{ color: t.nameColor }}>
              {currentUser.displayName}
            </div>
            {currentUser.jobTitle && (
              <div className={styles.drillCurrentTitle} style={{ color: t.accentText }}>
                {currentUser.jobTitle}
              </div>
            )}
            {showDepartment && currentUser.department && (
              <div className={styles.drillCurrentDept} style={{ color: t.deptText }}>
                {currentUser.department}
              </div>
            )}
          </div>
          <button
            className={styles.drillCurrentProfileBtn}
            onClick={() => onShowProfile(currentUser)}
            title="View profile"
            aria-label={`View ${currentUser.displayName}'s profile`}
          >
            <Icon iconName="Contact" />
            <span>Profile</span>
          </button>
        </div>
      )}

      {/* Direct reports grid */}
      <div className={styles.drillBody}>
        {drillLoadingId && !allReports.find(u => u.id === drillLoadingId) ? (
          <div className={styles.drillSpinner}>
            <Spinner size={SpinnerSize.medium} label="Loading..." />
          </div>
        ) : visibleReports.length > 0 ? (
          <>
            <div className={styles.drillSectionTitle}>
              Direct Reports &nbsp;
              <span className={styles.drillSectionCount}>{visibleReports.length}</span>
            </div>
            <div
              className={`${styles.drillReportsGrid} ${compactCards ? styles.compactMode : ''}`}
              role="list"
              aria-label={currentUser ? `Direct reports of ${currentUser.displayName}` : 'Direct reports'}
            >
              {visibleReports.map(report => {
                const count = drillReportCounts.get(report.id);
                const countKnown = count !== undefined;
                const noReports = countKnown && count === 0;
                const name = report.displayName;
                const node = getNode(report, noReports);
                // Clicking a drill card drills in (or opens the profile when
                // there's nobody below); the corner button opens the profile
                return (
                  <OrgNodeCard
                    key={report.id}
                    node={node}
                    photo={photos[report.id]}
                    presence={presenceMap.get(report.id)}
                    showDepartment={showDepartment}
                    showOffice={showOffice}
                    isExpanding={drillLoadingId === report.id}
                    isHighlighted={false}
                    theme={theme}
                    accentColor={accentColor}
                    directReportCount={count ?? 0}
                    totalReportCount={node.totalReportCount}
                    compactCards={compactCards}
                    variant="grid"
                    cardActionLabel={noReports ? `View ${name}'s profile` : `Show ${name}'s direct reports`}
                    cornerActionLabel={`View ${name}'s profile`}
                    cornerActionIcon="Contact"
                    expandActionLabel={`Show ${name}'s direct reports`}
                    onToggle={onDrillToggle}
                    onCardClick={onDrillInto}
                    onFocus={onShowProfile}
                  />
                );
              })}
            </div>
          </>
        ) : (
          <div className={styles.drillNoReports}>
            No direct reports
          </div>
        )}
      </div>
    </div>
  );
};

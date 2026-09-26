import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IGraphUser, IOrgNode, PresenceAvailability } from '../../../../services/GraphService';
import { PRESENCE_COLOR, getInitials } from '../personUtils';
import { OrgChartTheme, getThemeTokens } from './orgTheme';
import { formatString } from '../localeUtils';
import * as strings from 'SmartOrgChartWebPartStrings';
import styles from './OrgChart.module.scss';

/* ── Presence dot ────────────────────────── */

export const PresenceDot: React.FC<{ status: PresenceAvailability | undefined }> = ({ status }) => {
  if (!status || status === 'Unknown') return null;
  return <span className={styles.presenceDot} style={{ background: PRESENCE_COLOR[status] }} />;
};

/* ── Node card ───────────────────────────── */

export interface IOrgNodeCardProps {
  node: IOrgNode;
  photo: string | null | undefined;
  presence: PresenceAvailability | undefined;
  showDepartment: boolean;
  showOffice: boolean;
  isExpanding: boolean;
  isHighlighted: boolean;
  theme: OrgChartTheme;
  accentColor?: string;
  directReportCount: number;
  totalReportCount: number;
  managerUser?: IGraphUser;
  compactCards: boolean;
  /** 'tree': a treeitem in the tree layouts (roving tabindex). 'grid': an item in the drill-down grid. */
  variant: 'tree' | 'grid';
  /** Tree semantics (tree variant only) */
  ariaLevel?: number;
  ariaSetSize?: number;
  ariaPosInSet?: number;
  /** Roving tabindex: true for the single tree card that is in the Tab order */
  isTabStop?: boolean;
  /** What activating the card does (tooltip + accessible name). Default: view profile. */
  cardActionLabel?: string;
  /** What the top-right corner button does. Default: focus the chart on this person. */
  cornerActionLabel?: string;
  cornerActionIcon?: string;
  /** Label for the bottom expand button. Default: expand/collapse direct reports. */
  expandActionLabel?: string;
  onToggle: (node: IOrgNode) => void;
  onCardClick: (user: IGraphUser) => void;
  onFocus: (user: IGraphUser) => void;
  onItemFocus?: (userId: string) => void;
}

const fadeInPhoto = (e: React.SyntheticEvent<HTMLImageElement>): void => {
  e.currentTarget.style.opacity = '1';
};

const OrgNodeCardInner: React.FC<IOrgNodeCardProps> = ({
  node, photo, presence, showDepartment, showOffice, isExpanding, isHighlighted,
  theme, accentColor, directReportCount, totalReportCount, managerUser, compactCards, variant,
  ariaLevel, ariaSetSize, ariaPosInSet, isTabStop,
  cardActionLabel, cornerActionLabel, cornerActionIcon, expandActionLabel,
  onToggle, onCardClick, onFocus, onItemFocus,
}) => {
  const { user } = node;
  const t         = getThemeTokens(theme, accentColor);
  const initials  = getInitials(user.displayName);
  const isRoot    = node.level === 0;
  const isTree    = variant === 'tree';

  const hasReports = node.directReports.length > 0 || !node.childrenLoaded;
  const isDisabled = user.accountEnabled === false;
  const isGuest    = user.userType === 'Guest';

  const levelClass   = isRoot ? styles.rootCard : node.level === 1 ? styles.level1Card : '';
  const compactClass = compactCards ? styles.compactCard : '';
  const classes      = [styles.nodeCard, levelClass, compactClass, isHighlighted ? styles.highlightedCard : ''].filter(Boolean).join(' ');

  const borderStyle = theme === 'minimal'
    ? { borderLeft: `3px solid ${t.accent}`, borderTop: '1px solid #d8d8d8', opacity: isDisabled ? 0.55 : 1 }
    : { borderTopColor: t.accent, opacity: isDisabled ? 0.55 : 1 };

  const actionLabel = cardActionLabel || formatString(strings.OrgNode_ViewProfile, { name: user.displayName });
  const cornerLabel = cornerActionLabel || formatString(strings.OrgNode_FocusOrgChartOn, { name: user.displayName });
  const expandLabel = expandActionLabel ||
    `${node.isExpanded ? strings.OrgNode_Collapse : strings.OrgNode_Expand} ${formatString(strings.OrgNode_DirectReportsSuffix, { name: user.displayName })}`;
  const accessibleName = `${user.displayName}${user.jobTitle ? `, ${user.jobTitle}` : ''}. ${actionLabel}`;
  // Inner buttons follow the card's tab stop so Tab moves out of the tree
  // instead of walking every card's buttons.
  const innerTabIndex = isTree && !isTabStop ? -1 : undefined;

  const handleKeyDown = (e: React.KeyboardEvent<HTMLDivElement>): void => {
    // Keys pressed on the inner Expand/Focus buttons belong to those buttons
    if (e.target !== e.currentTarget) return;
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onCardClick(user); }
  };

  return (
    <div
      className={classes}
      style={{ ...borderStyle, cursor: 'pointer', position: 'relative' }}
      title={actionLabel}
      onClick={() => onCardClick(user)}
      onKeyDown={handleKeyDown}
      onFocus={onItemFocus ? (e => { if (e.target === e.currentTarget) onItemFocus(user.id); }) : undefined}
      role={isTree ? 'treeitem' : 'listitem'}
      aria-level={isTree ? ariaLevel : undefined}
      aria-setsize={isTree ? ariaSetSize : undefined}
      aria-posinset={isTree ? ariaPosInSet : undefined}
      aria-expanded={isTree && hasReports ? node.isExpanded : undefined}
      aria-busy={isExpanding || undefined}
      tabIndex={isTree ? (isTabStop ? 0 : -1) : 0}
      aria-label={accessibleName}
      data-soc-id={user.id}
    >
      {/* Corner action — shown on hover / keyboard focus */}
      <button
        className={styles.focusBtn}
        onClick={e => { e.stopPropagation(); onFocus(user); }}
        title={cornerLabel}
        aria-label={cornerLabel}
        tabIndex={innerTabIndex}
      >
        <Icon iconName={cornerActionIcon || 'FitPage'} />
      </button>

      <div className={styles.nodeAvatar}>
        {photo
          ? <img src={photo} alt="" className={styles.photo} style={{ opacity: 0, transition: 'opacity 0.35s ease' }} onLoad={fadeInPhoto} />
          : <div className={styles.initials} style={{ background: t.accent, color: t.onAccent }} aria-hidden="true">{initials}</div>
        }
        <PresenceDot status={presence} />
      </div>
      <div className={styles.nodeName} title={user.displayName}>{user.displayName}</div>
      {user.jobTitle && (
        <div className={styles.nodeTitle} title={user.jobTitle} style={{ color: t.accentText }}>
          {user.jobTitle}
        </div>
      )}
      {showDepartment && user.department && !isGuest && !isDisabled && (
        <div className={styles.nodeDept}>{user.department}</div>
      )}
      {showOffice && user.officeLocation && !isGuest && !isDisabled && (
        <div className={styles.nodeOffice}>
          <Icon iconName="POI" className={styles.nodeOfficeIcon} />
          {user.officeLocation}
        </div>
      )}
      {!showDepartment && managerUser && !isGuest && !isDisabled && (
        <div className={styles.managerLine} style={{ color: t.managerLine }}>
          ↑ {managerUser.displayName.split(' ')[0]}
        </div>
      )}
      {(isGuest || isDisabled) && (
        <div
          className={styles.nodeDept}
          style={{ background: isDisabled ? '#fde7e9' : '#fff4ce', color: isDisabled ? '#c50f1f' : '#835c00' }}
        >
          {isDisabled ? strings.OrgNode_StatusDisabled : strings.OrgNode_StatusGuest}
        </div>
      )}
      {hasReports && (
        <button
          className={styles.expandBtn}
          onClick={e => { e.stopPropagation(); onToggle(node); }}
          title={expandLabel}
          aria-label={expandLabel}
          aria-expanded={isTree ? node.isExpanded : undefined}
          tabIndex={innerTabIndex}
        >
          {isExpanding
            ? <Icon iconName="ProgressRingDots" className={styles.spinning} />
            : (
              <>
                <Icon iconName={node.isExpanded ? 'ChevronUp' : 'ChevronDown'} />
                {directReportCount > 0 && <span className={styles.reportCount}>{directReportCount}</span>}
              </>
            )
          }
        </button>
      )}
      {hasReports && !isExpanding && totalReportCount > directReportCount && (
        <span
          className={styles.totalReportBadge}
          title={formatString(strings.OrgNode_TotalReportsAcrossLevels, { count: totalReportCount, name: user.displayName })}
          aria-label={formatString(strings.OrgNode_TotalReportsAcrossLevels, { count: totalReportCount, name: user.displayName })}
          role="note"
        >
          <Icon iconName="Group" />
          {totalReportCount}
        </span>
      )}
    </div>
  );
};

export const OrgNodeCard = React.memo(OrgNodeCardInner);

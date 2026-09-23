import * as React from 'react';
import { IGraphUser, IOrgNode, PresenceAvailability } from '../../../../services/GraphService';
import { OrgChartTheme } from './orgTheme';
import { matchesQuery } from './orgTreeUtils';
import { OrgNodeCard } from './OrgNodeCard';
import styles from './OrgChart.module.scss';

/* ── Recursive tree ──────────────────────── */

export interface IOrgTreeProps {
  node: IOrgNode;
  photos: { [id: string]: string | null };
  presenceMap: Map<string, PresenceAvailability>;
  showDepartment: boolean;
  showOffice: boolean;
  expandingNodes: Set<string>;
  /** Lower-cased, trimmed query ('' when not searching) */
  searchQuery: string;
  theme: OrgChartTheme;
  accentColor?: string;
  /** Ids of nodes to draw (pass the filter, or have a descendant that does) */
  visibleIds: Set<string>;
  parentUser?: IGraphUser;
  compactCards: boolean;
  /** 1-based depth for aria-level */
  depth: number;
  posInSet: number;
  setSize: number;
  /** Id of the card that currently holds the roving tab stop */
  tabStopId: string;
  onToggle: (node: IOrgNode) => void;
  onCardClick: (user: IGraphUser) => void;
  onFocus: (user: IGraphUser) => void;
  onItemFocus: (userId: string) => void;
}

const OrgTreeInner: React.FC<IOrgTreeProps> = (props) => {
  const {
    node, photos, presenceMap, showDepartment, showOffice,
    expandingNodes, searchQuery, theme, accentColor, visibleIds, parentUser, compactCards,
    depth, posInSet, setSize, tabStopId,
    onToggle, onCardClick, onFocus, onItemFocus,
  } = props;
  const id = node.user.id;
  if (!visibleIds.has(id)) return null;

  const visibleReports     = node.directReports.filter(c => visibleIds.has(c.user.id));
  const hasVisibleChildren = node.isExpanded && visibleReports.length > 0;

  return (
    <div className={`${styles.nodeWrapper} ${hasVisibleChildren ? styles.hasChildren : ''}`} role="none">
      <OrgNodeCard
        node={node}
        photo={photos[id]}
        presence={presenceMap.get(id)}
        showDepartment={showDepartment}
        showOffice={showOffice}
        isExpanding={expandingNodes.has(id)}
        isHighlighted={!!searchQuery && matchesQuery(node, searchQuery)}
        theme={theme}
        accentColor={accentColor}
        directReportCount={visibleReports.length}
        totalReportCount={node.totalReportCount}
        managerUser={parentUser}
        compactCards={compactCards}
        variant="tree"
        ariaLevel={depth}
        ariaPosInSet={posInSet}
        ariaSetSize={setSize}
        isTabStop={tabStopId === id}
        onToggle={onToggle}
        onCardClick={onCardClick}
        onFocus={onFocus}
        onItemFocus={onItemFocus}
      />
      {hasVisibleChildren && (
        <div className={styles.children} role="group">
          {visibleReports.map((child, i) => (
            <OrgTree
              {...props}
              key={child.user.id}
              node={child}
              parentUser={node.user}
              depth={depth + 1}
              posInSet={i + 1}
              setSize={visibleReports.length}
            />
          ))}
        </div>
      )}
    </div>
  );
};

export const OrgTree: React.NamedExoticComponent<IOrgTreeProps> = React.memo(OrgTreeInner);

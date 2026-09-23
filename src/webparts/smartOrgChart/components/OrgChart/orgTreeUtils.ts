// Pure helpers for the org chart's immutable IOrgNode tree. Nothing in here
// touches React state, the DOM or the Graph service.
import { IGraphUser, IOrgNode } from '../../../../services/GraphService';

export interface IFilterCounts { members: number; guests: number; disabled: number; }

export interface IOrgStats { total: number; members: number; guests: number; depts: number; }

export interface IVisibilityFilters {
  filterMembers: boolean;
  filterGuests: boolean;
  filterDisabled: boolean;
  filterDepartments: Set<string>;
}

/* ── Tree mutation helpers ───────────────── */

// Replaces the node with the given id. Only the path from the root to the
// target is copied — untouched subtrees keep their references, so memoized
// OrgTree branches skip re-rendering. Returns `root` itself when not found.
export function updateNode(root: IOrgNode, targetId: string, fn: (n: IOrgNode) => IOrgNode): IOrgNode {
  const seen = new Set<string>();
  const visit = (n: IOrgNode): IOrgNode => {
    if (n.user.id === targetId) return fn(n);
    if (seen.has(n.user.id)) return n;
    seen.add(n.user.id);
    let changed = false;
    const kids = n.directReports.map(c => {
      const next = visit(c);
      if (next !== c) changed = true;
      return next;
    });
    return changed ? { ...n, directReports: kids } : n;
  };
  return visit(root);
}

export function setNodeExpanded(root: IOrgNode, targetId: string, expanded: boolean): IOrgNode {
  return updateNode(root, targetId, n => ({ ...n, isExpanded: expanded }));
}

export function injectChildren(root: IOrgNode, targetId: string, children: IOrgNode[]): IOrgNode {
  return updateNode(root, targetId, n => ({ ...n, directReports: children, childrenLoaded: true, isExpanded: true }));
}

export function collapseAll(node: IOrgNode): IOrgNode {
  return { ...node, isExpanded: false, directReports: node.directReports.map(collapseAll) };
}

export function expandLoaded(node: IOrgNode): IOrgNode {
  return { ...node, isExpanded: node.directReports.length > 0, directReports: node.directReports.map(expandLoaded) };
}

export function markNodeLoaded(root: IOrgNode, targetId: string): IOrgNode {
  return updateNode(root, targetId, n => ({ ...n, childrenLoaded: true, directReports: [] }));
}

/** Marks many nodes as loaded-with-no-reports in a single traversal. */
export function markNodesLoaded(root: IOrgNode, targetIds: Set<string>): IOrgNode {
  if (targetIds.size === 0) return root;
  const visit = (n: IOrgNode): IOrgNode => {
    if (targetIds.has(n.user.id)) return { ...n, childrenLoaded: true, directReports: [] };
    let changed = false;
    const kids = n.directReports.map(c => {
      const next = visit(c);
      if (next !== c) changed = true;
      return next;
    });
    return changed ? { ...n, directReports: kids } : n;
  };
  return visit(root);
}

// Injects children for many nodes in a single traversal — one full-tree clone
// per call instead of one per node (injectChildren is O(tree) per call).
export function injectChildrenBatch(root: IOrgNode, childrenById: Map<string, IOrgNode[]>): IOrgNode {
  const injected = childrenById.get(root.user.id);
  if (injected !== undefined) {
    return { ...root, childrenLoaded: true, isExpanded: true, directReports: injected };
  }
  return { ...root, directReports: root.directReports.map(c => injectChildrenBatch(c, childrenById)) };
}

// Removes repeated users (manager cycles or self-managed accounts in the
// directory) so each person appears at most once and React keys stay unique.
export function dedupeTree(root: IOrgNode): IOrgNode {
  const seen = new Set<string>();
  const visit = (n: IOrgNode): IOrgNode => {
    seen.add(n.user.id);
    const kids: IOrgNode[] = [];
    n.directReports.forEach(c => { if (!seen.has(c.user.id)) kids.push(visit(c)); });
    return kids.length === n.directReports.length && kids.every((k, i) => k === n.directReports[i])
      ? n
      : { ...n, directReports: kids };
  };
  return visit(root);
}

/** Normalizes a freshly built tree from the service: dedupe, then expand loaded levels. */
export function prepareTree(raw: IOrgNode): IOrgNode {
  return expandLoaded(dedupeTree(raw));
}

// Expands every ancestor of a node matching the query so search hits inside
// collapsed branches become visible. Only already-loaded nodes are affected.
export function expandToMatches(node: IOrgNode, lowerQ: string): { node: IOrgNode; hasMatch: boolean } {
  const children = node.directReports.map(c => expandToMatches(c, lowerQ));
  const childMatch = children.some(c => c.hasMatch);
  return {
    node: { ...node, directReports: children.map(c => c.node), isExpanded: node.isExpanded || childMatch },
    hasMatch: childMatch || matchesQuery(node, lowerQ),
  };
}

/* ── Queries ─────────────────────────────── */

export function matchUserQuery(user: IGraphUser, lowerQ: string): boolean {
  if (!lowerQ) return false;
  return (
    (user.displayName || '').toLowerCase().indexOf(lowerQ) !== -1 ||
    (user.jobTitle || '').toLowerCase().indexOf(lowerQ) !== -1 ||
    (user.department || '').toLowerCase().indexOf(lowerQ) !== -1 ||
    (user.mail || '').toLowerCase().indexOf(lowerQ) !== -1
  );
}

export function matchesQuery(node: IOrgNode, lowerQ: string): boolean {
  return matchUserQuery(node.user, lowerQ);
}

/** Builds the user-visibility predicate for the active user-type and department filters. */
export function buildIsVisible(f: IVisibilityFilters): (user: IGraphUser) => boolean {
  return (user: IGraphUser) => {
    if (user.accountEnabled === false) {
      if (!f.filterDisabled) return false;
    } else if (user.userType === 'Guest') {
      if (!f.filterGuests) return false;
    } else if (!f.filterMembers) {
      return false;
    }
    if (f.filterDepartments.size > 0 && !f.filterDepartments.has(user.department || '')) return false;
    return true;
  };
}

// Ids of every node that should be drawn: the node passes the filter itself,
// or any descendant does (ancestors stay visible for context). Computed once
// per (tree, filters) change instead of re-walking subtrees at every level.
export function computeVisibleIds(root: IOrgNode, isVisible: (u: IGraphUser) => boolean): Set<string> {
  const ids = new Set<string>();
  const seen = new Set<string>();
  const visit = (n: IOrgNode): boolean => {
    if (seen.has(n.user.id)) return false;
    seen.add(n.user.id);
    let keep = isVisible(n.user);
    n.directReports.forEach(c => { if (visit(c)) keep = true; });
    if (keep) ids.add(n.user.id);
    return keep;
  };
  visit(root);
  return ids;
}

export interface IRenderedNode {
  node: IOrgNode;
  parentId: string | null;
  /** 1-based level for aria-level */
  depth: number;
  /** Visible children (only meaningful when the node is expanded) */
  visibleChildren: IOrgNode[];
}

// Nodes currently drawn on screen in document (pre-)order: visible and not
// inside a collapsed branch. Drives keyboard navigation and presence polling.
export function getRenderedNodes(root: IOrgNode, visibleIds: Set<string>): IRenderedNode[] {
  const out: IRenderedNode[] = [];
  const seen = new Set<string>();
  const visit = (n: IOrgNode, parentId: string | null, depth: number): void => {
    if (seen.has(n.user.id) || !visibleIds.has(n.user.id)) return;
    seen.add(n.user.id);
    const visibleChildren = n.directReports.filter(c => visibleIds.has(c.user.id));
    out.push({ node: n, parentId, depth, visibleChildren });
    if (n.isExpanded) visibleChildren.forEach(c => visit(c, n.user.id, depth + 1));
  };
  visit(root, null, 1);
  return out;
}

// Prunes users hidden by the active filters. Mirrors what OrgTree renders:
// an ancestor is kept whenever any of its descendants survives.
export function filterTreeForExport(node: IOrgNode, isVisible: (u: IGraphUser) => boolean): IOrgNode | null {
  const directReports = node.directReports
    .map(c => filterTreeForExport(c, isVisible))
    .filter((c): c is IOrgNode => c !== null);
  if (!isVisible(node.user) && directReports.length === 0) return null;
  return { ...node, directReports };
}

// Counts matches among the nodes the chart actually draws (visibleIds),
// so the "N matches in tree" label agrees with the highlighted cards.
export function countSearchMatches(root: IOrgNode, lowerQ: string, visibleIds: Set<string>): number {
  let count = 0;
  const seen = new Set<string>();
  const visit = (n: IOrgNode): void => {
    if (seen.has(n.user.id) || !visibleIds.has(n.user.id)) return;
    seen.add(n.user.id);
    if (matchesQuery(n, lowerQ)) count++;
    n.directReports.forEach(visit);
  };
  visit(root);
  return count;
}

// Every user id currently in the tree. Used to drop fetched reports that are
// already present (self-managed accounts or manager cycles in Azure AD) —
// without this, expanding re-injects an ancestor and recursion never ends.
export function collectTreeIds(node: IOrgNode, into: Set<string> = new Set<string>()): Set<string> {
  if (into.has(node.user.id)) return into;
  into.add(node.user.id);
  node.directReports.forEach(c => collectTreeIds(c, into));
  return into;
}

export function getUnloadedFrontier(node: IOrgNode): IOrgNode[] {
  const result: IOrgNode[] = [];
  const seen = new Set<string>();
  const visit = (n: IOrgNode): void => {
    if (seen.has(n.user.id)) return;
    seen.add(n.user.id);
    if (!n.childrenLoaded) result.push(n);
    else n.directReports.forEach(visit);
  };
  visit(node);
  return result;
}

export function countTreeUsers(node: IOrgNode): IFilterCounts {
  const c: IFilterCounts = { members: 0, guests: 0, disabled: 0 };
  const seen = new Set<string>();
  const visit = (n: IOrgNode): void => {
    if (seen.has(n.user.id)) return;
    seen.add(n.user.id);
    const u = n.user;
    if (u.accountEnabled === false) c.disabled++;
    else if (u.userType === 'Guest') c.guests++;
    else c.members++;
    n.directReports.forEach(visit);
  };
  visit(node);
  return c;
}

export function getUniqueDepts(node: IOrgNode): Map<string, number> {
  const map = new Map<string, number>();
  const seen = new Set<string>();
  const visit = (n: IOrgNode): void => {
    if (seen.has(n.user.id)) return;
    seen.add(n.user.id);
    const dept = n.user.department || '';
    if (dept) map.set(dept, (map.get(dept) || 0) + 1);
    n.directReports.forEach(visit);
  };
  visit(node);
  return map;
}

export function computeStats(users: IGraphUser[]): IOrgStats {
  let members = 0, guests = 0;
  const deptSet = new Set<string>();
  for (const u of users) {
    if (u.department) deptSet.add(u.department);
    if (u.userType === 'Guest') guests++;
    else members++;
  }
  return { total: users.length, members, guests, depts: deptSet.size };
}

/** Removes repeated users (by id) and, optionally, one excluded id. Keeps first occurrence. */
export function dedupeUsers(users: IGraphUser[], excludeId?: string): IGraphUser[] {
  const seen = new Set<string>();
  if (excludeId) seen.add(excludeId);
  return users.filter(u => {
    if (seen.has(u.id)) return false;
    seen.add(u.id);
    return true;
  });
}

import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { IGraphUser, IOrgNode, PresenceAvailability } from '../../../../services/GraphService';
import { exportOrgChartToPdf, exportOrgChartToCsv } from '../../../../services/PdfExportService';
import { IOrgChartProps, IOrgChartState } from './IOrgChartProps';
import { getInitials } from '../personUtils';
import styles from './OrgChart.module.scss';
import {
  IFilterCounts, IOrgStats, IRenderedNode,
  buildIsVisible, collapseAll, collectTreeIds, computeStats, computeVisibleIds,
  countSearchMatches, countTreeUsers, dedupeUsers, expandLoaded, expandToMatches,
  filterTreeForExport, getRenderedNodes, getUniqueDepts, getUnloadedFrontier,
  injectChildren, injectChildrenBatch, markNodesLoaded, matchUserQuery, prepareTree, setNodeExpanded,
} from './orgTreeUtils';
import {
  ChartLayout, loadChartState, readUrlFocus, saveChartState, syncUrlFocus,
} from './chartPersistence';
import { OrgChartTheme, THEME_CONTAINER_CLASS, getThemeContainerStyle, getThemeTokens } from './orgTheme';
import { PersonCard } from './PersonCard';
import { OrgTree } from './OrgTree';
import { DrillView } from './DrillView';
import { NoConfigForm, UserFilterKey } from './ChartControls';
import { OrgChartToolbar } from './OrgChartToolbar';
import { formatString } from '../localeUtils';
import * as strings from 'SmartOrgChartWebPartStrings';

export type { OrgChartTheme };

/* ── Main OrgChart component ─────────────── */

interface IOrgChartLocalState extends IOrgChartState {
  /** Debounced, lower-cased query that drives highlighting, counts and results */
  appliedQuery: string;
  /** Keyboard-highlighted row in the search results listbox (-1 = none) */
  searchActiveIndex: number;
  presenceMap: Map<string, PresenceAvailability>;
  zoomLevel: number;
  selectedUser: IGraphUser | null;
  showFilters: boolean;
  filterMembers: boolean;
  filterGuests: boolean;
  filterDisabled: boolean;
  // Focus / navigation (full-tree mode)
  focusedUser: IGraphUser | null;
  ancestorChain: IGraphUser[];
  allUsers: IGraphUser[];
  showSearchResults: boolean;
  personCardManagerChain: IGraphUser[];
  personCardDottedManager: IGraphUser | null;
  personCardDottedReports: IGraphUser[];
  // Layout
  chartLayout: ChartLayout;
  // Find Me feedback
  findMeError: string;
  // Tier 3
  filterDepartments: Set<string>;
  showDeptFilter: boolean;
  showStats: boolean;
  // Drill-down mode
  drillPath: IGraphUser[];
  drillReports: IGraphUser[];
  drillLoadingId: string | null;
  drillReportCounts: Map<string, number>;
  showLayoutPicker: boolean;
  rootPickerQuery: string;
  rootPickerResults: IGraphUser[];
  runtimeRootUser: IGraphUser | null;
  /** Root chosen in the setup form when no topLevelUser is configured */
  setupRootId: string;
  /** Re-rooting with data already in memory — shown inline, not as a full-screen spinner */
  isRefocusing: boolean;
  isExpandingAll: boolean;
  /** Tree card that holds the roving tab stop */
  treeFocusId: string | null;
}

const SEARCH_DEBOUNCE_MS = 200;
const PRESENCE_POLL_MS   = 60000;
const MAX_SEARCH_RESULTS = 8;

const EMPTY_IDS = new Set<string>();

const yieldToBrowser = (): Promise<void> => new Promise<void>(resolve => { window.setTimeout(resolve, 0); });

export class OrgChart extends React.Component<IOrgChartProps, IOrgChartLocalState> {
  private _mounted          = false;
  private _uid              = `soc${Math.random().toString(36).slice(2, 8)}`;
  private _pendingFocusEmail: string | null = null;
  // True until a deep-link / saved focus has been restored — persisting
  // before then would record the default position and strip ?socFocus
  private _restorePending   = false;
  // The ?socFocus value this instance wrote — only that value may be removed
  private _ownUrlFocus: string | null = null;
  private _presenceInterval: number | null = null;
  private _presenceTimer: number | null = null;
  private _searchTimer: number | null = null;
  private _scrollEl: HTMLDivElement | null = null;
  private _searchRef        = React.createRef<HTMLDivElement>();
  private _rootPickerRef    = React.createRef<HTMLDivElement>();
  private _drillHeaderRef   = React.createRef<HTMLDivElement>();
  private _personCardOpener: HTMLElement | null = null;
  private _isPanning        = false;
  private _requestedReportCounts = new Set<string>();
  private _requestedPhotos  = new Set<string>();
  private _panStartX        = 0;
  private _panStartY        = 0;
  private _scrollStartX     = 0;
  private _scrollStartY     = 0;
  private _panDistance      = 0;
  private _lastPanEndTime   = 0;
  // Incremented by every navigation that replaces the view (load, focus,
  // re-root) so late responses from a superseded request are dropped
  private _navSeq           = 0;
  // Incremented whenever the whole tree is replaced; long-running tree jobs
  // (Expand All) stop when it changes
  private _treeGen          = 0;
  private _deptFilterValidated = false;
  // Tree snapshot taken when a search starts, restored when it is cleared.
  // Tree edits made while searching are applied to it too (see _mutateTree).
  private _preSearchRoot: IOrgNode | null = null;
  // Per-render tree scans are cached by reference — the tree is immutable,
  // so a changed rootNode/allUsers/filter reference is the only invalidation signal
  private _treeScanFor: IOrgNode | null = null;
  private _treeCounts: IFilterCounts = { members: 0, guests: 0, disabled: 0 };
  private _uniqueDepts: Map<string, number> = new Map();
  private _statsFor: IGraphUser[] | null = null;
  private _stats: IOrgStats | null = null;
  private _visKey: unknown[] = [];
  private _visibleIds: Set<string> = EMPTY_IDS;
  private _renderedFor: [IOrgNode | null, Set<string> | null] = [null, null];
  private _rendered: IRenderedNode[] = [];
  private _renderedIds: Set<string> = EMPTY_IDS;
  private _matchKey: unknown[] = [];
  private _matchCount = 0;
  private _resultsKey: unknown[] = [];
  private _results: IGraphUser[] = [];

  constructor(props: IOrgChartProps) {
    super(props);
    const stored = loadChartState(props.instanceId);
    // A shared deep link takes precedence over the user's own saved position
    this._pendingFocusEmail = readUrlFocus() || stored.focusEmail || null;
    this._restorePending = !!this._pendingFocusEmail;
    // Stored preferences only apply while the admin has the matching control
    // enabled — otherwise users could be stuck in a state they can't change.
    this.state = {
      rootNode: null, isLoading: false, error: null,
      photos: {}, expandingNodes: new Set(), searchQuery: '',
      appliedQuery: '', searchActiveIndex: -1,
      presenceMap: new Map(), zoomLevel: props.defaultZoom > 0 ? props.defaultZoom : 1,
      selectedUser: null,
      showFilters: false,
      filterMembers: props.enableUserFilter ? (stored.filterMembers ?? true) : true,
      filterGuests: props.enableUserFilter ? (stored.filterGuests ?? true) : true,
      filterDisabled: props.enableUserFilter ? (stored.filterDisabled ?? true) : true,
      focusedUser: null, ancestorChain: [], allUsers: [],
      showSearchResults: false, personCardManagerChain: [],
      personCardDottedManager: null, personCardDottedReports: [],
      chartLayout: props.enableLayoutToggle
        ? (stored.chartLayout ?? (props.defaultLayout || 'drill'))
        : (props.defaultLayout || 'drill'),
      findMeError: '',
      filterDepartments: props.enableDeptFilter ? new Set(stored.filterDepartments ?? []) : new Set(),
      showDeptFilter: false,
      showStats: props.enableStats ? (stored.showStats ?? false) : false,
      drillPath: [], drillReports: [], drillLoadingId: null,
      drillReportCounts: new Map(),
      showLayoutPicker: false,
      rootPickerQuery: '', rootPickerResults: [], runtimeRootUser: null,
      setupRootId: '',
      isRefocusing: false,
      isExpandingAll: false,
      treeFocusId: null,
    };
  }

  public async componentDidMount(): Promise<void> {
    this._mounted = true;
    // Register listeners before any await so an unmount mid-load can't leak them
    document.addEventListener('mousedown', this._handleOutsideClick);
    document.addEventListener('keydown', this._handleEscKey);
    document.addEventListener('visibilitychange', this._handleVisibilityChange);
    this._presenceInterval = window.setInterval(() => { this._refreshPresence(); }, PRESENCE_POLL_MS);
    this._syncSearchAria();

    if (this.props.graphService && this._getRootIdentifier()) await this._loadTree();
  }

  public async componentDidUpdate(prev: IOrgChartProps, prevState: IOrgChartLocalState): Promise<void> {
    this._syncSearchAria();

    const rootChanged = prev.topLevelUser !== this.props.topLevelUser;
    if (
      rootChanged ||
      prev.levelsBelow  !== this.props.levelsBelow  ||
      (!prev.graphService && this.props.graphService)
    ) {
      this._requestedReportCounts.clear();
      this._preSearchRoot = null;
      this._clearSearchTimer();
      this._treeGen++;
      const runtimeRootUser = rootChanged ? null : this.state.runtimeRootUser;
      // State updates here are batched, so resolve the identifier up front
      const identifier = this.props.topLevelUser || (rootChanged ? '' : this.state.setupRootId);
      this.setState({
        rootNode: null, error: null, searchQuery: '', appliedQuery: '', showSearchResults: false,
        focusedUser: null, ancestorChain: [],
        drillPath: [], drillReports: [], drillLoadingId: null,
        drillReportCounts: new Map(),
        // A new admin-configured root supersedes any runtime/setup choice
        runtimeRootUser,
        setupRootId: rootChanged ? '' : this.state.setupRootId,
      });
      if (this.props.graphService) {
        if (runtimeRootUser) await this._loadTreeForUser(runtimeRootUser);
        else await this._loadTree(false, identifier);
      }
    }

    // Admin property pane edits re-render (not remount) this component, so
    // defaults and feature flags must be applied to live state here. Turning
    // a feature off also resets its persisted state — otherwise users could
    // be stuck in a filter or layout they no longer have a control for.
    if (prev.defaultLayout !== this.props.defaultLayout) {
      this._setLayout(this.props.defaultLayout || 'drill');
    }
    if (prev.defaultZoom !== this.props.defaultZoom) {
      this.setState({ zoomLevel: this.props.defaultZoom > 0 ? this.props.defaultZoom : 1 },
        () => { if (!(this.props.defaultZoom > 0)) this._autoFitZoom(); });
    }
    if (prev.enableLayoutToggle && !this.props.enableLayoutToggle) {
      this._setLayout(this.props.defaultLayout || 'drill');
    }
    if (prev.enableUserFilter && !this.props.enableUserFilter) {
      this.setState({ filterMembers: true, filterGuests: true, filterDisabled: true });
    }
    if (prev.enableDeptFilter && !this.props.enableDeptFilter) {
      this.setState({ filterDepartments: new Set() });
    }
    if (prev.enableStats && !this.props.enableStats) {
      this.setState({ showStats: false });
    }

    const s = this.state;
    if (
      prevState.rootNode          !== s.rootNode          ||
      prevState.filterDepartments !== s.filterDepartments ||
      prevState.filterMembers     !== s.filterMembers     ||
      prevState.filterGuests      !== s.filterGuests      ||
      prevState.filterDisabled    !== s.filterDisabled    ||
      prevState.zoomLevel         !== s.zoomLevel         ||
      prevState.chartLayout       !== s.chartLayout
    ) {
      this._fixConnectorLines();
    }

    // The set of on-screen users changed — fetch presence for the newcomers
    // (the service only requests ids it hasn't resolved within its TTL)
    if (
      prevState.rootNode     !== s.rootNode     ||
      prevState.drillReports !== s.drillReports ||
      prevState.drillPath    !== s.drillPath    ||
      prevState.chartLayout  !== s.chartLayout
    ) {
      this._schedulePresenceRefresh();
    }

    if (!s.isLoading && (
      prevState.chartLayout       !== s.chartLayout       ||
      prevState.showStats         !== s.showStats         ||
      prevState.filterMembers     !== s.filterMembers     ||
      prevState.filterGuests      !== s.filterGuests      ||
      prevState.filterDisabled    !== s.filterDisabled    ||
      prevState.filterDepartments !== s.filterDepartments ||
      prevState.drillPath         !== s.drillPath         ||
      prevState.focusedUser       !== s.focusedUser
    )) {
      this._persistState();
    }
  }

  public componentWillUnmount(): void {
    this._mounted = false;
    if (this._presenceInterval !== null) { window.clearInterval(this._presenceInterval); this._presenceInterval = null; }
    if (this._presenceTimer !== null) { window.clearTimeout(this._presenceTimer); this._presenceTimer = null; }
    this._clearSearchTimer();
    this._setScrollRef(null);
    document.removeEventListener('mousedown', this._handleOutsideClick);
    document.removeEventListener('keydown', this._handleEscKey);
    document.removeEventListener('visibilitychange', this._handleVisibilityChange);
  }

  /* ── Persistence ── */

  private _persistState(): void {
    if (this._restorePending) return;
    const { chartLayout, showStats, filterMembers, filterGuests, filterDisabled,
            filterDepartments, drillPath, focusedUser, rootNode } = this.state;
    // Only persist a focus once the user actually navigated away from the
    // default position — otherwise every page view stamps ?socFocus into
    // the URL (and two instances on one page fight over the same param)
    const atDefault = chartLayout === 'drill'
      ? drillPath.length === 0 ||
        (drillPath.length === 1 && (!rootNode || drillPath[0].id === rootNode.user.id))
      : !focusedUser;
    const focusEmail = atDefault
      ? null
      : chartLayout === 'drill'
        ? (drillPath[drillPath.length - 1].mail || null)
        : (focusedUser?.mail || null);
    saveChartState(this.props.instanceId, {
      chartLayout, showStats, filterMembers, filterGuests, filterDisabled,
      filterDepartments: Array.from(filterDepartments),
      focusEmail,
    });
    this._ownUrlFocus = syncUrlFocus(focusEmail, this._ownUrlFocus);
  }

  /* ── Document listeners ── */

  private _handleEscKey = (e: KeyboardEvent): void => {
    if (e.key !== 'Escape') return;
    const { showFilters, showDeptFilter, showLayoutPicker } = this.state;
    if (showFilters || showDeptFilter || showLayoutPicker) {
      this.setState({ showFilters: false, showDeptFilter: false, showLayoutPicker: false });
    }
  }

  private _handleOutsideClick = (e: MouseEvent): void => {
    if (this._searchRef.current && !this._searchRef.current.contains(e.target as Node)) {
      if (this.state.showSearchResults) this.setState({ showSearchResults: false, searchActiveIndex: -1 });
    }
    if (this._rootPickerRef.current && !this._rootPickerRef.current.contains(e.target as Node)) {
      if (this.state.rootPickerResults.length > 0) this.setState({ rootPickerResults: [] });
    }
  }

  private _handleVisibilityChange = (): void => {
    if (!document.hidden) this._refreshPresence();
  }

  /* ── Presence ── */

  // Users actually drawn on screen: the current drill level, or the expanded,
  // visible part of the tree (collapsed branches are skipped).
  private _collectPresenceIds(): string[] {
    const { chartLayout, drillPath, selectedUser, rootNode } = this.state;
    const ids: string[] = [];
    if (chartLayout === 'drill') {
      if (drillPath.length > 0) ids.push(drillPath[drillPath.length - 1].id);
      this._getVisibleDrillReports().forEach(u => ids.push(u.id));
    } else if (rootNode) {
      this._getRenderedNodes().forEach(r => ids.push(r.node.user.id));
    }
    if (selectedUser) ids.push(selectedUser.id);
    return ids;
  }

  private async _refreshPresence(): Promise<void> {
    const { graphService } = this.props;
    if (!graphService || !this._mounted) return;
    if (document.hidden) return; // resumes on visibilitychange
    const ids = this._collectPresenceIds();
    if (ids.length === 0) return;
    try {
      const result = await graphService.getPresence(ids);
      if (this._mounted) this.setState({ presenceMap: new Map(result) });
    } catch { /* presence is best-effort */ }
  }

  private _schedulePresenceRefresh(): void {
    if (this._presenceTimer !== null) window.clearTimeout(this._presenceTimer);
    this._presenceTimer = window.setTimeout(() => {
      this._presenceTimer = null;
      this._refreshPresence();
    }, 800);
  }

  /* ── Derived data (cached by reference) ── */

  private _getIsVisible(): (user: IGraphUser) => boolean {
    const { filterMembers, filterGuests, filterDisabled, filterDepartments } = this.state;
    return buildIsVisible({ filterMembers, filterGuests, filterDisabled, filterDepartments });
  }

  private _getVisibleIds(): Set<string> {
    const { rootNode, filterMembers, filterGuests, filterDisabled, filterDepartments } = this.state;
    if (!rootNode) return EMPTY_IDS;
    const key = [rootNode, filterMembers, filterGuests, filterDisabled, filterDepartments];
    if (!sameKey(key, this._visKey)) {
      this._visKey = key;
      this._visibleIds = computeVisibleIds(rootNode, this._getIsVisible());
    }
    return this._visibleIds;
  }

  private _getRenderedNodes(): IRenderedNode[] {
    const { rootNode } = this.state;
    if (!rootNode) return [];
    const visibleIds = this._getVisibleIds();
    if (this._renderedFor[0] !== rootNode || this._renderedFor[1] !== visibleIds) {
      this._renderedFor = [rootNode, visibleIds];
      this._rendered = getRenderedNodes(rootNode, visibleIds);
      const ids = new Set<string>();
      this._rendered.forEach(r => ids.add(r.node.user.id));
      this._renderedIds = ids;
    }
    return this._rendered;
  }

  private _getVisibleDrillReports(): IGraphUser[] {
    const isVisible = this._getIsVisible();
    return this.state.drillReports.filter(u => isVisible(u));
  }

  private _getMatchCount(lowerQ: string): number {
    const { rootNode } = this.state;
    if (!rootNode || !lowerQ) return 0;
    const visibleIds = this._getVisibleIds();
    const key = [rootNode, visibleIds, lowerQ];
    if (!sameKey(key, this._matchKey)) {
      this._matchKey = key;
      this._matchCount = countSearchMatches(rootNode, lowerQ, visibleIds);
    }
    return this._matchCount;
  }

  private _getSearchResults(lowerQ: string): IGraphUser[] {
    const { allUsers } = this.state;
    if (!lowerQ || allUsers.length === 0) return [];
    const key = [allUsers, lowerQ];
    if (!sameKey(key, this._resultsKey)) {
      this._resultsKey = key;
      this._results = allUsers.filter(u => matchUserQuery(u, lowerQ)).slice(0, MAX_SEARCH_RESULTS);
    }
    return this._results;
  }

  // Builds the tree that matches what's on screen: the current drill level in
  // drill mode, otherwise the loaded tree with filtered-out users pruned.
  private _getExportTree(): { node: IOrgNode; note?: string } | null {
    const { rootNode, chartLayout, drillPath } = this.state;
    const isVisible = this._getIsVisible();

    if (chartLayout === 'drill' && drillPath.length > 0) {
      const current = drillPath[drillPath.length - 1];
      return {
        node: {
          user: current,
          directReports: this._getVisibleDrillReports().map(u => ({
            user: u, directReports: [], isExpanded: false, childrenLoaded: true, level: 1,
            totalReportCount: this.props.graphService?.getTotalReportCount(u.id) ?? 0,
          })),
          isExpanded: true, childrenLoaded: true, level: 0,
          totalReportCount: this.props.graphService?.getTotalReportCount(current.id) ?? 0,
        },
        note: strings.Chart_ExportNote_DrillLevel,
      };
    }

    if (!rootNode) return null;
    const filtered = filterTreeForExport(rootNode, isVisible);
    if (!filtered) return null;
    const note = getUnloadedFrontier(rootNode).length > 0
      ? strings.Chart_ExportNote_PartialLevels
      : undefined;
    return { node: filtered, note };
  }

  public exportPdf(): void {
    const tree = this._getExportTree();
    if (tree) exportOrgChartToPdf(tree.node, tree.note, this.props.locale);
  }

  private _exportCsv = (): void => {
    const tree = this._getExportTree();
    if (tree) exportOrgChartToCsv(tree.node);
  }

  private _exportPdfClick = (): void => { this.exportPdf(); }

  /* ── Tree edits ── */

  // Applies an edit to the displayed tree and, while a search is active, to
  // the pre-search snapshot too — so clearing the search keeps the user's
  // expands/collapses and loaded children instead of discarding them.
  private _mutateTree(fn: (root: IOrgNode) => IOrgNode): void {
    if (this._preSearchRoot) this._preSearchRoot = fn(this._preSearchRoot);
    this.setState(prev => (prev.rootNode ? { rootNode: fn(prev.rootNode) } : null));
  }

  /* ── Search ── */

  private _clearSearchTimer(): void {
    if (this._searchTimer !== null) { window.clearTimeout(this._searchTimer); this._searchTimer = null; }
  }

  private _onSearchChange = (_e?: React.ChangeEvent<HTMLInputElement>, value?: string): void => {
    const v = value || '';
    const q = v.trim().toLowerCase();
    // The input updates immediately; the tree work is debounced
    this.setState({ searchQuery: v, showSearchResults: !!q, showFilters: false, searchActiveIndex: -1 });
    this._clearSearchTimer();
    if (!q) { this._applySearch(''); return; }
    this._searchTimer = window.setTimeout(() => {
      this._searchTimer = null;
      this._applySearch(q);
    }, SEARCH_DEBOUNCE_MS);
  }

  private _applySearch(q: string): void {
    if (!this._mounted) return;
    const { chartLayout, rootNode } = this.state;
    if (q && chartLayout !== 'drill' && rootNode) {
      // Snapshot the tree when a search starts so clearing it restores the
      // user's expand/collapse state instead of leaving everything expanded.
      // Always expand from the snapshot so refining the query doesn't
      // accumulate expansions from earlier keystrokes.
      if (!this._preSearchRoot) this._preSearchRoot = rootNode;
      this.setState({ appliedQuery: q, rootNode: expandToMatches(this._preSearchRoot, q).node });
    } else if (!q && this._preSearchRoot) {
      const restore = this._preSearchRoot;
      this._preSearchRoot = null;
      this.setState({ appliedQuery: '', rootNode: restore });
    } else {
      this.setState({ appliedQuery: q });
    }
  }

  private _onSearchFocus = (): void => {
    if (this.state.searchQuery.trim()) this.setState({ showSearchResults: true });
  }

  private _onSearchKeyDown = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key !== 'ArrowDown' && e.key !== 'ArrowUp') return;
    // Navigate the list that is on screen (driven by the debounced query)
    const results = this._getSearchResults(this.state.appliedQuery);
    if (results.length === 0) return;
    e.preventDefault();
    const cur = this.state.showSearchResults ? this.state.searchActiveIndex : -1;
    const next = e.key === 'ArrowDown'
      ? (cur + 1) % results.length
      : (cur <= 0 ? results.length - 1 : cur - 1);
    this.setState({ showSearchResults: true, searchActiveIndex: next });
  }

  private _onSearchEnter = (value?: string): void => {
    const q = (value || '').trim().toLowerCase();
    const { appliedQuery, searchActiveIndex, showSearchResults } = this.state;
    const results = this._getSearchResults(q);
    if (results.length === 0) return;
    // The highlighted row only applies when it belongs to the list on screen
    const idx = q === appliedQuery && showSearchResults && searchActiveIndex >= 0 && searchActiveIndex < results.length
      ? searchActiveIndex : 0;
    this._selectSearchResult(results[idx]);
  }

  private _onSearchEscape = (): void => {
    if (this.state.showSearchResults) this.setState({ showSearchResults: false, searchActiveIndex: -1 });
  }

  private _selectSearchResult = (user: IGraphUser): void => {
    this._handleFocusUser(user);
  }

  // Fluent's SearchBox hard-codes role="searchbox" on its input; the combobox
  // pattern needs role="combobox" + aria-expanded there. React leaves the
  // attribute alone after mount because the prop value never changes.
  private _syncSearchAria(): void {
    const wrap = this._searchRef.current;
    const input = wrap ? wrap.querySelector('input') : null;
    if (!input) return;
    const open = this._isSearchListOpen();
    if (input.getAttribute('role') !== 'combobox') input.setAttribute('role', 'combobox');
    input.setAttribute('aria-haspopup', 'listbox');
    input.setAttribute('aria-expanded', open ? 'true' : 'false');
  }

  private _isSearchListOpen(): boolean {
    return this.state.showSearchResults && this._getSearchResults(this.state.appliedQuery).length > 0;
  }

  /* ── Loading ── */

  private _getRootIdentifier(): string {
    return this.props.topLevelUser || this.state.setupRootId;
  }

  private _handleSetupLoad = (identifier: string): void => {
    this.setState({ setupRootId: identifier }, () => { this._loadTree(); });
  }

  // Loads the configured (or setup-form) root. `light` keeps the current UI
  // on screen with an inline indicator instead of the full-screen spinner.
  private async _loadTree(light = false, identifierOverride?: string): Promise<void> {
    const { graphService } = this.props;
    if (!graphService) return;
    const identifier = identifierOverride !== undefined ? identifierOverride : this._getRootIdentifier();
    const seq = ++this._navSeq;
    this._preSearchRoot = null;
    this._clearSearchTimer();
    if (!identifier) {
      // Nothing configured — show the setup form
      this.setState({ isLoading: false, isRefocusing: false, error: null, rootNode: null });
      return;
    }
    this.setState({
      isLoading: !light, isRefocusing: light, error: null,
      searchQuery: '', appliedQuery: '', showSearchResults: false,
    });
    try {
      const [rootUser, allUsers] = await Promise.all([
        graphService.findUser(identifier),
        graphService.getAllUsers().catch(() => [] as IGraphUser[]),
      ]);
      if (!this._mounted || seq !== this._navSeq) return;
      if (!rootUser) {
        const fromSetup = !this.props.topLevelUser;
        this._restorePending = false;
        this.setState({
          isLoading: false, isRefocusing: false,
          // Clearing the setup choice lets Retry fall back to the setup form
          setupRootId: '',
          error: fromSetup
            ? formatString(strings.Chart_NotFoundFromSetup, { identifier })
            : formatString(strings.Chart_NotFoundWithSettings, { identifier }),
        });
        return;
      }
      await this._applyRootUser(rootUser, allUsers, seq);
    } catch (err) {
      const detail = err instanceof Error && err.message ? ` ${err.message}` : strings.Chart_LoadFailedPermissions;
      this._restorePending = false;
      if (this._mounted && seq === this._navSeq) {
        this.setState({ isLoading: false, isRefocusing: false, error: formatString(strings.Chart_LoadFailedGeneric, { detail }) });
      }
    }
  }

  // Loads the tree for a specific user (root picker / runtime root).
  private async _loadTreeForUser(user: IGraphUser, light = false): Promise<void> {
    const seq = ++this._navSeq;
    this._preSearchRoot = null;
    this._clearSearchTimer();
    this.setState({
      isLoading: !light, isRefocusing: light, error: null,
      searchQuery: '', appliedQuery: '', showSearchResults: false,
    });
    try {
      await this._applyRootUser(user, null, seq);
    } catch {
      this._restorePending = false;
      if (this._mounted && seq === this._navSeq) {
        this.setState({ isLoading: false, isRefocusing: false, error: strings.Chart_LoadFailedForPerson });
      }
    }
  }

  private async _applyRootUser(rootUser: IGraphUser, allUsers: IGraphUser[] | null, seq: number): Promise<void> {
    const gs = this.props.graphService;
    if (!gs) return;
    const rawRoot = await gs.buildOrgTree(rootUser.id, this.props.levelsBelow);
    if (!this._mounted || seq !== this._navSeq) return;

    const rootNode = prepareTree(rawRoot);
    const drillReports = rootNode.directReports.map(n => n.user);
    const updates: Partial<IOrgChartLocalState> = {
      rootNode, isLoading: false, isRefocusing: false, error: null,
      drillPath: [rootUser], drillReports, drillLoadingId: null,
      focusedUser: null, ancestorChain: [], treeFocusId: null,
    };
    if (allUsers) updates.allUsers = allUsers;

    // A saved department filter that matches nobody would leave an empty chart
    if (!this._deptFilterValidated) {
      this._deptFilterValidated = true;
      const { filterDepartments } = this.state;
      if (filterDepartments.size > 0) {
        const present = getUniqueDepts(rootNode);
        const kept = new Set<string>();
        filterDepartments.forEach(d => { if (present.has(d)) kept.add(d); });
        if (kept.size !== filterDepartments.size) updates.filterDepartments = kept;
      }
    }

    this._treeGen++;
    this.setState(updates as IOrgChartLocalState, () => {
      this._autoFitZoom();
      this._refreshPresence();
    });
    this._loadPhotosForTree(rootNode);
    this._checkFrontier(getUnloadedFrontier(rootNode));
    this._loadDrillReportCounts(drillReports);

    // Restore the deep-linked / last saved navigation position
    if (this._pendingFocusEmail) {
      const emailToRestore = this._pendingFocusEmail.toLowerCase();
      this._pendingFocusEmail = null;
      if (emailToRestore !== (rootUser.mail || '').toLowerCase()) {
        const pool = allUsers || this.state.allUsers;
        const foundUser = pool.find(u => (u.mail || '').toLowerCase() === emailToRestore)
          || await gs.findUser(emailToRestore).catch(() => null);
        if (foundUser && this._mounted && seq === this._navSeq) {
          await this._handleFocusUser(foundUser, false);
        }
      }
    }
    if (this._restorePending) {
      this._restorePending = false;
      if (this._mounted) this._persistState();
    }
  }

  private _fixConnectorLines(): void {
    requestAnimationFrame(() => {
      const container = this._scrollEl;
      if (!container || !this._mounted) return;
      const zoom = this.state.zoomLevel || 1;
      const allChildren = container.querySelectorAll(`.${styles.children}`) as NodeListOf<HTMLElement>;
      allChildren.forEach(childrenDiv => {
        const first = childrenDiv.firstElementChild as HTMLElement;
        const last  = childrenDiv.lastElementChild  as HTMLElement;
        if (!first || !last) return;
        const left  = first.getBoundingClientRect().width / (2 * zoom);
        const right = last.getBoundingClientRect().width  / (2 * zoom);
        childrenDiv.style.setProperty('--conn-left',  `${left}px`);
        childrenDiv.style.setProperty('--conn-right', `${right}px`);
      });
    });
  }

  private _autoFitZoom(): void {
    if (this.props.defaultZoom > 0) return; // fixed zoom configured — don't override
    requestAnimationFrame(() => {
      const container = this._scrollEl;
      if (!container || !this._mounted) return;
      const treeWrapper = container.firstElementChild as HTMLElement;
      if (!treeWrapper) return;
      const naturalW = treeWrapper.offsetWidth;
      const naturalH = treeWrapper.offsetHeight;
      if (!naturalW || !naturalH) return;
      const cW = container.clientWidth;
      const cH = container.clientHeight;
      const scaleX = cW / naturalW;
      const scaleY = cH / naturalH;
      const fitZoom = Math.min(scaleX, scaleY) * 0.90;
      // Shrink to fit, or grow back toward 100% when a smaller subtree is shown
      const newZoom = Math.max(0.25, Math.min(1, fitZoom));
      if (Math.abs(newZoom - this.state.zoomLevel) > 0.01) {
        this.setState({ zoomLevel: newZoom });
      }
    });
  }

  private _loadPhotosForTree(node: IOrgNode): void {
    this._loadPhotos(Array.from(collectTreeIds(node)));
  }

  // Photos are fetched in chunks and applied in as few state updates as
  // possible, yielding to the browser between flushes so large trees don't
  // lock up the page.
  private async _loadPhotos(ids: string[]): Promise<void> {
    const { graphService } = this.props;
    if (!graphService) return;
    const todo = ids.filter(id => !(id in this.state.photos) && !this._requestedPhotos.has(id));
    if (todo.length === 0) return;
    todo.forEach(id => this._requestedPhotos.add(id));

    const CHUNK = 50;
    const FLUSH_AT = 200;
    let batch: { [id: string]: string | null } = {};
    let pending = 0;
    const flush = (): void => {
      const toApply = batch;
      batch = {};
      pending = 0;
      if (this._mounted && Object.keys(toApply).length > 0) {
        this.setState(prev => ({ photos: { ...prev.photos, ...toApply } }));
      }
    };
    for (let i = 0; i < todo.length; i += CHUNK) {
      if (!this._mounted) return;
      const slice = todo.slice(i, i + CHUNK);
      const results = await Promise.all(slice.map(id => graphService.getUserPhoto(id).catch(() => null)));
      slice.forEach((id, j) => { batch[id] = results[j]; });
      pending += slice.length;
      if (pending >= FLUSH_AT) { flush(); await yieldToBrowser(); }
    }
    flush();
  }

  // Marks frontier nodes that turn out to have no reports, so their expand
  // button disappears. One tree update per batch.
  private async _checkFrontier(frontier: IOrgNode[]): Promise<void> {
    const { graphService } = this.props;
    if (!graphService || frontier.length === 0) return;
    const batchSize = 10;
    for (let i = 0; i < frontier.length; i += batchSize) {
      if (!this._mounted) return;
      const slice = frontier.slice(i, i + batchSize);
      const results = await Promise.all(slice.map(n =>
        graphService.hasDirectReports(n.user.id).catch(() => true)));
      if (!this._mounted) return;
      const empty = new Set<string>();
      slice.forEach((n, j) => { if (!results[j]) empty.add(n.user.id); });
      if (empty.size > 0) this._mutateTree(r => markNodesLoaded(r, empty));
    }
  }

  /* ── Toggle expand (full-tree mode) ── */

  private _handleToggle = async (node: IOrgNode): Promise<void> => {
    if (!this.state.rootNode) return;
    const id = node.user.id;
    if (node.isExpanded) { this._mutateTree(r => setNodeExpanded(r, id, false)); return; }
    const graphService = this.props.graphService;
    if (!node.childrenLoaded && graphService) {
      if (this.state.expandingNodes.has(id)) return;
      this.setState(prev => { const s = new Set(prev.expandingNodes); s.add(id); return { expandingNodes: s }; });
      const clearExpanding = (): void => {
        this.setState(prev => { const s = new Set(prev.expandingNodes); s.delete(id); return { expandingNodes: s }; });
      };
      try {
        const reports = await graphService.getDirectReports(id);
        if (!this._mounted) return;
        if (this.state.rootNode) {
          // Drop reports already in the tree (or repeated in the response) —
          // a self-managed account or manager cycle would otherwise
          // re-inject an ancestor forever
          const treeIds  = collectTreeIds(this.state.rootNode);
          const children: IOrgNode[] = [];
          reports.forEach(u => {
            if (treeIds.has(u.id)) return;
            treeIds.add(u.id);
            children.push({
              user: u, directReports: [], isExpanded: false, childrenLoaded: false, level: node.level + 1,
              totalReportCount: graphService.getTotalReportCount(u.id),
            });
          });
          this._mutateTree(r => injectChildren(r, id, children));
          this._loadPhotos(children.map(c => c.user.id));
          this._checkFrontier(children);
        }
        clearExpanding();
      } catch {
        if (this._mounted) clearExpanding();
      }
    } else {
      this._mutateTree(r => setNodeExpanded(r, id, true));
    }
  }

  private _handleCollapseAll = (): void => {
    this._mutateTree(collapseAll);
  }

  // Expand All: expand already-loaded nodes for immediate feedback, then
  // BFS-load every unloaded frontier level. Each level is assembled off-state
  // and applied with a single tree update; the loop yields between batches.
  private _handleExpandAll = async (): Promise<void> => {
    const start = this.state.rootNode;
    if (!start || this.state.isExpandingAll) return;
    const gen = this._treeGen;
    const root0 = expandLoaded(start);
    this._mutateTree(expandLoaded);

    const gs = this.props.graphService;
    if (!gs) return;

    // treeIds guards against self-managed accounts and manager cycles;
    // the level cap is a safety valve so a guard regression can't hang the browser
    const treeIds = collectTreeIds(root0);
    let frontier = getUnloadedFrontier(root0);
    if (frontier.length === 0) { this._autoFitZoom(); return; }
    this.setState({ isExpandingAll: true });
    const MAX_LEVELS = 50;
    const BATCH = 25;
    try {
      for (let level = 0; frontier.length > 0 && level < MAX_LEVELS; level++) {
        const levelChildren = new Map<string, IOrgNode[]>();
        const next: IOrgNode[] = [];
        const newIds: string[] = [];
        for (let i = 0; i < frontier.length; i += BATCH) {
          if (!this._mounted || gen !== this._treeGen) return;
          const slice = frontier.slice(i, i + BATCH);
          const results = await Promise.all(
            slice.map(n => gs.getDirectReports(n.user.id)
              .then(r  => ({ node: n, reports: r }))
              .catch(() => ({ node: n, reports: [] as IGraphUser[] }))
            )
          );
          for (const { node: n, reports } of results) {
            const kids: IOrgNode[] = [];
            reports.forEach(u => {
              if (treeIds.has(u.id)) return;
              treeIds.add(u.id);
              kids.push({
                user: u, directReports: [], isExpanded: false, childrenLoaded: false, level: n.level + 1,
                totalReportCount: gs.getTotalReportCount(u.id),
              });
              newIds.push(u.id);
            });
            levelChildren.set(n.user.id, kids);
            next.push(...kids);
          }
          await yieldToBrowser();
        }
        if (!this._mounted || gen !== this._treeGen) return;
        this._mutateTree(r => expandLoaded(injectChildrenBatch(r, levelChildren)));
        if (newIds.length) this._loadPhotos(newIds);
        frontier = next;
      }
    } finally {
      if (this._mounted) this.setState({ isExpandingAll: false });
    }
    if (this._mounted) this._autoFitZoom();
  }

  /* ── Keyboard navigation (tree layouts) ── */

  private _handleTreeKeyDown = (e: React.KeyboardEvent<HTMLDivElement>): void => {
    const target = e.target as HTMLElement;
    if (target.getAttribute('role') !== 'treeitem') return;
    const id = target.getAttribute('data-soc-id');
    if (!id) return;
    const list = this._getRenderedNodes();
    let idx = -1;
    for (let i = 0; i < list.length; i++) { if (list[i].node.user.id === id) { idx = i; break; } }
    if (idx < 0) return;
    const cur = list[idx];
    const hasReports = cur.node.directReports.length > 0 || !cur.node.childrenLoaded;
    let targetId: string | null = null;

    switch (e.key) {
      case 'ArrowDown': if (idx < list.length - 1) targetId = list[idx + 1].node.user.id; break;
      case 'ArrowUp':   if (idx > 0) targetId = list[idx - 1].node.user.id; break;
      case 'Home':      targetId = list[0].node.user.id; break;
      case 'End':       targetId = list[list.length - 1].node.user.id; break;
      case 'ArrowRight':
        if (!cur.node.isExpanded && hasReports) this._handleToggle(cur.node);
        else if (cur.node.isExpanded && cur.visibleChildren.length > 0) targetId = cur.visibleChildren[0].user.id;
        break;
      case 'ArrowLeft':
        if (cur.node.isExpanded && hasReports) this._handleToggle(cur.node);
        else targetId = cur.parentId;
        break;
      default:
        return;
    }
    e.preventDefault();
    if (targetId) this._moveTreeFocus(targetId);
  }

  private _handleTreeItemFocus = (userId: string): void => {
    if (this.state.treeFocusId !== userId) this.setState({ treeFocusId: userId });
  }

  private _moveTreeFocus(userId: string): void {
    this.setState({ treeFocusId: userId }, () => this._focusTreeCard(userId));
  }

  private _focusTreeCard(userId: string): void {
    const container = this._scrollEl;
    if (!container) return;
    const cards = container.querySelectorAll('[data-soc-id]');
    for (let i = 0; i < cards.length; i++) {
      if (cards[i].getAttribute('data-soc-id') === userId) { (cards[i] as HTMLElement).focus(); return; }
    }
  }

  // After a drill/refocus the activated card may be gone — land on the header
  private _focusDrillHeaderIfLost(force = false): void {
    const active = document.activeElement;
    if (!force && active && active !== document.body && document.body.contains(active)) return;
    if (this._drillHeaderRef.current) this._drillHeaderRef.current.focus();
  }

  /* ── Card click → profile popup ── */

  private _handleCardClick = (user: IGraphUser): void => {
    if (Date.now() - this._lastPanEndTime < 150) return;
    const active = document.activeElement;
    this._personCardOpener = active instanceof HTMLElement && active !== document.body ? active : null;
    this.setState({
      selectedUser: user, showFilters: false, personCardManagerChain: [],
      personCardDottedManager: null, personCardDottedReports: [],
    });
    const gs = this.props.graphService;
    if (!gs) return;
    gs.getManagerChain(user.id, 8).then(chain => {
      if (this._mounted && this.state.selectedUser === user) this.setState({ personCardManagerChain: dedupeUsers(chain, user.id) });
    }).catch(() => { /* ignore */ });
    gs.getDottedLineReports(user.id).then(reports => {
      if (this._mounted && this.state.selectedUser === user && reports.length > 0) {
        this.setState({ personCardDottedReports: dedupeUsers(reports, user.id) });
      }
    }).catch(() => { /* ignore */ });
    if (user.dottedManagerId) {
      gs.findUser(user.dottedManagerId).then(mgr => {
        if (this._mounted && this.state.selectedUser === user && mgr) this.setState({ personCardDottedManager: mgr });
      }).catch(() => { /* ignore */ });
    }
  }

  private _closePersonCard = (): void => {
    const opener = this._personCardOpener;
    this._personCardOpener = null;
    this.setState({
      selectedUser: null, personCardManagerChain: [],
      personCardDottedManager: null, personCardDottedReports: [],
    }, () => {
      // Return focus to whatever opened the card, if it's still on the page
      if (opener && document.body.contains(opener)) opener.focus();
    });
  }

  /* ── Drill-down handlers ── */

  private _handleDrillInto = async (user: IGraphUser): Promise<void> => {
    const { graphService } = this.props;
    if (!graphService) return;
    // A manager cycle could lead back to someone already in the path
    const existing = this.state.drillPath.map(u => u.id).indexOf(user.id);
    if (existing !== -1) { this._handleDrillNavigate(existing); return; }
    const seq = this._navSeq;
    this.setState({ drillLoadingId: user.id });
    try {
      const reports = dedupeUsers(await graphService.getDirectReports(user.id), user.id);
      if (!this._mounted || seq !== this._navSeq) return;
      if (reports.length === 0) {
        // No reports — open profile popup instead of drilling into a dead end
        this.setState({ drillLoadingId: null });
        this._handleCardClick(user);
        return;
      }
      this.setState(prev => ({
        drillPath: [...prev.drillPath, user],
        drillReports: reports,
        drillLoadingId: null,
      }), () => this._focusDrillHeaderIfLost());
      this._loadPhotos([user.id, ...reports.map(u => u.id)]);
      this._loadDrillReportCounts(reports);
    } catch {
      if (this._mounted) this.setState({ drillLoadingId: null });
    }
  }

  private _handleDrillToggle = (node: IOrgNode): void => {
    this._handleDrillInto(node.user);
  }

  private _handleDrillNavigate = async (index: number): Promise<void> => {
    const { graphService } = this.props;
    const { drillPath } = this.state;
    if (!graphService || index >= drillPath.length) return;
    if (index === drillPath.length - 1) return; // already at this level
    const targetUser = drillPath[index];
    const seq = this._navSeq;
    this.setState({ drillLoadingId: targetUser.id });
    try {
      const reports = dedupeUsers(await graphService.getDirectReports(targetUser.id), targetUser.id);
      if (!this._mounted || seq !== this._navSeq) return;
      this.setState({
        drillPath: drillPath.slice(0, index + 1),
        drillReports: reports,
        drillLoadingId: null,
      }, () => this._focusDrillHeaderIfLost());
      this._loadPhotos(reports.map(u => u.id));
      this._loadDrillReportCounts(reports);
    } catch {
      if (this._mounted) this.setState({ drillLoadingId: null });
    }
  }

  /* ── Focus on person (full-tree mode or drill mode) ── */

  // Data is already in memory, so the current view stays up with an inline
  // indicator. `moveFocus` lands keyboard focus on the new root afterward
  // (skipped for the automatic deep-link restore on page load).
  private _handleFocusUser = async (user: IGraphUser, moveFocus = true): Promise<void> => {
    const { graphService, levelsBelow, levelsAbove } = this.props;
    if (!graphService) return;
    const { chartLayout } = this.state;
    const seq = ++this._navSeq;
    this._clearSearchTimer();
    if (moveFocus && this._restorePending) {
      // A user navigation supersedes a still-pending deep-link restore
      this._pendingFocusEmail = null;
      this._restorePending = false;
    }

    this.setState({
      isRefocusing: true, error: null,
      searchQuery: '', appliedQuery: '', showSearchResults: false, searchActiveIndex: -1,
    });

    if (chartLayout === 'drill') {
      try {
        const [reports, managerChain] = await Promise.all([
          graphService.getDirectReports(user.id),
          graphService.getManagerChain(user.id, levelsAbove),
        ]);
        if (!this._mounted || seq !== this._navSeq) return;
        const chain = dedupeUsers(managerChain, user.id);
        const drillReports = dedupeUsers(reports, user.id);
        this.setState({
          drillPath: [...chain, user],
          drillReports,
          isRefocusing: false,
          drillLoadingId: null,
        }, () => { if (moveFocus) this._focusDrillHeaderIfLost(true); });
        this._loadPhotos([user.id, ...chain.map(u2 => u2.id), ...drillReports.map(u2 => u2.id)]);
        this._loadDrillReportCounts(drillReports);
      } catch {
        if (this._mounted && seq === this._navSeq) this.setState({ isRefocusing: false });
      }
      return;
    }

    // Full-tree mode
    this._preSearchRoot = null;
    try {
      const [rawRoot, ancestorChain] = await Promise.all([
        graphService.buildOrgTree(user.id, levelsBelow),
        graphService.getManagerChain(user.id, levelsAbove),
      ]);
      if (!this._mounted || seq !== this._navSeq) return;
      const expanded = prepareTree(rawRoot);
      this._preSearchRoot = null;
      this._treeGen++;
      this.setState({
        rootNode: expanded, focusedUser: user,
        ancestorChain: dedupeUsers(ancestorChain, user.id),
        isRefocusing: false, treeFocusId: expanded.user.id,
      }, () => {
        if (this._scrollEl) {
          this._scrollEl.scrollLeft = 0;
          this._scrollEl.scrollTop  = 0;
        }
        this._autoFitZoom();
        if (moveFocus) this._focusTreeCard(expanded.user.id);
      });
      this._loadPhotosForTree(expanded);
      this._checkFrontier(getUnloadedFrontier(expanded));
    } catch {
      if (this._mounted && seq === this._navSeq) {
        this.setState({ isRefocusing: false, error: strings.Chart_LoadFailedForPerson });
      }
    }
  }

  private _handleFocusFromCard = (user: IGraphUser): void => {
    this._handleFocusUser(user);
  }

  // Reloads the current root: the runtime root when one is picked, otherwise
  // the configured / setup-form root.
  private _reloadRoot(light: boolean): void {
    const { runtimeRootUser } = this.state;
    if (runtimeRootUser) this._loadTreeForUser(runtimeRootUser, light);
    else this._loadTree(light);
  }

  private _handleRetry = (): void => { this._reloadRoot(false); }

  private _handleReturnToRoot = (): void => {
    const { chartLayout, rootNode } = this.state;
    if (chartLayout === 'drill' && rootNode) {
      const rootReports = rootNode.directReports.map(n => n.user);
      this._navSeq++;
      this.setState({
        drillPath: [rootNode.user],
        drillReports: rootReports,
        focusedUser: null,
        ancestorChain: [],
        drillLoadingId: null,
      }, () => this._focusDrillHeaderIfLost());
      this._loadDrillReportCounts(rootReports);
      return;
    }
    this._reloadRoot(true);
  }

  /* ── Background-fetch direct report counts for drill cards ── */

  private _loadDrillReportCounts(users: IGraphUser[]): void {
    const { graphService } = this.props;
    if (!graphService) return;
    const todo = users.filter(u => !this._requestedReportCounts.has(u.id));
    if (todo.length === 0) return;
    todo.forEach(u => this._requestedReportCounts.add(u.id));
    // One state update for the whole level instead of one per card
    Promise.all(todo.map(u => graphService.getDirectReports(u.id)
      .then(r => ({ id: u.id, count: r.filter(x => x.id !== u.id).length }))
      .catch(() => { this._requestedReportCounts.delete(u.id); return null; })
    )).then(results => {
      if (!this._mounted) return;
      this.setState(prev => {
        const next = new Map(prev.drillReportCounts);
        results.forEach(r => { if (r) next.set(r.id, r.count); });
        return { drillReportCounts: next };
      });
    }).catch(() => { /* ignore */ });
  }

  /* ── Find Me ── */

  private _handleFindMe = async (): Promise<void> => {
    const { graphService, currentUserEmail } = this.props;
    if (!graphService || !currentUserEmail) return;
    const user = await graphService.findUser(currentUserEmail).catch(() => null);
    if (user) {
      await this._handleFocusUser(user);
    } else if (this._mounted) {
      this.setState({ findMeError: strings.Chart_FindMeNotFound });
      setTimeout(() => { if (this._mounted) this.setState({ findMeError: '' }); }, 3000);
    }
  }

  /* ── Root picker ── */

  private _onRootPickerChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const query = e.target.value;
    const { allUsers } = this.state;
    if (!query.trim()) {
      this.setState({ rootPickerQuery: query, rootPickerResults: [] });
      return;
    }
    const q = query.toLowerCase();
    const results = allUsers.filter(u =>
      (u.displayName || '').toLowerCase().indexOf(q) !== -1 ||
      (u.mail || '').toLowerCase().indexOf(q) !== -1
    ).slice(0, 8);
    this.setState({ rootPickerQuery: query, rootPickerResults: results });
  }

  private _onRootPickerSelect = (user: IGraphUser): void => {
    this.setState({ rootPickerQuery: '', rootPickerResults: [], runtimeRootUser: user });
    this._requestedReportCounts.clear();
    this._loadTreeForUser(user);
  }

  private _resetRoot = (): void => {
    this.setState({ runtimeRootUser: null, rootPickerQuery: '', rootPickerResults: [] },
      () => { this._loadTree(); });
  }

  /* ── Layout picker ── */

  private _setLayout = (layout: ChartLayout): void => {
    const { rootNode, drillPath } = this.state;
    const updates: Partial<IOrgChartLocalState> = { chartLayout: layout };
    if (layout === 'drill' && drillPath.length === 0 && rootNode) {
      const rootReports = rootNode.directReports.map(n => n.user);
      updates.drillPath = [rootNode.user];
      updates.drillReports = rootReports;
      this._loadDrillReportCounts(rootReports);
    }
    this.setState(updates as IOrgChartLocalState);
  }

  private _toggleLayoutPicker = (): void => {
    this.setState(p => ({ showLayoutPicker: !p.showLayoutPicker, showFilters: false, showDeptFilter: false }));
  }

  private _toggleStats = (): void => { this.setState(p => ({ showStats: !p.showStats })); }

  private _toggleDeptFilter = (): void => {
    this.setState(p => ({ showDeptFilter: !p.showDeptFilter, showFilters: false, showLayoutPicker: false }));
  }

  private _toggleUserFilterPanel = (): void => {
    this.setState(p => ({ showFilters: !p.showFilters, showDeptFilter: false, showLayoutPicker: false }));
  }

  private _toggleUserFilter = (key: UserFilterKey): void => {
    this.setState(p => ({
      filterMembers:  key === 'members'  ? !p.filterMembers  : p.filterMembers,
      filterGuests:   key === 'guests'   ? !p.filterGuests   : p.filterGuests,
      filterDisabled: key === 'disabled' ? !p.filterDisabled : p.filterDisabled,
    }));
  }

  private _clearDeptFilter = (): void => { this.setState({ filterDepartments: new Set() }); }

  private _toggleDept = (dept: string): void => {
    this.setState(p => {
      const next = new Set(p.filterDepartments);
      if (next.has(dept)) next.delete(dept); else next.add(dept);
      return { filterDepartments: next };
    });
  }

  private _selectLayout = (layout: ChartLayout): void => {
    this._setLayout(layout);
    this.setState({ showLayoutPicker: false });
  }

  private _zoomOut = (): void => {
    this.setState(p => ({ zoomLevel: Math.max(0.25, p.zoomLevel - 0.1) }));
  }

  private _zoomIn = (): void => {
    this.setState(p => {
      const next = Math.min(1.5, p.zoomLevel + 0.1);
      return { zoomLevel: p.zoomLevel < 1 && next > 1 ? 1 : next };
    });
  }

  private _zoomReset = (): void => {
    this.setState({ zoomLevel: this.props.defaultZoom > 0 ? this.props.defaultZoom : 1 });
  }

  private _closePopups = (): void => {
    this.setState({ showFilters: false, showDeptFilter: false, showLayoutPicker: false });
  }

  /* ── Drag-to-pan (mouse) ── */

  // Callback ref: the scroll area unmounts in drill layout and while loading,
  // so the non-passive touchmove listener follows the element's lifetime.
  private _setScrollRef = (el: HTMLDivElement | null): void => {
    if (this._scrollEl === el) return;
    if (this._scrollEl) this._scrollEl.removeEventListener('touchmove', this._handleTouchMoveDirect);
    this._scrollEl = el;
    if (el) el.addEventListener('touchmove', this._handleTouchMoveDirect, { passive: false });
  }

  // The grabbing cursor is toggled directly on the element — a state update
  // would re-render the whole chart at the start and end of every drag
  private _setPanningClass(on: boolean): void {
    if (!this._scrollEl) return;
    if (on) this._scrollEl.classList.add(styles.treeScrollPanning);
    else this._scrollEl.classList.remove(styles.treeScrollPanning);
  }

  private _handlePanStart = (e: React.MouseEvent<HTMLDivElement>): void => {
    if ((e.target as HTMLElement).closest('button, a, input')) return;
    this._isPanning    = true;
    this._panDistance  = 0;
    this._panStartX    = e.clientX;
    this._panStartY    = e.clientY;
    this._scrollStartX = this._scrollEl?.scrollLeft ?? 0;
    this._scrollStartY = this._scrollEl?.scrollTop  ?? 0;
    this._setPanningClass(true);
    e.preventDefault();
  }

  private _handlePanMove = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (!this._isPanning || !this._scrollEl) return;
    const dx = e.clientX - this._panStartX;
    const dy = e.clientY - this._panStartY;
    this._panDistance = Math.sqrt(dx * dx + dy * dy);
    this._scrollEl.scrollLeft = this._scrollStartX - dx;
    this._scrollEl.scrollTop  = this._scrollStartY - dy;
  }

  private _handlePanEnd = (): void => {
    if (!this._isPanning) return;
    if (this._panDistance > 8) this._lastPanEndTime = Date.now();
    this._isPanning = false;
    this._setPanningClass(false);
  }

  /* ── Drag-to-pan (touch) ── */

  private _handleTouchStart = (e: React.TouchEvent<HTMLDivElement>): void => {
    if ((e.target as HTMLElement).closest('button, a, input')) return;
    const touch = e.touches[0];
    this._isPanning    = true;
    this._panDistance  = 0;
    this._panStartX    = touch.clientX;
    this._panStartY    = touch.clientY;
    this._scrollStartX = this._scrollEl?.scrollLeft ?? 0;
    this._scrollStartY = this._scrollEl?.scrollTop  ?? 0;
    this._setPanningClass(true);
  }

  private _handleTouchMoveDirect = (e: TouchEvent): void => {
    if (!this._isPanning || !this._scrollEl) return;
    e.preventDefault();
    const touch = e.touches[0];
    const dx = touch.clientX - this._panStartX;
    const dy = touch.clientY - this._panStartY;
    this._panDistance = Math.sqrt(dx * dx + dy * dy);
    this._scrollEl.scrollLeft = this._scrollStartX - dx;
    this._scrollEl.scrollTop  = this._scrollStartY - dy;
  }

  private _handleTouchEnd = (): void => {
    if (!this._isPanning) return;
    if (this._panDistance > 8) this._lastPanEndTime = Date.now();
    this._isPanning = false;
    this._setPanningClass(false);
  }

  /* ── Render ── */

  public render(): React.ReactElement {
    const {
      rootNode, isLoading, error, photos, expandingNodes, searchQuery, appliedQuery,
      showSearchResults, searchActiveIndex,
      presenceMap, zoomLevel, selectedUser, showFilters,
      filterMembers, filterGuests, filterDisabled,
      focusedUser, ancestorChain, allUsers,
      personCardManagerChain, chartLayout, findMeError,
      filterDepartments, showDeptFilter, showStats, showLayoutPicker,
      rootPickerQuery, rootPickerResults, runtimeRootUser,
      isRefocusing, isExpandingAll, treeFocusId,
    } = this.state;
    const { showDepartment, showOffice, theme, accentColor, currentUserEmail, compactCards,
      enableFindMe, enableLayoutToggle, enableStats, enableDeptFilter, enableUserFilter } = this.props;

    if (isLoading) return (
      <div className={styles.centered}><Spinner size={SpinnerSize.large} label={strings.Chart_BuildingLabel} /></div>
    );

    if (error) return (
      <div className={styles.errorState} role="alert">
        <Icon iconName="Warning" className={styles.errorIcon} />
        <div className={styles.errorText}>{error}</div>
        <DefaultButton text={strings.Chart_RetryButton} onClick={this._handleRetry} />
      </div>
    );

    if (!rootNode) return <NoConfigForm onLoad={this._handleSetupLoad} />;

    const t = getThemeTokens(theme, accentColor);
    // Full-tree scans are cached by reference — renders fire on every pan/zoom
    // state change, and the tree only changes when rootNode is replaced
    if (this._treeScanFor !== rootNode) {
      this._treeScanFor = rootNode;
      this._treeCounts  = countTreeUsers(rootNode);
      this._uniqueDepts = getUniqueDepts(rootNode);
    }
    const treeCounts   = this._treeCounts;
    const uniqueDepts  = this._uniqueDepts;
    if (showStats && this._statsFor !== allUsers) {
      this._statsFor = allUsers;
      this._stats    = computeStats(allUsers);
    }
    const stats        = showStats ? this._stats : null;
    const isDrillMode  = chartLayout === 'drill';
    const resetZoom    = this.props.defaultZoom > 0 ? this.props.defaultZoom : 1;
    const isBusy       = isRefocusing || isExpandingAll;

    const containerClasses = [
      styles.container,
      THEME_CONTAINER_CLASS[theme],
    ].filter(Boolean).join(' ');

    const treeScrollClasses = [
      styles.treeScroll,
      chartLayout === 'horizontal' ? styles.layoutHorizontal : '',
      compactCards ? styles.compactMode : '',
      isRefocusing ? styles.refocusing : '',
    ].filter(Boolean).join(' ');

    let visibleIds = EMPTY_IDS;
    let tabStopId = rootNode.user.id;
    if (!isDrillMode) {
      visibleIds = this._getVisibleIds();
      this._getRenderedNodes();
      if (treeFocusId && this._renderedIds.has(treeFocusId)) tabStopId = treeFocusId;
    }

    return (
      <div className={containerClasses} style={getThemeContainerStyle(theme, accentColor)}>

        {/* ── Toolbar ── */}
        <OrgChartToolbar
          theme={theme}
          accentColor={accentColor}
          isDrillMode={isDrillMode}
          searchRef={this._searchRef}
          idPrefix={this._uid}
          searchQuery={searchQuery}
          appliedQuery={appliedQuery}
          matchCount={appliedQuery && !isDrillMode ? this._getMatchCount(appliedQuery) : 0}
          searchResults={this._getSearchResults(appliedQuery)}
          showSearchResults={showSearchResults}
          searchActiveIndex={searchActiveIndex}
          onSearchChange={this._onSearchChange}
          onSearchFocus={this._onSearchFocus}
          onSearchKeyDown={this._onSearchKeyDown}
          onSearchEnter={this._onSearchEnter}
          onSearchEscape={this._onSearchEscape}
          onSelectSearchResult={this._selectSearchResult}
          isExpandingAll={isExpandingAll}
          isRefocusing={isRefocusing}
          onExpandAll={this._handleExpandAll}
          onCollapseAll={this._handleCollapseAll}
          showFindMe={!!currentUserEmail && enableFindMe}
          onFindMe={this._handleFindMe}
          enableLayoutToggle={enableLayoutToggle}
          chartLayout={chartLayout}
          showLayoutPicker={showLayoutPicker}
          onToggleLayoutPicker={this._toggleLayoutPicker}
          onSelectLayout={this._selectLayout}
          enableStats={enableStats}
          showStats={showStats}
          onToggleStats={this._toggleStats}
          enableDeptFilter={enableDeptFilter}
          uniqueDepts={uniqueDepts}
          filterDepartments={filterDepartments}
          showDeptFilter={showDeptFilter}
          onToggleDeptFilter={this._toggleDeptFilter}
          onToggleDept={this._toggleDept}
          onClearDeptFilter={this._clearDeptFilter}
          enableUserFilter={enableUserFilter}
          filterMembers={filterMembers}
          filterGuests={filterGuests}
          filterDisabled={filterDisabled}
          treeCounts={treeCounts}
          showFilters={showFilters}
          onToggleUserFilterPanel={this._toggleUserFilterPanel}
          onToggleUserFilter={this._toggleUserFilter}
          onExportPdf={this._exportPdfClick}
          onExportCsv={this._exportCsv}
          rootPickerRef={this._rootPickerRef}
          runtimeRootUser={runtimeRootUser}
          rootPickerQuery={rootPickerQuery}
          rootPickerResults={rootPickerResults}
          onRootPickerChange={this._onRootPickerChange}
          onRootPickerSelect={this._onRootPickerSelect}
          onResetRoot={this._resetRoot}
          zoomLevel={zoomLevel}
          resetZoom={resetZoom}
          onZoomOut={this._zoomOut}
          onZoomIn={this._zoomIn}
          onZoomReset={this._zoomReset}
        />

        {/* ── Find Me error toast ── */}
        {findMeError && (
          <div className={styles.findMeToast} role="alert">{findMeError}</div>
        )}

        {/* ── Stats bar ── */}
        {stats && (
          <div className={styles.statsBar}>
            <div className={styles.statItem}><span className={styles.statValue}>{stats.total}</span><span className={styles.statLabel}>{strings.Chart_StatPeople}</span></div>
            <div className={styles.statItem}><span className={styles.statValue}>{stats.members}</span><span className={styles.statLabel}>{strings.Chart_StatMembers}</span></div>
            {stats.guests > 0 && <div className={styles.statItem}><span className={styles.statValue}>{stats.guests}</span><span className={styles.statLabel}>{strings.Chart_StatGuests}</span></div>}
            <div className={styles.statItem}><span className={styles.statValue}>{stats.depts}</span><span className={styles.statLabel}>{strings.Chart_StatDepts}</span></div>
          </div>
        )}

        {/* ── Ancestor strip (full-tree mode only) ── */}
        {!isDrillMode && focusedUser && (
          <nav className={styles.ancestorStrip} aria-label={strings.Chart_ReportingLineAria}>
            <button className={styles.ancestorReturnBtn} onClick={this._handleReturnToRoot} title={strings.Chart_BackToFullOrgTitle}>
              <Icon iconName="Home" /> {strings.Chart_FullOrgLabel}
            </button>
            <Icon iconName="ChevronRight" className={styles.ancestorChevron} />
            {ancestorChain.map(ancestor => (
              <React.Fragment key={ancestor.id}>
                <button
                  className={styles.ancestorLink}
                  onClick={() => this._handleFocusUser(ancestor)}
                  title={formatString(strings.Chart_FocusOnPerson, { name: ancestor.displayName })}
                >
                  <span className={styles.ancestorInitials} style={{ background: t.accent, color: t.onAccent }}>
                    {getInitials(ancestor.displayName)}
                  </span>
                  <span className={styles.ancestorName}>{ancestor.displayName.split(' ')[0]}</span>
                </button>
                <Icon iconName="ChevronRight" className={styles.ancestorChevron} />
              </React.Fragment>
            ))}
            <span className={styles.ancestorCurrent} aria-current="page">
              <span className={styles.ancestorInitials} style={{ background: t.accent, color: t.onAccent }}>
                {getInitials(focusedUser.displayName)}
              </span>
              <span className={styles.ancestorName}>{focusedUser.displayName}</span>
            </span>
          </nav>
        )}

        {/* ── DRILL-DOWN VIEW ── */}
        {isDrillMode && (
          <DrillView
            drillPath={this.state.drillPath}
            visibleReports={this._getVisibleDrillReports()}
            allReports={this.state.drillReports}
            drillLoadingId={this.state.drillLoadingId}
            photos={photos}
            presenceMap={presenceMap}
            drillReportCounts={this.state.drillReportCounts}
            graphService={this.props.graphService}
            theme={theme}
            accentColor={accentColor}
            showDepartment={showDepartment}
            showOffice={showOffice}
            compactCards={compactCards}
            isBusy={isRefocusing}
            headerRef={this._drillHeaderRef}
            onReturnToRoot={this._handleReturnToRoot}
            onNavigate={this._handleDrillNavigate}
            onDrillInto={this._handleDrillInto}
            onDrillToggle={this._handleDrillToggle}
            onShowProfile={this._handleCardClick}
          />
        )}

        {/* ── FULL TREE VIEW ── */}
        {!isDrillMode && (
          <div
            ref={this._setScrollRef}
            className={treeScrollClasses}
            style={{ position: 'relative' }}
            onMouseDown={this._handlePanStart}
            onMouseMove={this._handlePanMove}
            onMouseUp={this._handlePanEnd}
            onMouseLeave={this._handlePanEnd}
            onTouchStart={this._handleTouchStart}
            onTouchEnd={this._handleTouchEnd}
            aria-busy={isBusy}
          >
            <div
              style={
                chartLayout === 'horizontal'
                  ? { zoom: zoomLevel, display: 'inline-block', minWidth: '100%' }
                  // Vertical (top-down): the root's own box shrinks to fit its
                  // subtree width, so without an explicit centering flex wrapper
                  // it lands flush left instead of over the middle of its tree.
                  : { zoom: zoomLevel, display: 'flex', justifyContent: 'center', minWidth: '100%' }
              }
              role="tree"
              aria-label={strings.Chart_OrgChartAria}
              onKeyDown={this._handleTreeKeyDown}
            >
              <OrgTree
                node={rootNode}
                photos={photos}
                presenceMap={presenceMap}
                showDepartment={showDepartment}
                showOffice={showOffice}
                expandingNodes={expandingNodes}
                searchQuery={appliedQuery}
                theme={theme}
                accentColor={accentColor}
                visibleIds={visibleIds}
                compactCards={compactCards}
                depth={1}
                posInSet={1}
                setSize={1}
                tabStopId={tabStopId}
                onToggle={this._handleToggle}
                onCardClick={this._handleCardClick}
                onFocus={this._handleFocusFromCard}
                onItemFocus={this._handleTreeItemFocus}
              />
            </div>
          </div>
        )}

        {/* ── Popups backdrop ── */}
        {(showFilters || showDeptFilter || showLayoutPicker) && (
          <div className={styles.popupBackdrop} onClick={this._closePopups} />
        )}

        {/* ── Person card ── */}
        {selectedUser && (
          <PersonCard
            user={selectedUser}
            photo={photos[selectedUser.id] ?? null}
            presence={presenceMap.get(selectedUser.id)}
            theme={theme}
            accentColor={accentColor}
            managerChain={personCardManagerChain}
            dottedManager={this.state.personCardDottedManager}
            dottedReports={this.state.personCardDottedReports}
            directReportCount={this.props.graphService?.getDirectReportCount(selectedUser.id) ?? 0}
            totalReportCount={this.props.graphService?.getTotalReportCount(selectedUser.id) ?? 0}
            customAttributes={this.props.customAttributes}
            onClose={this._closePersonCard}
            onFocus={this._handleFocusFromCard}
          />
        )}
      </div>
    );
  }
}

function sameKey(a: unknown[], b: unknown[]): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

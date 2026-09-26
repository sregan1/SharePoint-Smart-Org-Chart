import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { IGraphUser, PresenceAvailability, ICustomAttributeConfig } from '../../../../services/GraphService';
import { IEmployeeDirectoryProps } from './IEmployeeDirectoryProps';
import { exportDirectoryToExcel } from '../../../../services/PdfExportService';
import { PRESENCE_COLOR, PRESENCE_LABEL, getInitials } from '../personUtils';
import { getAccentCssVars } from '../colorUtils';
import { formatString } from '../localeUtils';
import * as strings from 'SmartOrgChartWebPartStrings';
import styles from './EmployeeDirectory.module.scss';

// '#' collects names that don't start with A–Z (digits, symbols, other scripts)
const OTHER_BUCKET = '#';
const ALPHABET = ['All', ...'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split(''), OTHER_BUCKET];

const SEARCH_DEBOUNCE_MS = 200;

// Letters that don't decompose under NFD but read as a Latin base letter
const SPECIAL_LETTERS: { [ch: string]: string } = {
  'Ø': 'O', 'Œ': 'O', 'Æ': 'A', 'Ð': 'D', 'Đ': 'D', 'Ł': 'L', 'Þ': 'T', 'ß': 'S', 'ẞ': 'S',
};

/** Alphabet bucket for a name: its first letter without diacritics, or '#'. */
function letterBucket(name: string): string {
  const first = (name || '').trim().charAt(0);
  if (!first) return OTHER_BUCKET;
  let ch = first.normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().charAt(0);
  if (SPECIAL_LETTERS[ch]) ch = SPECIAL_LETTERS[ch];
  return ch >= 'A' && ch <= 'Z' ? ch : OTHER_BUCKET;
}

type ViewMode = 'card' | 'list';

const LS_VIEWMODE_KEY = 'smartOrgChart_dirViewMode';

function viewModeKey(instanceId: string): string {
  return instanceId ? `${LS_VIEWMODE_KEY}_${instanceId}` : LS_VIEWMODE_KEY;
}

function readViewMode(instanceId: string): ViewMode {
  try {
    const v = localStorage.getItem(viewModeKey(instanceId));
    if (v === 'card' || v === 'list') return v;
  } catch { /* ignore */ }
  return 'card';
}

/** Per-user values derived once when users load, so filtering stays cheap. */
interface IUserIndex {
  searchKey: string;
  firstLetter: string;
  lastLetter: string;
}

interface IFilterCache {
  users: IGraphUser[];
  query: string;
  letter: string;
  dept: string;
  office: string;
  field: 'firstName' | 'lastName';
  result: IGraphUser[];
}

interface IEmployeeDirectoryState {
  users: IGraphUser[];
  isLoading: boolean;
  error: string | null;
  selectedLetter: string;
  /** What the search box shows — updates on every keystroke */
  searchInput: string;
  /** Debounced copy of searchInput that the filter actually uses */
  searchQuery: string;
  photos: { [id: string]: string | null };
  presenceMap: Map<string, PresenceAvailability>;
  currentPage: number;
  viewMode: ViewMode;
  selectedDepartment: string;
  selectedOffice: string;
}

export class EmployeeDirectory extends React.Component<IEmployeeDirectoryProps, IEmployeeDirectoryState> {
  private _photoQueue: string[] = [];
  private _photoQueueSet: Set<string> = new Set();
  private _processingPhotos = false;
  private _mounted = false;
  private _presenceInterval: number | null = null;
  private _presenceDebounce: number | null = null;
  private _searchDebounce: number | null = null;
  private _containerRef = React.createRef<HTMLDivElement>();
  private _gridRef = React.createRef<HTMLDivElement>();

  // Memoization — rebuilt only when their inputs change
  private _index: Map<string, IUserIndex> = new Map();
  private _indexUsers: IGraphUser[] | null = null;
  private _filterCache: IFilterCache | null = null;
  private _optionsUsers: IGraphUser[] | null = null;
  private _deptOptions: IDropdownOption[] = [];
  private _officeOptions: IDropdownOption[] = [];
  // Admin-configured attributes visible in the Directory — recomputed only
  // when props.customAttributes changes, not on every render/keystroke.
  private _dirAttrsSource: ICustomAttributeConfig[] | null = null;
  private _dirAttrs: ICustomAttributeConfig[] = [];

  constructor(props: IEmployeeDirectoryProps) {
    super(props);
    this.state = {
      users: [], isLoading: true, error: null,
      selectedLetter: 'All', searchInput: '', searchQuery: '',
      photos: {}, presenceMap: new Map(), currentPage: 1,
      viewMode: readViewMode(props.instanceId),
      selectedDepartment: '', selectedOffice: ''
    };
  }

  public async componentDidMount(): Promise<void> {
    this._mounted = true;
    await this._loadUsers();
    if (!this._mounted) return;
    this._refreshPresence().catch(() => { /* ignore */ });
    this._presenceInterval = window.setInterval(() => {
      this._refreshPresence().catch(() => { /* ignore */ });
    }, 60_000);
  }

  public componentDidUpdate(prevProps: IEmployeeDirectoryProps, prevState: IEmployeeDirectoryState): void {
    if (prevProps.alphabetFilterField !== this.props.alphabetFilterField) {
      this.setState({ selectedLetter: 'All', currentPage: 1 });
      return;
    }
    if (
      prevState.users !== this.state.users ||
      prevState.selectedLetter !== this.state.selectedLetter ||
      prevState.searchQuery !== this.state.searchQuery ||
      prevState.currentPage !== this.state.currentPage ||
      prevState.selectedDepartment !== this.state.selectedDepartment ||
      prevState.selectedOffice !== this.state.selectedOffice ||
      prevProps.pageSize !== this.props.pageSize
    ) {
      const paged = this._getCurrentPageUsers();
      this._enqueuePhotos(paged.map(u => u.id));
      // Presence for the newly visible page — debounced so rapid filter
      // changes don't fire a Graph call for every transient page
      if (this._presenceDebounce !== null) window.clearTimeout(this._presenceDebounce);
      this._presenceDebounce = window.setTimeout(() => {
        this._presenceDebounce = null;
        this._refreshPresence().catch(() => { /* ignore */ });
      }, 500);
    }
  }

  public componentWillUnmount(): void {
    this._mounted = false;
    if (this._presenceInterval !== null) {
      window.clearInterval(this._presenceInterval);
      this._presenceInterval = null;
    }
    if (this._presenceDebounce !== null) {
      window.clearTimeout(this._presenceDebounce);
      this._presenceDebounce = null;
    }
    if (this._searchDebounce !== null) {
      window.clearTimeout(this._searchDebounce);
      this._searchDebounce = null;
    }
  }

  private async _refreshPresence(): Promise<void> {
    if (!this._mounted) return;
    // Only request presence for the page of users currently on screen
    const paged = this._getCurrentPageUsers();
    if (paged.length === 0) return;
    const result = await this.props.graphService.getPresence(paged.map(u => u.id));
    // Copy so a map the service keeps mutating can't change state behind React's back
    if (this._mounted) this.setState({ presenceMap: new Map(result) });
  }

  private _getCurrentPageUsers(): IGraphUser[] {
    const filtered = this._getFilteredUsers();
    const { pageSize } = this.props;
    const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
    const safePage = Math.min(this.state.currentPage, totalPages);
    return filtered.slice((safePage - 1) * pageSize, safePage * pageSize);
  }

  public exportExcel(): void {
    const { showEmail, showPhone, showDepartment, showOffice } = this.props;
    exportDirectoryToExcel(this._getFilteredUsers(), { showEmail, showPhone, showDepartment, showOffice });
  }

  private async _loadUsers(): Promise<void> {
    try {
      // GraphService/MockGraphService both return the list pre-sorted by display name
      const users = await this.props.graphService.getAllUsers();
      if (!this._mounted) return;
      this._buildIndex(users);
      this.setState({ users, isLoading: false });
    } catch (err) {
      const detail = err instanceof Error && err.message
        ? err.message
        : strings.Directory_LoadFailedFallback;
      if (this._mounted) this.setState({ isLoading: false, error: formatString(strings.Directory_LoadFailedPrefix, { detail }) });
    }
  }

  private _buildIndex(users: IGraphUser[]): void {
    if (this._indexUsers === users) return;
    const index = new Map<string, IUserIndex>();
    users.forEach(u => {
      index.set(u.id, {
        searchKey: [u.displayName, u.mail, u.jobTitle, u.department]
          .map(v => (v || '').toLowerCase())
          .join('\n'),
        firstLetter: letterBucket(this._getFirstName(u.displayName)),
        lastLetter: letterBucket(this._getLastName(u.displayName)),
      });
    });
    this._index = index;
    this._indexUsers = users;
  }

  private _enqueuePhotos(ids: string[]): void {
    const toLoad = ids.filter(id => !(id in this.state.photos) && !this._photoQueueSet.has(id));
    toLoad.forEach(id => { this._photoQueue.push(id); this._photoQueueSet.add(id); });
    if (!this._processingPhotos) this._drainPhotoQueue().catch(() => { /* ignore */ });
  }

  private async _drainPhotoQueue(): Promise<void> {
    this._processingPhotos = true;
    // Batch setState calls — one render per 10 photos instead of one per photo
    let batch: { [id: string]: string | null } = {};
    const flush = (): void => {
      const toApply = batch;
      batch = {};
      if (this._mounted && Object.keys(toApply).length > 0) {
        this.setState(prev => ({ photos: { ...prev.photos, ...toApply } }));
      }
    };
    try {
      while (this._photoQueue.length > 0 && this._mounted) {
        const id = this._photoQueue.shift();
        if (!id) break;
        this._photoQueueSet.delete(id);
        if (id in this.state.photos || id in batch) continue;
        try {
          batch[id] = await this.props.graphService.getUserPhoto(id);
        } catch {
          batch[id] = null;   // show initials rather than stalling the queue
        }
        if (Object.keys(batch).length >= 10) flush();
      }
      flush();
    } finally {
      this._processingPhotos = false;
    }
  }

  /** A photo URL that fails to load falls back to initials. */
  private _onPhotoError = (userId: string): void => {
    if (!this._mounted) return;
    this.setState(prev => ({ photos: { ...prev.photos, [userId]: null } }));
  }

  private _getFirstName(dn: string): string { return (dn || '').split(' ')[0] || ''; }
  private _getLastName(dn: string): string {
    const p = (dn || '').split(' ').filter(x => x);
    return p.length > 1 ? p[p.length - 1] : p[0] || '';
  }
  private _getPhone(user: IGraphUser): string {
    return user.mobilePhone || (user.businessPhones && user.businessPhones[0]) || '';
  }

  /** Directory-visible custom attribute configs — memoized on the props array reference. */
  private _getDirectoryAttributes(): ICustomAttributeConfig[] {
    const { customAttributes } = this.props;
    if (this._dirAttrsSource !== customAttributes) {
      this._dirAttrs = (customAttributes || []).filter(c => c.showInDirectory && c.graphField && c.label);
      this._dirAttrsSource = customAttributes;
    }
    return this._dirAttrs;
  }

  /** Department and office dropdown options — rebuilt only when the user list changes. */
  private _ensureOptions(): void {
    const { users } = this.state;
    if (this._optionsUsers === users) return;
    const collect = (get: (u: IGraphUser) => string | undefined): string[] => {
      const seen: { [k: string]: boolean } = {};
      const out: string[] = [];
      users.forEach(u => { const v = get(u); if (v && !seen[v]) { seen[v] = true; out.push(v); } });
      return out.sort();
    };
    this._deptOptions = [{ key: '', text: strings.Directory_AllDepartments }, ...collect(u => u.department).map(d => ({ key: d, text: d }))];
    this._officeOptions = [{ key: '', text: strings.Directory_AllOffices }, ...collect(u => u.officeLocation).map(o => ({ key: o, text: o }))];
    this._optionsUsers = users;
  }

  private _getFilteredUsers(): IGraphUser[] {
    const { users, selectedLetter, searchQuery, selectedDepartment, selectedOffice } = this.state;
    const { alphabetFilterField } = this.props;
    const query = searchQuery.trim().toLowerCase();

    const c = this._filterCache;
    if (c && c.users === users && c.query === query && c.letter === selectedLetter &&
        c.dept === selectedDepartment && c.office === selectedOffice && c.field === alphabetFilterField) {
      return c.result;
    }

    this._buildIndex(users);
    const index = this._index;
    let result = users;

    if (query) {
      result = result.filter(u => {
        const entry = index.get(u.id);
        return !!entry && entry.searchKey.indexOf(query) >= 0;
      });
    } else if (selectedLetter !== 'All') {
      result = result.filter(u => {
        const entry = index.get(u.id);
        if (!entry) return false;
        return (alphabetFilterField === 'firstName' ? entry.firstLetter : entry.lastLetter) === selectedLetter;
      });
    }

    if (selectedDepartment) result = result.filter(u => u.department === selectedDepartment);
    if (selectedOffice) result = result.filter(u => u.officeLocation === selectedOffice);

    this._filterCache = {
      users, query, letter: selectedLetter, dept: selectedDepartment,
      office: selectedOffice, field: alphabetFilterField, result,
    };
    return result;
  }

  private _cancelSearchDebounce(): void {
    if (this._searchDebounce !== null) {
      window.clearTimeout(this._searchDebounce);
      this._searchDebounce = null;
    }
  }

  private _selectLetter = (letter: string): void => {
    this._cancelSearchDebounce();
    this.setState({ selectedLetter: letter, searchInput: '', searchQuery: '', currentPage: 1 });
  }

  private _onSearch = (value: string): void => {
    const v = value || '';
    this._cancelSearchDebounce();
    // The input updates immediately; the (potentially large) filter waits
    // until typing pauses. Clearing the box applies at once.
    if (!v.trim()) {
      this.setState({ searchInput: v, searchQuery: '', selectedLetter: 'All', currentPage: 1 });
      return;
    }
    this.setState({ searchInput: v });
    this._searchDebounce = window.setTimeout(() => {
      this._searchDebounce = null;
      if (this._mounted) this.setState({ searchQuery: v, selectedLetter: 'All', currentPage: 1 });
    }, SEARCH_DEBOUNCE_MS);
  }

  private _clearFilters = (): void => {
    this._cancelSearchDebounce();
    this.setState({ selectedDepartment: '', selectedOffice: '', selectedLetter: 'All', searchInput: '', searchQuery: '', currentPage: 1 });
  }

  private _setViewMode = (mode: ViewMode): void => {
    try { localStorage.setItem(viewModeKey(this.props.instanceId), mode); } catch { /* ignore */ }
    this.setState({ viewMode: mode });
  }

  private _goToPage = (page: number): void => {
    this.setState({ currentPage: page }, () => {
      const grid = this._gridRef.current;
      if (grid) grid.scrollTop = 0;
      // Bring the top of the directory back into view when the user paged
      // from further down (e.g. the bottom pagination bar)
      const el = this._containerRef.current;
      if (!el) return;
      let top = el.getBoundingClientRect().top;
      let scrollParent = el.parentElement;
      while (scrollParent && scrollParent !== document.body) {
        const oy = window.getComputedStyle(scrollParent).overflowY;
        if ((oy === 'auto' || oy === 'scroll') && scrollParent.scrollHeight > scrollParent.clientHeight) break;
        scrollParent = scrollParent.parentElement;
      }
      const parentTop = scrollParent && scrollParent !== document.body
        ? Math.max(0, scrollParent.getBoundingClientRect().top)
        : 0;
      top -= parentTop;
      if (top < 0 && el.scrollIntoView) el.scrollIntoView({ block: 'start' });
    });
  }

  /* ── Render helpers ── */

  private _renderPresence(userId: string, small?: boolean): React.ReactElement | null {
    const s = this.state.presenceMap.get(userId);
    if (!s || s === 'Unknown') return null;
    const label = PRESENCE_LABEL[s] || s;
    return (
      <>
        <span
          className={styles.presenceDot}
          style={small ? { background: PRESENCE_COLOR[s], width: 10, height: 10 } : { background: PRESENCE_COLOR[s] }}
          aria-hidden="true"
          title={label}
        />
        <span className={styles.srOnly}>{formatString(strings.Directory_PresenceStatus, { status: label })}</span>
      </>
    );
  }

  private _renderAvatar(user: IGraphUser, imgClass: string, initialsClass: string): React.ReactElement {
    const photoUrl = this.state.photos[user.id] || null;
    return photoUrl
      ? (
        <img
          src={photoUrl}
          alt={user.displayName}
          className={imgClass}
          loading="lazy"
          onError={() => this._onPhotoError(user.id)}
        />
      )
      : <div className={initialsClass} aria-hidden="true">{getInitials(user.displayName)}</div>;
  }

  private _teamsChatUrl(mail: string): string {
    return `https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(mail)}`;
  }

  private _renderCardGrid(paged: IGraphUser[]): React.ReactElement {
    const { cardSize, showEmail, showPhone, showDepartment, showOffice } = this.props;
    const dirAttrs = this._getDirectoryAttributes();

    return (
      <div ref={this._gridRef} className={`${styles.grid} ${styles[`size_${cardSize}`]}`}>
        {paged.map(user => {
          const phone = this._getPhone(user);
          const visibleAttrs = dirAttrs.filter(a => !!(user.customAttributes && user.customAttributes[a.graphField]));
          const hasDetails =
            (showDepartment && !!user.department) ||
            (showOffice && !!user.officeLocation) ||
            (showEmail && !!user.mail) ||
            (showPhone && !!phone) ||
            visibleAttrs.length > 0;
          return (
            <div key={user.id} className={styles.card}>
              <div className={styles.cardHeader}>
                <div className={styles.avatarWrap}>
                  {this._renderAvatar(user, styles.avatar, styles.initials)}
                  {this._renderPresence(user.id)}
                </div>
                <div className={styles.cardBody}>
                  <div className={styles.userName}>{user.displayName}</div>
                  {user.jobTitle && <div className={styles.jobTitle}>{user.jobTitle}</div>}
                  <div className={styles.statusBadges}>
                    {user.accountEnabled === false && (
                      <span className={`${styles.statusBadge} ${styles.statusDisabled}`}>{strings.Directory_StatusDisabled}</span>
                    )}
                    {user.userType === 'Guest' && (
                      <span className={`${styles.statusBadge} ${styles.statusGuest}`}>{strings.Directory_StatusGuest}</span>
                    )}
                  </div>
                </div>
                {user.mail && (
                  <a
                    href={this._teamsChatUrl(user.mail)}
                    target="_blank" rel="noopener noreferrer"
                    className={styles.chatBtn} title={strings.Directory_ChatInTeams}
                    aria-label={formatString(strings.Directory_ChatWithInTeams, { name: user.displayName })}
                    onClick={e => e.stopPropagation()}
                  >
                    <Icon iconName="Chat" aria-hidden="true" />
                  </a>
                )}
              </div>
              {hasDetails && (
                <div className={styles.cardDetails}>
                  {showDepartment && user.department && (
                    <div className={styles.detail}>
                      <Icon iconName="Work" className={styles.detailIcon} />
                      <span>{user.department}</span>
                    </div>
                  )}
                  {showOffice && user.officeLocation && (
                    <div className={styles.detail}>
                      <Icon iconName="POI" className={styles.detailIcon} />
                      <span>{user.officeLocation}</span>
                    </div>
                  )}
                  {showEmail && user.mail && (
                    <div className={styles.detail}>
                      <Icon iconName="Mail" className={styles.detailIcon} />
                      <a href={`mailto:${user.mail}`} className={styles.link}>{user.mail}</a>
                    </div>
                  )}
                  {showPhone && phone && (
                    <div className={styles.detail}>
                      <Icon iconName="Phone" className={styles.detailIcon} />
                      <a href={`tel:${phone}`} className={styles.link}>{phone}</a>
                    </div>
                  )}
                  {visibleAttrs.map(attr => (
                    <div className={styles.detail} key={attr.id} title={attr.label}>
                      <Icon iconName="Tag" className={styles.detailIcon} />
                      <span>{attr.label}: {user.customAttributes?.[attr.graphField]}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>
    );
  }

  private _renderListView(paged: IGraphUser[]): React.ReactElement {
    const { showEmail, showPhone, showDepartment, showOffice } = this.props;
    const dirAttrs = this._getDirectoryAttributes();

    return (
      <table className={styles.listTable}>
        <thead>
          <tr className={styles.listHead}>
            <th scope="col" className={styles.listTh}>{strings.Directory_ColumnName}</th>
            <th scope="col" className={styles.listTh}>{strings.Directory_ColumnJobTitle}</th>
            {showDepartment && <th scope="col" className={styles.listTh}>{strings.Directory_ColumnDepartment}</th>}
            {showOffice && <th scope="col" className={styles.listTh}>{strings.Directory_ColumnOffice}</th>}
            {dirAttrs.map(attr => (
              <th scope="col" className={styles.listTh} key={attr.id} title={attr.label}>{attr.label}</th>
            ))}
            {showEmail && <th scope="col" className={styles.listTh}>{strings.Directory_ColumnEmail}</th>}
            {showPhone && <th scope="col" className={styles.listTh}>{strings.Directory_ColumnPhone}</th>}
            <th scope="col" className={styles.listTh}><span className={styles.srOnly}>{strings.Directory_ColumnChat}</span></th>
          </tr>
        </thead>
        <tbody>
          {paged.map(user => {
            const phone = this._getPhone(user);
            return (
              <tr key={user.id} className={styles.listRow}>
                <td className={styles.listTd}>
                  <div className={styles.listNameCell}>
                    <div className={styles.listAvatarWrap}>
                      {this._renderAvatar(user, styles.listAvatar, styles.listInitials)}
                      {this._renderPresence(user.id, true)}
                    </div>
                    <span className={styles.listName}>{user.displayName}</span>
                    {user.accountEnabled === false && (
                      <span className={`${styles.statusBadge} ${styles.statusDisabled}`}>{strings.Directory_StatusDisabled}</span>
                    )}
                    {user.userType === 'Guest' && (
                      <span className={`${styles.statusBadge} ${styles.statusGuest}`}>{strings.Directory_StatusGuest}</span>
                    )}
                  </div>
                </td>
                <td className={styles.listTd}><span className={styles.listCell}>{user.jobTitle || '—'}</span></td>
                {showDepartment && <td className={styles.listTd}><span className={styles.listCell}>{user.department || '—'}</span></td>}
                {showOffice && <td className={styles.listTd}><span className={styles.listCell}>{user.officeLocation || '—'}</span></td>}
                {dirAttrs.map(attr => {
                  const value = user.customAttributes && user.customAttributes[attr.graphField];
                  return (
                    <td className={styles.listTd} key={attr.id}>
                      {value
                        ? <span className={styles.listCustomAttrCell} title={value}>{value}</span>
                        : <span className={styles.listCell}>—</span>}
                    </td>
                  );
                })}
                {showEmail && (
                  <td className={styles.listTd}>
                    {user.mail
                      ? <a href={`mailto:${user.mail}`} className={styles.listLink}>{user.mail}</a>
                      : <span className={styles.listCell}>—</span>}
                  </td>
                )}
                {showPhone && (
                  <td className={styles.listTd}>
                    {phone
                      ? <a href={`tel:${phone}`} className={styles.listLink}>{phone}</a>
                      : <span className={styles.listCell}>—</span>}
                  </td>
                )}
                <td className={styles.listTd}>
                  {user.mail && (
                    <a
                      href={this._teamsChatUrl(user.mail)}
                      target="_blank" rel="noopener noreferrer"
                      className={styles.listChatBtn} title={strings.Directory_ChatInTeams}
                      aria-label={formatString(strings.Directory_ChatWithInTeams, { name: user.displayName })}
                    >
                      <Icon iconName="Chat" aria-hidden="true" />
                    </a>
                  )}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    );
  }

  private _renderPager(safePage: number, totalPages: number): React.ReactElement {
    return (
      <>
        <button
          className={styles.pageBtn}
          onClick={() => this._goToPage(Math.max(1, safePage - 1))}
          disabled={safePage === 1}
          aria-label={strings.Directory_PreviousPageAria}
        >
          <Icon iconName="ChevronLeft" aria-hidden="true" /> {strings.Directory_PrevPage}
        </button>
        <span className={styles.pageInfo}>{safePage} / {totalPages}</span>
        <button
          className={styles.pageBtn}
          onClick={() => this._goToPage(Math.min(totalPages, safePage + 1))}
          disabled={safePage === totalPages}
          aria-label={strings.Directory_NextPageAria}
        >
          {strings.Directory_NextPage} <Icon iconName="ChevronRight" aria-hidden="true" />
        </button>
      </>
    );
  }

  public render(): React.ReactElement {
    const { isLoading, error, selectedLetter, currentPage, searchInput,
            viewMode, selectedDepartment, selectedOffice } = this.state;
    const { pageSize } = this.props;

    if (isLoading) return (
      <div className={styles.centered}><Spinner size={SpinnerSize.large} label={strings.Directory_LoadingLabel} /></div>
    );

    if (error) return (
      <div className={styles.errorMsg}><Icon iconName="Warning" /><span>{error}</span></div>
    );

    this._ensureOptions();
    const deptOptions = this._deptOptions;
    const officeOptions = this._officeOptions;
    const filtered = this._getFilteredUsers();
    const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
    const safePage = Math.min(currentPage, totalPages);
    const paged = filtered.slice((safePage - 1) * pageSize, safePage * pageSize);

    const activeFilters = (selectedDepartment ? 1 : 0) + (selectedOffice ? 1 : 0);
    const hasSearch = !!searchInput.trim();
    const anythingToClear = activeFilters > 0 || selectedLetter !== 'All' || hasSearch;

    const THEME_CLASS: Partial<Record<string, string>> = {
      minimal: styles.themeMinimal,
      corporate: styles.themeCorporate,
      dark: styles.themeDark,
      custom: styles.themeCustom,
    };
    const themeClass = THEME_CLASS[this.props.theme] || '';
    const customStyle = this.props.theme === 'custom'
      ? getAccentCssVars(this.props.accentColor)
      : undefined;

    return (
      <div ref={this._containerRef} className={[styles.container, themeClass].filter(Boolean).join(' ')} style={customStyle}>

        {/* ── Top bar: search + filters + view toggle ── */}
        <div className={styles.toolbar}>
          <SearchBox
            placeholder={strings.Directory_SearchPlaceholder}
            ariaLabel={strings.Directory_SearchAria}
            value={searchInput}
            onChange={(_, v) => this._onSearch(v || '')}
            className={styles.searchBox}
            underlined
          />

          {deptOptions.length > 2 && (
            <Dropdown
              placeholder={strings.Directory_ColumnDepartment}
              ariaLabel={strings.Directory_FilterByDepartmentAria}
              selectedKey={selectedDepartment}
              options={deptOptions}
              onChange={(_, o) => o && this.setState({ selectedDepartment: o.key as string, currentPage: 1 })}
              className={styles.filterDropdown}
            />
          )}

          {officeOptions.length > 2 && (
            <Dropdown
              placeholder={strings.Directory_ColumnOffice}
              ariaLabel={strings.Directory_FilterByOfficeAria}
              selectedKey={selectedOffice}
              options={officeOptions}
              onChange={(_, o) => o && this.setState({ selectedOffice: o.key as string, currentPage: 1 })}
              className={styles.filterDropdown}
            />
          )}

          {anythingToClear && (
            <button className={styles.clearBtn} onClick={this._clearFilters} title={strings.Directory_ClearSearchTitle}>
              <Icon iconName="Cancel" aria-hidden="true" /> {strings.Directory_ClearButton}
            </button>
          )}

          <button
            className={styles.exportExcelBtn}
            onClick={() => this.exportExcel()}
            disabled={filtered.length === 0}
            title={strings.Directory_ExportCsvTitle}
          >
            <Icon iconName="Download" aria-hidden="true" />
            <span>{strings.Directory_ExportCsvButton}</span>
          </button>

          <div className={styles.viewToggle} role="group" aria-label={strings.Directory_ViewModeAria}>
            <button
              className={`${styles.viewBtn} ${viewMode === 'card' ? styles.viewBtnActive : ''}`}
              onClick={() => this._setViewMode('card')}
              title={strings.Directory_CardViewTitle}
              aria-label={strings.Directory_CardViewTitle}
              aria-pressed={viewMode === 'card'}
            >
              <Icon iconName="GridViewMedium" aria-hidden="true" />
            </button>
            <button
              className={`${styles.viewBtn} ${viewMode === 'list' ? styles.viewBtnActive : ''}`}
              onClick={() => this._setViewMode('list')}
              title={strings.Directory_ListViewTitle}
              aria-label={strings.Directory_ListViewTitle}
              aria-pressed={viewMode === 'list'}
            >
              <Icon iconName="BulletedList" aria-hidden="true" />
            </button>
          </div>
        </div>

        {/* ── Alphabet bar ── */}
        <div className={styles.alphabetBar} role="toolbar" aria-label={strings.Directory_AlphabetFilterAria}>
          {ALPHABET.map(letter => {
            const isActive = selectedLetter === letter && !hasSearch;
            return (
              <button
                key={letter}
                className={`${styles.letterBtn} ${isActive ? styles.active : ''}`}
                onClick={() => this._selectLetter(letter)}
                aria-pressed={isActive}
                title={letter === 'All' ? strings.Directory_ShowAllTitle : letter === OTHER_BUCKET ? strings.Directory_OtherLettersTitle : formatString(strings.Directory_FilterByLetterTitle, { letter })}
              >
                {letter}
              </button>
            );
          })}
        </div>

        {/* ── Result count + top pagination ── */}
        <div className={styles.resultMeta}>
          <span aria-live="polite">
            {filtered.length} {filtered.length === 1 ? strings.Directory_ResultCountSingular : strings.Directory_ResultCountPlural}
            {totalPages > 1 && formatString(strings.Directory_PageOf, { page: safePage, totalPages })}
            {activeFilters > 0 && (
              <span className={styles.filterBadge}>
                {formatString(activeFilters > 1 ? strings.Directory_FilterActivePlural : strings.Directory_FilterActiveSingular, { count: activeFilters })}
              </span>
            )}
          </span>
          {totalPages > 1 && (
            <div className={styles.paginationInline}>
              {this._renderPager(safePage, totalPages)}
            </div>
          )}
        </div>

        {/* ── Content ── */}
        {viewMode === 'card' ? this._renderCardGrid(paged) : this._renderListView(paged)}

        {filtered.length === 0 && (
          <div className={styles.noResults}>
            <Icon iconName="SearchIssue" />
            <span>{strings.Directory_NoResults}</span>
            {anythingToClear && (
              <button className={styles.clearBtn} onClick={this._clearFilters}>{strings.Directory_ClearFiltersButton}</button>
            )}
          </div>
        )}

        {/* ── Bottom pagination ── */}
        {totalPages > 1 && (
          <nav className={styles.pagination} aria-label={strings.Directory_PaginationAria}>
            {this._renderPager(safePage, totalPages)}
          </nav>
        )}
      </div>
    );
  }
}

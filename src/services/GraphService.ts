import { SPHttpClient, MSGraphClientV3 } from '@microsoft/sp-http';
import { createHttpError, getErrorStatus, getRetryAfterMs, parseRetryAfter, sleep, withRetry } from './requestUtils';
import { cacheDelete, cacheGet, cacheSet } from './persistentCache';

export interface IGraphUser {
  id: string;               // UPN (from AccountName) — matches what Manager field stores
  displayName: string;
  mail: string;
  jobTitle: string;
  mobilePhone: string;
  businessPhones: string[];
  department: string;
  officeLocation: string;
  userPrincipalName: string;
  accountEnabled?: boolean;  // false = disabled / blocked sign-in
  userType?: string;         // 'Member' | 'Guest' | undefined
  dottedManagerId?: string;  // secondary "dotted line" manager (resolved user id)
  // Admin-configured Entra ID attributes (see ICustomAttributeConfig), keyed by
  // graphField. Only populated for fields the admin configured; Search mode
  // never populates this (SharePoint Search has no generic way to read them).
  customAttributes?: { [graphField: string]: string };
}

export type PresenceAvailability =
  'Available' | 'Busy' | 'DoNotDisturb' | 'BeRightBack' | 'Away' | 'Offline' | 'Unknown';

export interface IOrgNode {
  user: IGraphUser;
  directReports: IOrgNode[];
  isExpanded: boolean;
  childrenLoaded: boolean;
  level: number;
  // Everyone below this person in the (filtered) org, at every level — not
  // just node.directReports.length, which is only the immediate children and
  // only once loaded. Computed from the whole in-memory manager graph, so it's
  // correct even for a node whose own children haven't been fetched yet.
  totalReportCount: number;
}

// An admin-configured Entra ID attribute to surface in the UI. `graphField`
// must be one of CUSTOM_ATTRIBUTE_ALLOWED_FIELDS or extensionAttribute1-15
// (see isValidCustomAttributeField) — anything else is ignored so it can't be
// used to smuggle an arbitrary value into a Graph $select.
export interface ICustomAttributeConfig {
  id: string;               // stable key for editing/React lists — not the Graph field name
  graphField: string;
  label: string;
  showInDirectory: boolean;
  showInOrgChart: boolean;
}

// Standard Graph /users properties safe to add to $select on demand. Deliberately
// excludes anything already always-selected, and anything that isn't a simple
// user-facing string (no navigation properties, no arrays other than businessPhones).
export const CUSTOM_ATTRIBUTE_ALLOWED_FIELDS: string[] = [
  'employeeId', 'employeeType', 'companyName', 'city', 'state', 'country',
  'postalCode', 'streetAddress', 'preferredLanguage', 'usageLocation',
  'faxNumber', 'employeeHireDate', 'onPremisesSamAccountName',
];

const EXTENSION_ATTR_RE = /^extensionAttribute(?:[1-9]|1[0-5])$/i;

export function isExtensionAttributeField(field: string): boolean {
  return EXTENSION_ATTR_RE.test(field);
}

export function isValidCustomAttributeField(field: string): boolean {
  return EXTENSION_ATTR_RE.test(field) || CUSTOM_ATTRIBUTE_ALLOWED_FIELDS.indexOf(field) >= 0;
}

// Pulls the configured fields off a raw Graph user object (flat properties
// directly, extensionAttributeN via onPremisesExtensionAttributes). Shared by
// the main listing fetch and the unlicensed-reports supplement.
function extractCustomAttributes(
  item: { [key: string]: unknown; onPremisesExtensionAttributes?: { [k: string]: string | null } },
  fields: string[]
): { [field: string]: string } | undefined {
  if (fields.length === 0) return undefined;
  const out: { [field: string]: string } = {};
  const ext = item.onPremisesExtensionAttributes;
  for (const field of fields) {
    if (EXTENSION_ATTR_RE.test(field)) {
      if (!ext) continue;
      const extKey = Object.keys(ext).filter(k => k.toLowerCase() === field.toLowerCase())[0];
      const v = extKey ? ext[extKey] : undefined;
      if (v) out[field] = v;
    } else {
      const v = item[field];
      if (typeof v === 'string' && v) out[field] = v;
    }
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

// Everything downloaded from the network, before the admin user filters are
// applied. All derived state (visible list, children/manager maps, lookups)
// is rebuilt from this, so filter changes and cache restores need no network.
export interface IRawUserData {
  users:      IGraphUser[];           // all fetched users, sorted by display name
  managerIds: Map<string, string>;    // userId → raw managerId (pre-resolution)
  dottedIds:  Map<string, string>;    // userId → raw dotted-line managerId
  objectIds:  Map<string, string>;    // userId → AAD object id (for presence)
  noPhotoIds: Set<string>;            // users known to have no profile photo
  fetchedAt:  number;                 // when this data came from the network
}

// Serialized form stored in IndexedDB (plain arrays rather than Map/Set)
interface IPersistedUserData {
  schema:     number;
  fetchedAt:  number;
  users:      IGraphUser[];
  managerIds: Array<[string, string]>;
  dottedIds:  Array<[string, string]>;
  objectIds:  Array<[string, string]>;
  noPhotoIds: string[];
}

interface IPresenceEntry {
  availability: PresenceAvailability;
  fetchedAt:    number;
}

interface IGraphDirectoryUser {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  jobTitle: string;
  department: string;
  officeLocation: string;
  mobilePhone: string;
  businessPhones: string[];
  accountEnabled?: boolean;
  userType?: string;
  manager?: { id: string; userPrincipalName: string; mail: string };
  onPremisesExtensionAttributes?: { [key: string]: string | null };
  // Admin-configured custom attributes (see CUSTOM_ATTRIBUTE_ALLOWED_FIELDS)
  // are requested dynamically, so they aren't declared here individually.
  [customField: string]: unknown;
}

interface IBatchResponse {
  id: string;
  status: number;
  headers?: { [key: string]: string };
  body?: { value?: IGraphDirectoryUser[]; '@odata.nextLink'?: string };
}

const PEOPLE_SOURCE = 'b09a7990-05ea-4af9-81ef-edfab16c4e31';
const SELECT_PROPS   = 'AccountName,DisplayName,PreferredName,JobTitle,Department,' +
                       'WorkEmail,WorkPhone,MobilePhone,OfficeNumber,PictureURL,Manager';
const BATCH_SIZE = 500;
const MAX_SEARCH_PAGES = 200;                // 100k users — guards against a paging loop
const GRAPH_BATCH_LIMIT = 20;                // Graph JSON $batch request limit
const REPORT_SELECT_BASE = 'id,displayName,mail,jobTitle,department,officeLocation,mobilePhone,' +
                           'businessPhones,userPrincipalName,accountEnabled,userType';

const DATA_TTL = 4 * 60 * 60_000;            // user data (memory + IndexedDB) lifetime
const STALE_RETRY = 5 * 60_000;              // after a failed background refresh, keep stale data this long
const CACHE_SCHEMA = 1;

export type DataSource = 'auto' | 'graph' | 'search';

export interface IUserFilterOptions {
  tenantDomain?: string;       // only show users whose email domain matches (e.g. 'contoso.com')
  excludedPatterns?: string[]; // lower-case substrings — hide any user whose name/UPN/mail contains one
  hideGuestUsers?: boolean;    // hide userType === 'Guest'
  hideDisabledAccounts?: boolean; // hide accountEnabled === false
  hideNoJobTitle?: boolean;    // hide users with no jobTitle
  hideNoDepartment?: boolean;  // hide users with no department
}

// Copy without dottedManagerId, which _buildMaps sets per filter pass
function cloneUser(u: IGraphUser): IGraphUser {
  const copy: IGraphUser = { ...u };
  delete copy.dottedManagerId;
  return copy;
}

function emptyRawData(): IRawUserData {
  return {
    users: [], managerIds: new Map(), dottedIds: new Map(), objectIds: new Map(),
    noPhotoIds: new Set(), fetchedAt: Date.now(),
  };
}

function mapToPairs(m: Map<string, string>): Array<[string, string]> {
  const out: Array<[string, string]> = [];
  m.forEach((v, k) => out.push([k, v]));
  return out;
}

export class GraphService {
  private _client:         SPHttpClient | null;
  private _graphClient:    MSGraphClientV3 | undefined;
  private _webUrl:         string;
  private _dataSource:     DataSource;
  private _filterOptions:  IUserFilterOptions;
  private _dottedLineAttribute: string;                          // AAD extension attribute holding the dotted-line manager
  private _customAttributeFields: string[];                      // admin-configured extra Graph fields to fetch/expose

  // Raw (unfiltered) network data — the single source for all derived state
  private _raw: IRawUserData | null = null;
  private _loadingPromise: Promise<IGraphUser[]> | null = null;
  private _generation = 0;                                       // bumped by refresh() to discard in-flight loads
  private _staleRetryAt = 0;                                     // don't retry an expired-data reload before this

  // Derived state — rebuilt by _buildMaps from _raw + _filterOptions
  private _allUsersCache:  IGraphUser[] | null = null;
  private _childrenMap:    Map<string, IGraphUser[]> = new Map();
  private _managerMap:     Map<string, string> = new Map();      // userId → managerId (resolved, visible)
  protected _dottedReportsMap: Map<string, IGraphUser[]> = new Map(); // managerId → dotted-line reports
  private _byId:   Map<string, IGraphUser> = new Map();
  private _byMail: Map<string, IGraphUser> = new Map();
  private _byUpn:  Map<string, IGraphUser> = new Map();
  private _photoCache:     Map<string, string | null> = new Map();
  private _totalCountCache: Map<string, number> = new Map();     // userId → total descendants; rebuilt with _childrenMap
  private _cycleWarned = false;

  // Presence
  private _upnToObjectId:   Map<string, string> = new Map();     // upn → AAD object id
  private _unresolvableUpns: Set<string> = new Set();            // UPNs Graph returned no user for
  private _objectIdPromise: Promise<void> | null = null;         // shared in-flight object id lookup
  private _presenceCache:   Map<string, IPresenceEntry> = new Map();
  private _presenceFailedUntil = 0;                              // back-off after a presence failure
  private readonly _PRESENCE_TTL = 60_000;
  private readonly _PRESENCE_RETRY = 15 * 60_000;                // missing permission
  private readonly _PRESENCE_TRANSIENT_RETRY = 60_000;           // network / server errors

  constructor(
    client: SPHttpClient | null,
    webUrl: string,
    graphClient?: MSGraphClientV3,
    dataSource: DataSource = 'auto',
    filterOptions: IUserFilterOptions = {},
    dottedLineAttribute = '',
    customAttributeFields: string[] = []
  ) {
    this._client        = client;
    this._webUrl        = webUrl.replace(/\/$/, '');
    this._graphClient   = graphClient;
    this._dataSource    = dataSource;
    this._filterOptions = filterOptions;
    this._dottedLineAttribute = dottedLineAttribute.trim();
    this._customAttributeFields = (customAttributeFields || []).filter(isValidCustomAttributeField);
  }

  /* ── Public API ──────────────────────────────────────────────────── */

  public getAllUsers(): Promise<IGraphUser[]> {
    if (this._allUsersCache && !this._isStale()) return Promise.resolve(this._allUsersCache.slice());
    return this._load().then(users => users.slice());
  }

  // Replaces the admin user filters and re-derives everything from the raw
  // download already in memory — no network call. Before the first load
  // completes, the options are simply used by that load.
  public updateFilterOptions(options: IUserFilterOptions): void {
    this._filterOptions = options || {};
    if (this._raw) this._buildMaps();
  }

  // Drops the in-memory data and this service's persistent cache entry so the
  // next data call re-downloads from the network.
  public async refresh(): Promise<void> {
    this._generation++;
    this._raw = null;
    this._loadingPromise = null;
    this._staleRetryAt = 0;
    this._allUsersCache = null;
    this._childrenMap = new Map();
    this._managerMap = new Map();
    this._dottedReportsMap = new Map();
    this._byId = new Map();
    this._byMail = new Map();
    this._byUpn = new Map();
    this._photoCache = new Map();
    this._totalCountCache = new Map();
    this._upnToObjectId = new Map();
    this._unresolvableUpns = new Set();
    this._presenceCache = new Map();
    const key = this._cacheKey();
    if (key) await cacheDelete(key);
  }

  // When the current user data was fetched from the network (for data
  // restored from the persistent cache, the original fetch time).
  public getLastLoaded(): Date | null {
    return this._raw ? new Date(this._raw.fetchedAt) : null;
  }

  // Returns a photo URL, or null when the user is known to have no photo.
  // Graph mode can't tell cheaply, so it always returns the SharePoint
  // userphoto.aspx URL (which serves a silhouette when there is no photo) —
  // callers should fall back to initials if the image fails to load.
  public async getUserPhoto(userId: string): Promise<string | null> {
    const key = userId.toLowerCase();
    if (this._photoCache.has(key)) return this._photoCache.get(key) ?? null;
    await this._ensureLoaded(); // photo cache is populated during user load
    return this._photoCache.get(key) ?? null;
  }

  public async getDirectReports(userId: string): Promise<IGraphUser[]> {
    await this._ensureLoaded();
    return (this._childrenMap.get(userId.toLowerCase()) || []).slice();
  }

  public async getDottedLineReports(userId: string): Promise<IGraphUser[]> {
    await this._ensureLoaded();
    return (this._dottedReportsMap.get(userId.toLowerCase()) || []).slice();
  }

  public async hasDirectReports(userId: string): Promise<boolean> {
    await this._ensureLoaded();
    const kids = this._childrenMap.get(userId.toLowerCase());
    return !!(kids && kids.length > 0);
  }

  // Immediate reports only. A synchronous counterpart to getDirectReports()
  // for callers that just need the number — safe to call anywhere the data is
  // already known to be loaded (e.g. while rendering a tree built from it).
  public getDirectReportCount(userId: string): number {
    return (this._childrenMap.get(userId.toLowerCase()) || []).length;
  }

  // Everyone below this person at any depth, from the whole in-memory manager
  // graph — independent of how much of the tree has actually been expanded.
  // Memoized; the cache is cleared whenever _childrenMap is rebuilt.
  public getTotalReportCount(userId: string): number {
    return this._computeTotalReportCount(userId.toLowerCase(), new Set<string>());
  }

  private _computeTotalReportCount(id: string, guard: Set<string>): number {
    const cached = this._totalCountCache.get(id);
    if (cached !== undefined) return cached;
    if (guard.has(id)) return 0; // defense in depth — _breakManagerCycles should prevent this
    guard.add(id);
    const kids = this._childrenMap.get(id) || [];
    let total = kids.length;
    for (const kid of kids) total += this._computeTotalReportCount(kid.id.toLowerCase(), guard);
    guard.delete(id);
    this._totalCountCache.set(id, total);
    return total;
  }

  public async getManagerChain(userId: string, levels: number): Promise<IGraphUser[]> {
    await this._ensureLoaded();
    const chain: IGraphUser[] = [];
    let curId = userId.toLowerCase();
    const visited = new Set<string>([curId]);   // defense in depth — cycles are broken in _buildMaps
    for (let i = 0; i < levels; i++) {
      const mgrId = this._managerMap.get(curId);
      if (!mgrId || visited.has(mgrId)) break;
      const mgrUser = this._byId.get(mgrId);
      if (!mgrUser) break;
      visited.add(mgrId);
      chain.unshift(mgrUser);
      curId = mgrId;
    }
    return chain;
  }

  public async findUser(identifier: string): Promise<IGraphUser | null> {
    if (!identifier) return null;
    await this._ensureLoaded();
    const q = identifier.toLowerCase();
    return (
      this._byId.get(q) ||
      this._byMail.get(q) ||
      this._byUpn.get(q) ||
      (this._allUsersCache || []).find(u => (u.displayName || '').toLowerCase().startsWith(q)) ||
      null
    );
  }

  // Fetches presence for the given user IDs only (pass the users currently on
  // screen — fetching the whole tenant every poll does not scale). Omitting
  // userIds falls back to all visible users. Returns a snapshot copy.
  public async getPresence(userIds?: string[]): Promise<Map<string, PresenceAvailability>> {
    const graph = this._graphClient;
    if (!graph) return new Map();
    if (Date.now() < this._presenceFailedUntil) return this._presenceSnapshot();

    // Object ids come with the Graph user download — let it finish first
    if (this._loadingPromise) {
      try { await this._loadingPromise; } catch { /* surfaced by the data loaders */ }
    }

    const wanted = userIds
      ? Array.from(new Set(userIds.map(id => id.toLowerCase())))
      : (this._allUsersCache || []).map(u => u.id);

    // Each entry has its own fetch time — only refetch missing or stale ones
    const now = Date.now();
    const toFetch = wanted.filter(upn => {
      const entry = this._presenceCache.get(upn);
      return !entry || now - entry.fetchedAt >= this._PRESENCE_TTL;
    });
    if (toFetch.length === 0) return this._presenceSnapshot();

    try {
      await this._ensureObjectIds(toFetch);
    } catch (err) {
      this._presenceFailedUntil = Date.now() + this._presenceBackoff(err);
      return this._presenceSnapshot();
    }

    // Build reverse map so we can key results back by UPN
    const objectIds: string[] = [];
    const reverseMap = new Map<string, string>();
    for (const upn of toFetch) {
      const objId = this._upnToObjectId.get(upn);
      if (objId) { objectIds.push(objId); reverseMap.set(objId, upn); }
    }
    if (objectIds.length === 0) return this._presenceSnapshot();

    const CHUNK = 650;
    for (let i = 0; i < objectIds.length; i += CHUNK) {
      try {
        const response = await graph
          .api('/communications/getPresencesByUserId')
          .version('v1.0')
          .post({ ids: objectIds.slice(i, i + CHUNK) });
        const presences: Array<{ id: string; availability: string }> = response?.value || [];
        this._presenceFailedUntil = 0;
        const fetchedAt = Date.now();
        for (const p of presences) {
          const upn = reverseMap.get(p.id);
          if (upn) this._presenceCache.set(upn, { availability: this._normalizeAvailability(p.availability), fetchedAt });
        }
      } catch (err) {
        // Back off so the 60-second polls don't hammer Graph — long for a
        // missing Presence.Read.All permission, short for throttling/outages
        this._presenceFailedUntil = Date.now() + this._presenceBackoff(err);
        break;
      }
    }

    return this._presenceSnapshot();
  }

  public async buildOrgTree(rootUserId: string, levelsBelow: number): Promise<IOrgNode> {
    await this._ensureLoaded();
    const q = rootUserId.toLowerCase();
    const user = this._byId.get(q) || this._byMail.get(q);
    if (!user) throw new Error(`User not found: ${rootUserId}`);

    const root: IOrgNode = {
      user, directReports: [], isExpanded: true, childrenLoaded: false, level: 0,
      totalReportCount: this.getTotalReportCount(user.id)
    };
    this._loadChildren(root, levelsBelow, new Set<string>([user.id]));
    return root;
  }

  /* ── Loading & caching ───────────────────────────────────────────── */

  // Persistent cache key — null disables the persistent cache (demo data)
  protected _cacheKey(): string | null {
    // Changing which custom attributes are configured changes what's fetched,
    // so it must change the key too — otherwise a newly-added attribute would
    // silently stay blank until the cache's normal TTL expires.
    const attrs = this._customAttributeFields.slice().sort().join(',');
    return `users|${this._webUrl.toLowerCase()}|${this._dataSource}|${this._dottedLineAttribute.toLowerCase()}|${attrs}`;
  }

  private _isStale(): boolean {
    if (!this._raw) return true;
    const age = Date.now() - this._raw.fetchedAt;
    return (age >= DATA_TTL || age < 0) && Date.now() >= this._staleRetryAt;
  }

  private async _ensureLoaded(): Promise<void> {
    if (!this._allUsersCache || this._isStale()) await this._load();
  }

  private _load(): Promise<IGraphUser[]> {
    if (this._loadingPromise) return this._loadingPromise;
    const gen = this._generation;

    const p: Promise<IGraphUser[]> = this._acquireRawData()
      .then(({ raw, fromNetwork }) => {
        if (this._loadingPromise === p) this._loadingPromise = null;
        // refresh() ran mid-load — discard this result and load again
        if (gen !== this._generation) return this._load();
        this._raw = raw;
        this._staleRetryAt = 0;
        this._buildMaps();
        if (fromNetwork) this._persist(raw);
        return this._allUsersCache as IGraphUser[];
      }, err => {
        if (this._loadingPromise === p) this._loadingPromise = null;
        // A background refresh of expired data failed — keep serving the old
        // data for a while rather than breaking an already-rendered page
        if (gen === this._generation && this._allUsersCache) {
          console.warn('[SmartOrgChart] Refreshing user data failed; keeping cached data:', err);
          this._staleRetryAt = Date.now() + STALE_RETRY;
          return this._allUsersCache;
        }
        throw err;
      });

    this._loadingPromise = p;
    return p;
  }

  private async _acquireRawData(): Promise<{ raw: IRawUserData; fromNetwork: boolean }> {
    const key = this._cacheKey();
    if (key) {
      const cached = await this._readPersisted(key);
      if (cached) return { raw: cached, fromNetwork: false };
    }
    return { raw: await this._fetchRawData(), fromNetwork: true };
  }

  private async _readPersisted(key: string): Promise<IRawUserData | null> {
    try {
      const rec = await cacheGet<IPersistedUserData>(key);
      if (!rec || rec.schema !== CACHE_SCHEMA || !Array.isArray(rec.users)) return null;
      const age = Date.now() - rec.fetchedAt;
      if (!(age >= 0 && age < DATA_TTL)) {
        cacheDelete(key).catch(() => undefined);
        return null;
      }
      return {
        users:      rec.users,
        managerIds: new Map(rec.managerIds || []),
        dottedIds:  new Map(rec.dottedIds || []),
        objectIds:  new Map(rec.objectIds || []),
        noPhotoIds: new Set(rec.noPhotoIds || []),
        fetchedAt:  rec.fetchedAt,
      };
    } catch {
      return null;
    }
  }

  // Fire-and-forget — caching is an optimization and must never fail a load.
  // Photos and presence are deliberately not persisted.
  private _persist(raw: IRawUserData): void {
    const key = this._cacheKey();
    if (!key) return;
    try {
      const noPhotoIds: string[] = [];
      raw.noPhotoIds.forEach(id => noPhotoIds.push(id));
      const rec: IPersistedUserData = {
        schema:     CACHE_SCHEMA,
        fetchedAt:  raw.fetchedAt,
        users:      raw.users.map(cloneUser),
        managerIds: mapToPairs(raw.managerIds),
        dottedIds:  mapToPairs(raw.dottedIds),
        objectIds:  mapToPairs(raw.objectIds),
        noPhotoIds,
      };
      cacheSet(key, rec).catch(() => undefined);
    } catch {
      // unserializable or storage unavailable — skip caching
    }
  }

  // Downloads the raw directory. Overridden by MockGraphService.
  protected async _fetchRawData(): Promise<IRawUserData> {
    const useGraph = this._dataSource === 'graph' ||
                     (this._dataSource === 'auto' && !!this._graphClient);

    if (this._dataSource === 'graph' && !this._graphClient) {
      throw new Error(
        'Graph API data source selected but Microsoft Graph permissions have not been granted. ' +
        'Please approve the app permissions in the SharePoint App Catalog, or switch the Data Source setting to "SharePoint Search".'
      );
    }

    let graphFailed = false;
    if (useGraph) {
      try {
        const data = await this._fetchUsersFromGraph();
        if (data.users.length === 0) throw new Error('Graph API returned 0 users');
        return data;
      } catch (err) {
        if (this._dataSource === 'graph') throw err;
        // 'auto' mode: fall back to SharePoint Search. Nothing from the failed
        // Graph attempt was committed, so there is no partial state to undo.
        console.warn('[SmartOrgChart] Graph API unavailable, falling back to SharePoint Search:', err);
        graphFailed = true;
      }
    }

    // SharePoint Search path (primary or fallback)
    const data = await this._fetchUsersFromSearch();
    // Graph just failed — don't make hundreds of directReports calls to it too
    if (!graphFailed) await this._supplementUnlicensedReports(data);
    return data;
  }

  /* ── Filters & derived maps ──────────────────────────────────────── */

  protected _applyUserFilters(users: IGraphUser[]): IGraphUser[] {
    const { excludedPatterns, hideGuestUsers, hideDisabledAccounts,
            hideNoJobTitle, hideNoDepartment } = this._filterOptions;
    const tenantDomain = (this._filterOptions.tenantDomain || '').toLowerCase();
    const hasFilters = tenantDomain || (excludedPatterns && excludedPatterns.length > 0) ||
                       hideGuestUsers || hideDisabledAccounts || hideNoJobTitle || hideNoDepartment;
    if (!hasFilters) return users;

    return users.filter(user => {
      if (hideDisabledAccounts && user.accountEnabled === false) return false;
      if (hideGuestUsers && user.userType === 'Guest') return false;
      if (hideNoJobTitle && !user.jobTitle) return false;
      if (hideNoDepartment && !user.department) return false;
      if (tenantDomain) {
        // Exact domain or any subdomain of it (e.g. uk.contoso.com)
        const emailDomain = ((user.mail || user.id || '').split('@')[1] || '').toLowerCase();
        if (emailDomain && emailDomain !== tenantDomain && !emailDomain.endsWith('.' + tenantDomain)) return false;
      }
      if (excludedPatterns && excludedPatterns.length > 0) {
        const haystack = [
          user.displayName || '',
          user.id || '',
          user.mail || '',
          user.userPrincipalName || '',
        ].join('\0').toLowerCase();
        if (excludedPatterns.some(p => haystack.includes(p))) return false;
      }
      return true;
    });
  }

  // Rebuilds every derived structure from _raw and the current filters.
  // Operates on copies of the raw users, so re-filtering is side-effect free.
  private _buildMaps(): void {
    const raw = this._raw;
    if (!raw) return;

    // Resolve manager references against ALL fetched users (before the admin
    // user filters), so a report whose manager is hidden by a filter can be
    // bridged to the nearest visible ancestor instead of orphaning the branch.
    const allUsers = raw.users;
    const users = this._applyUserFilters(allUsers).map(cloneUser);

    // Display name is a last-resort manager lookup (SP Search sometimes stores
    // only the manager's name). Skip names shared by multiple users — guessing
    // would silently attach reports to the wrong person.
    const nameCount = new Map<string, number>();
    for (const user of allUsers) {
      const dn = (user.displayName || '').toLowerCase();
      nameCount.set(dn, (nameCount.get(dn) || 0) + 1);
    }

    const canonicalId = new Map<string, string>();
    for (const user of allUsers) {
      canonicalId.set(user.id, user.id);
      if (user.mail && !canonicalId.has(user.mail.toLowerCase())) canonicalId.set(user.mail.toLowerCase(), user.id);
      const dn = (user.displayName || '').toLowerCase();
      if (nameCount.get(dn) === 1 && !canonicalId.has(dn)) canonicalId.set(dn, user.id);
    }

    // Raw manager edges across all fetched users. Self-managed accounts
    // (common for CEOs in Azure AD) are dropped here — without this guard a
    // user renders as their own child and Expand All never terminates.
    const rawManagerOf = new Map<string, string>();
    for (const user of allUsers) {
      const rawMgr = raw.managerIds.get(user.id);
      if (!rawMgr) continue;
      const mgrId = canonicalId.get(rawMgr) || rawMgr;
      if (mgrId !== user.id) rawManagerOf.set(user.id, mgrId);
    }

    const byId = new Map<string, IGraphUser>();
    const byMail = new Map<string, IGraphUser>();
    const byUpn = new Map<string, IGraphUser>();
    for (const user of users) {
      byId.set(user.id, user);
      const mail = (user.mail || '').toLowerCase();
      if (mail && !byMail.has(mail)) byMail.set(mail, user);
      const upn = (user.userPrincipalName || '').toLowerCase();
      if (upn && !byUpn.has(upn)) byUpn.set(upn, user);
    }

    const managerMap = new Map<string, string>();
    for (const user of users) {
      // Walk up through hidden managers to the nearest visible one; the
      // visited set stops manager cycles (A→B→A) from looping forever
      let mgrId = rawManagerOf.get(user.id);
      const visited = new Set<string>([user.id]);
      while (mgrId && !byId.has(mgrId) && !visited.has(mgrId)) {
        visited.add(mgrId);
        mgrId = rawManagerOf.get(mgrId);
      }
      if (!mgrId || !byId.has(mgrId) || mgrId === user.id) continue;
      managerMap.set(user.id, mgrId);
    }
    this._breakManagerCycles(managerMap, byId);

    // Children in display-name order (users is sorted)
    const childrenMap = new Map<string, IGraphUser[]>();
    for (const user of users) {
      const mgrId = managerMap.get(user.id);
      if (!mgrId) continue;
      if (!childrenMap.has(mgrId)) childrenMap.set(mgrId, []);
      (childrenMap.get(mgrId) as IGraphUser[]).push(user);
    }

    const dottedReportsMap = new Map<string, IGraphUser[]>();
    for (const user of users) {
      const rawDotted = raw.dottedIds.get(user.id);
      if (!rawDotted) continue;
      const dottedId = canonicalId.get(rawDotted);
      if (!dottedId || dottedId === user.id || !byId.has(dottedId)) continue; // unresolvable, hidden, or self-reference
      user.dottedManagerId = dottedId;
      if (!dottedReportsMap.has(dottedId)) dottedReportsMap.set(dottedId, []);
      (dottedReportsMap.get(dottedId) as IGraphUser[]).push(user);
    }

    // Photo URLs are derived (cheap strings) — never persisted
    const photoCache = new Map<string, string | null>();
    for (const user of allUsers) {
      const photoEmail = (user.mail || user.userPrincipalName || '').toLowerCase();
      photoCache.set(user.id, photoEmail && !raw.noPhotoIds.has(user.id)
        ? `${this._webUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(photoEmail)}`
        : null);
    }

    raw.objectIds.forEach((objId, upn) => this._upnToObjectId.set(upn, objId));

    this._allUsersCache    = users;
    this._byId             = byId;
    this._byMail           = byMail;
    this._byUpn            = byUpn;
    this._managerMap       = managerMap;
    this._childrenMap      = childrenMap;
    this._dottedReportsMap = dottedReportsMap;
    this._photoCache       = photoCache;
    this._totalCountCache  = new Map();   // stale once childrenMap changes
  }

  // Manager data can contain cycles between visible users (A→B→C→A), which
  // would leave the whole loop without a root and make trees recurse forever.
  // Break each cycle by making its alphabetically-first member a root.
  private _breakManagerCycles(managerMap: Map<string, string>, byId: Map<string, IGraphUser>): void {
    const state = new Map<string, number>();   // 1 = on current path, 2 = done
    const cycles: string[][] = [];
    managerMap.forEach((_mgr, startId) => {
      if (state.has(startId)) return;
      const path: string[] = [];
      let cur: string | undefined = startId;
      while (cur && !state.has(cur)) {
        state.set(cur, 1);
        path.push(cur);
        cur = managerMap.get(cur);
      }
      if (cur && state.get(cur) === 1) cycles.push(path.slice(path.indexOf(cur)));
      for (const id of path) state.set(id, 2);
    });
    if (cycles.length === 0) return;

    const nameOf = (id: string): string => ((byId.get(id) || { displayName: id }).displayName || id).toLowerCase();
    for (const cycle of cycles) {
      const root = cycle.slice().sort((a, b) => nameOf(a).localeCompare(nameOf(b)))[0];
      managerMap.delete(root);
    }
    if (!this._cycleWarned) {
      this._cycleWarned = true;
      console.warn(`[SmartOrgChart] Found ${cycles.length} manager cycle(s) in directory data; ` +
        'broke each by treating one member as a top-level user:', cycles);
    }
  }

  private _loadChildren(node: IOrgNode, remaining: number, ancestors: Set<string>): void {
    // ancestors guards against cycles (defense in depth — _buildMaps breaks them)
    const reports = (this._childrenMap.get(node.user.id) || []).filter(u => !ancestors.has(u.id));
    node.childrenLoaded = true;
    node.directReports = reports.map(u => {
      const hasKids = (this._childrenMap.get(u.id) || []).length > 0;
      return {
        user: u,
        directReports: [],
        isExpanded: remaining > 1,
        childrenLoaded: remaining <= 1 ? !hasKids : false,
        level: node.level + 1,
        totalReportCount: this.getTotalReportCount(u.id)
      };
    });
    if (remaining > 1) {
      node.directReports.forEach(child => {
        ancestors.add(child.user.id);
        this._loadChildren(child, remaining - 1, ancestors);
        ancestors.delete(child.user.id);
      });
    }
  }

  /* ── Presence helpers ────────────────────────────────────────────── */

  private _presenceSnapshot(): Map<string, PresenceAvailability> {
    const out = new Map<string, PresenceAvailability>();
    this._presenceCache.forEach((entry, upn) => out.set(upn, entry.availability));
    return out;
  }

  private _presenceBackoff(err: unknown): number {
    const status = getErrorStatus(err);
    if (status === 401 || status === 403) return this._PRESENCE_RETRY;   // permission not granted
    if (status === 429) {
      const retryAfter = getRetryAfterMs(err);
      return Math.min(5 * 60_000, Math.max(5_000, retryAfter !== undefined ? retryAfter : 30_000));
    }
    return this._PRESENCE_TRANSIENT_RETRY;
  }

  // Resolves AAD object ids for just the given UPNs. Concurrent callers
  // (directory + org chart) share one in-flight lookup.
  private async _ensureObjectIds(upns: string[]): Promise<void> {
    while (this._objectIdPromise) {
      try { await this._objectIdPromise; } catch { /* the caller that started it handles failures */ }
    }
    const missing = upns.filter(u => !this._upnToObjectId.has(u) && !this._unresolvableUpns.has(u));
    if (missing.length === 0) return;

    const p = this._fetchObjectIds(missing);
    this._objectIdPromise = p;
    try {
      await p;
    } finally {
      if (this._objectIdPromise === p) this._objectIdPromise = null;
    }
  }

  private async _fetchObjectIds(upns: string[]): Promise<void> {
    const graph = this._graphClient;
    if (!graph) return;
    const CHUNK = 15;   // Graph limits the number of values in an 'in' filter
    for (let i = 0; i < upns.length; i += CHUNK) {
      const chunk = upns.slice(i, i + CHUNK);
      const list = chunk.map(u => `'${u.replace(/'/g, "''")}'`).join(',');
      const url = `/users?$filter=${encodeURIComponent(`userPrincipalName in (${list})`)}` +
                  `&$select=id,userPrincipalName&$top=${CHUNK}`;
      const response = await graph.api(url).version('v1.0').get();
      const found: Array<{ id: string; userPrincipalName: string }> = response?.value || [];
      for (const u of found) {
        if (u.id && u.userPrincipalName) this._upnToObjectId.set(u.userPrincipalName.toLowerCase(), u.id);
      }
      // Remember misses (e.g. ids that aren't UPNs) so polls don't re-query them
      for (const upn of chunk) {
        if (!this._upnToObjectId.has(upn)) this._unresolvableUpns.add(upn);
      }
    }
  }

  private _normalizeAvailability(raw: string): PresenceAvailability {
    const map: { [k: string]: PresenceAvailability } = {
      Available: 'Available', AvailableIdle: 'Available',
      Busy: 'Busy', BusyIdle: 'Busy',
      DoNotDisturb: 'DoNotDisturb',
      BeRightBack: 'BeRightBack',
      Away: 'Away',
      Offline: 'Offline',
      PresenceUnknown: 'Unknown',
    };
    return map[raw] || 'Unknown';
  }

  /* ── Network fetchers ────────────────────────────────────────────── */

  private _reportSelect(): string {
    const flatCustomFields = this._customAttributeFields.filter(f => !isExtensionAttributeField(f));
    const needsExtensionAttrs = this._customAttributeFields.some(isExtensionAttributeField);
    let select = REPORT_SELECT_BASE;
    if (flatCustomFields.length > 0) select += ',' + flatCustomFields.join(',');
    if (needsExtensionAttrs) select += ',onPremisesExtensionAttributes';
    return select;
  }

  // Queries Graph directReports for each SP-found manager (via JSON $batch)
  // and merges in any unlicensed users that SP Search omitted. Best-effort:
  // stops on the first permission error and never fails the load.
  private async _supplementUnlicensedReports(data: IRawUserData): Promise<void> {
    const graph = this._graphClient;
    if (!graph) return;

    const knownIds = new Set<string>();
    for (const u of data.users) {
      knownIds.add(u.id);
      if (u.mail) knownIds.add(u.mail.toLowerCase());
    }

    // Only UPN/email managers — display-name managers would just 404
    const managerIds: string[] = [];
    const seenMgr = new Set<string>();
    data.managerIds.forEach(m => {
      if (m.indexOf('@') >= 0 && !seenMgr.has(m)) { seenMgr.add(m); managerIds.push(m); }
    });
    if (managerIds.length === 0) return;

    let added = 0;
    const addReports = (managerId: string, reports: IGraphDirectoryUser[]): void => {
      for (const rep of reports) {
        const upn  = (rep.userPrincipalName || '').toLowerCase();
        const mail = (rep.mail || '').toLowerCase();
        const id   = upn || mail || rep.id?.toLowerCase() || '';
        if (!id || knownIds.has(id) || knownIds.has(upn) || knownIds.has(mail)) continue;

        knownIds.add(id);
        if (mail) knownIds.add(mail);
        data.noPhotoIds.add(id);
        data.managerIds.set(id, managerId);
        if (rep.id) data.objectIds.set(id, rep.id);
        added++;

        data.users.push({
          id,
          displayName:       rep.displayName    || '',
          mail:              rep.mail            || '',
          jobTitle:          rep.jobTitle        || '',
          mobilePhone:       rep.mobilePhone     || '',
          businessPhones:    rep.businessPhones  || [],
          department:        rep.department      || '',
          officeLocation:    rep.officeLocation  || '',
          userPrincipalName: upn,
          accountEnabled:    rep.accountEnabled,
          userType:          rep.userType,
          customAttributes:  extractCustomAttributes(rep, this._customAttributeFields),
        });
      }
    };

    let denied = false;
    try {
      for (let i = 0; i < managerIds.length && !denied; i += GRAPH_BATCH_LIMIT) {
        let pending = managerIds.slice(i, i + GRAPH_BATCH_LIMIT);
        // Individually throttled requests inside a batch are retried a couple of times
        for (let attempt = 0; pending.length > 0 && attempt < 3 && !denied; attempt++) {
          const batchIds = pending;
          const requests = batchIds.map((m, idx) => ({
            id: String(idx),
            method: 'GET',
            url: `/users/${encodeURIComponent(m)}/directReports?$select=${this._reportSelect()}&$top=999`,
          }));
          const resp = await withRetry(() => graph.api('/$batch').version('v1.0').post({ requests }));
          const responses: IBatchResponse[] = resp?.responses || [];

          const throttled: string[] = [];
          let waitMs = 0;
          for (const r of responses) {
            const managerId = batchIds[Number(r.id)];
            if (!managerId) continue;
            if (r.status === 401 || r.status === 403) { denied = true; break; }
            if (r.status === 429 || r.status >= 500) {
              throttled.push(managerId);
              const h = r.headers || {};
              waitMs = Math.max(waitMs, parseRetryAfter(h['Retry-After'] || h['retry-after']) || 0);
              continue;
            }
            if (r.status !== 200) continue;   // 404 — external user, deleted account, etc.

            addReports(managerId, r.body?.value || []);
            let next = r.body?.['@odata.nextLink'];
            while (next) {
              const nextUrl: string = next;
              const page = await withRetry(() => graph.api(nextUrl).get());
              addReports(managerId, page?.value || []);
              next = page?.['@odata.nextLink'];
            }
          }
          pending = throttled;
          if (pending.length > 0 && !denied) await sleep(Math.min(30_000, waitMs || 2000 * (attempt + 1)));
        }
      }
    } catch (err) {
      const status = getErrorStatus(err);
      if (status === 401 || status === 403) denied = true;
      else console.warn('[SmartOrgChart] Could not supplement unlicensed users from Graph:', err);
    }
    if (denied) console.warn('[SmartOrgChart] Graph directReports not permitted — skipping unlicensed-user supplement.');

    if (added > 0) data.users.sort((a, b) => (a.displayName || '').localeCompare(b.displayName || ''));
  }

  private async _fetchUsersFromGraph(): Promise<IRawUserData> {
    const graph = this._graphClient;
    if (!graph) throw new Error('Graph client not available');

    // Collected into locals and returned — nothing is committed to the
    // service unless every page succeeds
    const data = emptyRawData();
    const seen = new Set<string>();
    const attrKey = this._dottedLineAttribute.toLowerCase();
    const flatCustomFields = this._customAttributeFields.filter(f => !isExtensionAttributeField(f));
    const needsExtensionAttrs = !!attrKey || this._customAttributeFields.some(isExtensionAttributeField);
    let SELECT = 'id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,businessPhones,accountEnabled,userType';
    if (flatCustomFields.length > 0) SELECT += ',' + flatCustomFields.join(',');
    if (needsExtensionAttrs) SELECT += ',onPremisesExtensionAttributes';
    let url: string | null =
      `/users?$select=${SELECT}&$expand=manager($select=id,userPrincipalName,mail)&$top=999`;

    while (url) {
      // Retry a failed page (throttling / transient errors) before giving up
      const pageUrl: string = url;
      const response = await withRetry(() => {
        const req = graph.api(pageUrl);
        if (!pageUrl.startsWith('https://')) req.version('v1.0');
        return req.get();
      });
      const items: IGraphDirectoryUser[] = response?.value || [];

      for (const item of items) {
        const upn  = (item.userPrincipalName || '').toLowerCase();
        const mail = (item.mail || '').toLowerCase();
        const id   = upn || mail;
        if (!id || !item.displayName || seen.has(id)) continue;
        seen.add(id);

        // Store AAD object ID for presence lookups
        if (item.id) data.objectIds.set(id, item.id);

        // Manager relationship — prefer UPN, fall back to mail
        const mgrUpn  = (item.manager?.userPrincipalName || '').toLowerCase();
        const mgrMail = (item.manager?.mail || '').toLowerCase();
        const mgrId   = mgrUpn || mgrMail;
        if (mgrId) data.managerIds.set(id, mgrId);

        // Dotted-line manager from the configured extension attribute (email
        // or UPN); the attribute name is matched case-insensitively
        const ext = item.onPremisesExtensionAttributes;
        if (attrKey && ext) {
          const extKey = Object.keys(ext).filter(k => k.toLowerCase() === attrKey)[0];
          const dotted = (extKey ? ext[extKey] || '' : '').trim().toLowerCase();
          if (dotted) data.dottedIds.set(id, dotted);
        }

        data.users.push({
          id,
          displayName:       item.displayName,
          mail:              mail,
          jobTitle:          item.jobTitle          || '',
          mobilePhone:       item.mobilePhone        || '',
          businessPhones:    item.businessPhones     || [],
          department:        item.department         || '',
          officeLocation:    item.officeLocation     || '',
          userPrincipalName: upn,
          accountEnabled:    item.accountEnabled,
          userType:          item.userType,
          customAttributes:  extractCustomAttributes(item, this._customAttributeFields),
        });
      }

      url = response?.['@odata.nextLink'] || null;
    }

    data.users.sort((a, b) => (a.displayName || '').localeCompare(b.displayName || ''));
    data.fetchedAt = Date.now();
    return data;
  }

  private async _fetchUsersFromSearch(): Promise<IRawUserData> {
    const client = this._client;
    if (!client) throw new Error('SharePoint client not available');

    const data = emptyRawData();
    const seen = new Set<string>();
    let complete = false;

    for (let page = 0; page < MAX_SEARCH_PAGES; page++) {
      // Stable sort so startrow paging doesn't skip or repeat rows
      const url =
        `${this._webUrl}/_api/search/query` +
        `?querytext='*'` +
        `&sourceid='${PEOPLE_SOURCE}'` +
        `&selectproperties='${SELECT_PROPS}'` +
        `&sortlist=${encodeURIComponent("'[DocId]:ascending'")}` +
        `&rowlimit=${BATCH_SIZE}` +
        `&startrow=${page * BATCH_SIZE}` +
        `&trimduplicates=false`;

      const json = await withRetry(async () => {
        const resp = await client.get(
          url,
          SPHttpClient.configurations.v1,
          { headers: { 'Accept': 'application/json;odata=nometadata' } }
        );
        if (!resp.ok) {
          throw createHttpError(`People search failed: ${resp.status}`, resp.status, resp.headers.get('Retry-After'));
        }
        return resp.json();
      });

      const results = json?.PrimaryQueryResult?.RelevantResults;
      if (!results) { complete = true; break; }

      const rows: Array<{ Cells: Array<{ Key: string; Value: string }> }> =
        results.Table?.Rows || [];

      let duplicates = 0;
      for (const row of rows) {
        const parsed = this._rowToUser(row.Cells);
        if (!parsed) continue;
        const { user, managerId, hasPicture } = parsed;
        if (seen.has(user.id)) { duplicates++; continue; }
        seen.add(user.id);
        data.users.push(user);
        if (managerId) data.managerIds.set(user.id, managerId);
        if (!hasPicture) data.noPhotoIds.add(user.id);
      }

      // Only a short page marks the end (TotalRows is an estimate); a page of
      // nothing but repeats means paging has stopped advancing
      if (rows.length < BATCH_SIZE || (rows.length > 0 && duplicates === rows.length)) { complete = true; break; }
    }
    if (!complete) console.warn(`[SmartOrgChart] People search stopped after ${MAX_SEARCH_PAGES} pages; results may be incomplete.`);

    data.users.sort((a, b) => (a.displayName || '').localeCompare(b.displayName || ''));
    data.fetchedAt = Date.now();
    return data;
  }

  private _rowToUser(cells: Array<{ Key: string; Value: string }>):
      { user: IGraphUser; managerId: string; hasPicture: boolean } | null {
    const p: { [key: string]: string } = {};
    for (const cell of cells) p[cell.Key] = cell.Value || '';

    const accountName = p['AccountName'] || '';
    const displayName = p['PreferredName'] || p['DisplayName'] || '';
    if (!accountName || !displayName) return null;

    // Strip claims prefix from AccountName: i:0#.f|membership|user@domain → user@domain
    const upn = (accountName.includes('|')
      ? (accountName.split('|').pop() || '')
      : accountName).toLowerCase();

    // Email: use WorkEmail if set, fall back to UPN
    const email = (p['WorkEmail'] || upn).toLowerCase();

    // Use UPN as id — same format Manager field uses, so parent-child matching
    // works correctly even in tenants where UPN ≠ WorkEmail.
    const id = upn || email;
    if (!id || !displayName) return null;

    // Manager: extract from claims format, bare email, or store raw (display name fallback)
    const managerRaw = p['Manager'] || '';
    let managerId = '';
    if (managerRaw.includes('|')) {
      managerId = (managerRaw.split('|').pop() || '').toLowerCase();
    } else if (managerRaw.trim()) {
      // bare email or display name — store as-is for _buildMaps to resolve
      managerId = managerRaw.trim().toLowerCase();
    }

    // Photo: only use a URL when the user actually has a profile picture
    const hasPicture = !!(p['PictureURL'] && p['PictureURL'].trim());

    const user: IGraphUser = {
      id,
      displayName,
      mail:              p['WorkEmail'] || upn,
      jobTitle:          p['JobTitle']   || '',
      mobilePhone:       p['MobilePhone'] || '',
      businessPhones:    p['WorkPhone'] ? [p['WorkPhone']] : [],
      department:        p['Department'] || '',
      officeLocation:    p['OfficeNumber'] || '',
      userPrincipalName: upn,
    };
    // B2B guest accounts carry '#ext#' in their UPN. Search has no
    // account-enabled flag, so accountEnabled stays undefined and the
    // "Hide disabled accounts" filter can't apply in Search mode.
    if (accountName.toLowerCase().indexOf('#ext#') >= 0) user.userType = 'Guest';
    // customAttributes is intentionally left unset here: SharePoint Search has
    // no generic way to read Entra ID/AAD extension attributes — only
    // properties an admin has explicitly mapped to a managed property, which
    // this app has no way to discover. Custom attributes need Graph.
    return { user, managerId, hasPicture };
  }
}

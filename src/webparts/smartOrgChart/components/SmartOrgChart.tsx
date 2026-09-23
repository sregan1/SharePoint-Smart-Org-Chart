import * as React from 'react';
import { IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { GraphService, ICustomAttributeConfig, IUserFilterOptions } from '../../../services/GraphService';
import { MockGraphService, MockCompanySize } from '../../../services/MockGraphService';
import { ISmartOrgChartProps, IUserSettings } from './ISmartOrgChartProps';
import { EmployeeDirectory } from './EmployeeDirectory/EmployeeDirectory';
import { OrgChart } from './OrgChart/OrgChart';
import { SettingsPanel } from './SettingsPanel/SettingsPanel';
import styles from './SmartOrgChart.module.scss';

const VIEW_META = {
  directory: { label: 'Employee Directory', icon: 'People', toggleIcon: 'Org',  toggleTitle: 'Switch to Org Chart' },
  orgchart:  { label: 'Org Chart',          icon: 'Org',    toggleIcon: 'People', toggleTitle: 'Switch to Employee Directory' },
};

const LS_KEY      = 'smartOrgChart_userSettings';
const LS_MOCK_KEY = 'smartOrgChart_mockSize';
const LS_VIEW_KEY = 'smartOrgChart_currentView';

// Stored alongside the user settings: true when the user explicitly picked a
// font size. Without it, the admin's Default Font Size applies (and follows
// later changes to that default).
const FONT_OVERRIDE_FLAG = 'fontScaleOverride';

// All SharePoint sites in a tenant share one origin, so bare keys would be
// shared by every web part instance on every page. Scope them by instance ID,
// reading the legacy un-scoped key as a migration fallback.
function scopedKey(base: string, instanceId: string): string {
  return instanceId ? `${base}_${instanceId}` : base;
}

function lsGet(base: string, instanceId: string): string | null {
  try {
    return localStorage.getItem(scopedKey(base, instanceId)) ?? localStorage.getItem(base);
  } catch {
    return null;
  }
}

function lsSet(base: string, instanceId: string, value: string): void {
  try { localStorage.setItem(scopedKey(base, instanceId), value); } catch { /* ignore */ }
}

interface ISmartOrgChartState {
  currentView: 'directory' | 'orgchart';
  isSettingsOpen: boolean;
  graphService: GraphService | null;
  userSettings: IUserSettings;
  mockSize: MockCompanySize;
  serviceGen: number;
  isRefreshing: boolean;
}

interface ILoadedSettings {
  settings: IUserSettings;
  fontScaleOverridden: boolean;
}

function loadUserSettings(defaultFontScale: number, instanceId: string): ILoadedSettings {
  const defaults = buildDefaultSettings(defaultFontScale);
  try {
    const stored = lsGet(LS_KEY, instanceId);
    if (stored) {
      const parsed = JSON.parse(stored) || {};
      // Legacy entries have no flag and always stored fontScale; treat them as
      // an override only when they differ from the current admin default.
      const overridden = typeof parsed.fontScale === 'number' && (
        parsed[FONT_OVERRIDE_FLAG] === true ||
        (parsed[FONT_OVERRIDE_FLAG] === undefined && parsed.fontScale !== defaultFontScale)
      );
      delete parsed[FONT_OVERRIDE_FLAG];
      if (!overridden) delete parsed.fontScale;
      return { settings: { ...defaults, ...parsed }, fontScaleOverridden: overridden };
    }
  } catch {
    // ignore
  }
  return { settings: defaults, fontScaleOverridden: false };
}

function buildDefaultSettings(defaultFontScale = 1): IUserSettings {
  return {
    alphabetFilterField: 'firstName',
    cardSize: 'medium',
    showEmail: true,
    showPhone: true,
    showDepartment: true,
    showOffice: true,
    levelsAbove: 1,
    compactCards: false,
    fontScale: defaultFontScale,
  };
}

function readMockSize(instanceId: string): MockCompanySize {
  const v = lsGet(LS_MOCK_KEY, instanceId);
  if (v === '500')  return 500;
  if (v === '1000') return 1000;
  return 150;
}

function readCurrentView(fallback: 'directory' | 'orgchart', instanceId: string): 'directory' | 'orgchart' {
  const v = lsGet(LS_VIEW_KEY, instanceId);
  if (v === 'directory' || v === 'orgchart') return v;
  return fallback;
}

function formatLastLoaded(date: Date | null): string {
  if (!date) return 'not loaded yet';
  const mins = Math.floor((Date.now() - date.getTime()) / 60000);
  if (mins < 1) return 'just now';
  if (mins < 60) return `${mins} min ago`;
  const hours = Math.floor(mins / 60);
  if (hours < 24) return `${hours} hr ago`;
  return date.toLocaleString();
}

export class SmartOrgChart extends React.Component<ISmartOrgChartProps, ISmartOrgChartState> {
  private _instanceId: string;
  private _mounted = false;
  // Incremented on every service (re)creation and on unmount; async init work
  // bails out when its request number is no longer current
  private _initRequest = 0;
  private _fontScaleOverridden: boolean;

  constructor(props: ISmartOrgChartProps) {
    super(props);
    this._instanceId   = props.context?.instanceId || '';
    const loaded       = loadUserSettings(props.defaultFontScale || 1, this._instanceId);
    this._fontScaleOverridden = loaded.fontScaleOverridden;
    const mockSize     = readMockSize(this._instanceId);
    this.state = {
      currentView: readCurrentView(props.defaultView || 'directory', this._instanceId),
      isSettingsOpen: false,
      graphService: null,
      userSettings: loaded.settings,
      mockSize,
      serviceGen: 0,
      isRefreshing: false,
    };
  }

  public async componentDidMount(): Promise<void> {
    this._mounted = true;
    await this._initGraphService();
  }

  public componentWillUnmount(): void {
    this._mounted = false;
    this._initRequest++;
  }

  public async componentDidUpdate(prev: ISmartOrgChartProps): Promise<void> {
    // Admin changed Default Font Size — apply it unless this user picked their own
    if (prev.defaultFontScale !== this.props.defaultFontScale && !this._fontScaleOverridden) {
      const fontScale = this.props.defaultFontScale || 1;
      this.setState(s => ({ userSettings: { ...s.userSettings, fontScale } }));
    }

    const serviceChanged =
      prev.useDemoData         !== this.props.useDemoData ||
      prev.dataSource          !== this.props.dataSource  ||
      prev.dottedLineAttribute !== this.props.dottedLineAttribute;
    const filtersChanged =
      prev.excludedAccounts       !== this.props.excludedAccounts ||
      prev.hideDisabledAccounts   !== this.props.hideDisabledAccounts ||
      prev.hideGuestUsers         !== this.props.hideGuestUsers ||
      prev.restrictToTenantDomain !== this.props.restrictToTenantDomain ||
      prev.hideNoJobTitle         !== this.props.hideNoJobTitle ||
      prev.hideNoDepartment       !== this.props.hideNoDepartment;
    // Compared by content, not reference — which custom attributes are
    // configured changes what GraphService fetches ($select), so it needs a
    // fresh service; unlike the filters above, updateFilterOptions can't apply it.
    const customFieldsChanged = this._customFieldsKey(prev.customAttributes) !== this._customFieldsKey(this.props.customAttributes);

    if (serviceChanged || customFieldsChanged) {
      await this._initGraphService();
    } else if (filtersChanged && this.state.graphService) {
      // Re-filter the already-downloaded data in memory — no tenant re-download.
      // Bumping serviceGen remounts the views so they read the new result.
      this.state.graphService.updateFilterOptions(this._buildFilterOptions(this._getTenantDomain()));
      this.setState(s => ({ serviceGen: s.serviceGen + 1 }));
    }
  }

  private _isDemoMode(): boolean {
    return window.location.hostname === 'localhost' || !!this.props.useDemoData;
  }

  /**
   * Tenant domain for "Only show tenant users", derived from the current
   * user's email. Undefined when the restriction is off, and in demo mode —
   * demo users are @contoso.com, so restricting to the real tenant's domain
   * would hide everyone.
   */
  private _getTenantDomain(): string | undefined {
    if (this._isDemoMode() || !this.props.restrictToTenantDomain) return undefined;
    const userEmail = (this.props.context?.pageContext?.user?.email || '').toLowerCase();
    const atIdx = userEmail.lastIndexOf('@');
    return atIdx > 0 ? userEmail.substring(atIdx + 1) : undefined;
  }

  private _customFieldsKey(config: ICustomAttributeConfig[] | undefined): string {
    return (config || []).map(c => c.graphField).sort().join(',');
  }

  private _buildFilterOptions(tenantDomain?: string): IUserFilterOptions {
    return {
      tenantDomain,
      excludedPatterns: (this.props.excludedAccounts || '')
        .split(',')
        .map(s => s.trim().toLowerCase())
        .filter(s => s.length > 0),
      hideGuestUsers:       this.props.hideGuestUsers       || false,
      hideDisabledAccounts: this.props.hideDisabledAccounts || false,
      hideNoJobTitle:       this.props.hideNoJobTitle       || false,
      hideNoDepartment:     this.props.hideNoDepartment     || false,
    };
  }

  private async _initGraphService(): Promise<void> {
    const request = ++this._initRequest;
    if (this._isDemoMode()) {
      const filterOptions = this._buildFilterOptions(this._getTenantDomain());
      this.setState(prev => ({
        graphService: new MockGraphService(prev.mockSize, filterOptions) as unknown as GraphService,
        serviceGen: prev.serviceGen + 1,
      }));
      return;
    }
    const { spHttpClient, msGraphClientFactory, pageContext } = this.props.context;
    let graphClient: MSGraphClientV3 | undefined;
    try {
      graphClient = await msGraphClientFactory.getClient('3');
    } catch {
      // Graph client unavailable — fall back to SP Search only
    }
    // A newer init (or an unmount) happened while awaiting the client
    if (!this._mounted || request !== this._initRequest) return;

    const filterOptions = this._buildFilterOptions(this._getTenantDomain());

    const service = new GraphService(
      spHttpClient,
      pageContext.web.absoluteUrl,
      graphClient,
      this.props.dataSource || 'auto',
      filterOptions,
      this.props.dottedLineAttribute || '',
      (this.props.customAttributes || []).map(c => c.graphField)
    );
    this.setState(prev => ({ graphService: service, serviceGen: prev.serviceGen + 1 }));
  }

  private _setMockSize = (size: MockCompanySize): void => {
    lsSet(LS_MOCK_KEY, this._instanceId, String(size));
    this._initRequest++;
    this.setState(prev => ({
      mockSize: size,
      graphService: new MockGraphService(size, this._buildFilterOptions()) as unknown as GraphService,
      serviceGen: prev.serviceGen + 1,
    }));
  }

  private _toggleView = (): void => {
    const newView: 'directory' | 'orgchart' = this.state.currentView === 'directory' ? 'orgchart' : 'directory';
    lsSet(LS_VIEW_KEY, this._instanceId, newView);
    this.setState({ currentView: newView });
  }

  private _refreshTitle(): string {
    const service = this.state.graphService;
    return `Refresh data (last updated ${formatLastLoaded(service ? service.getLastLoaded() : null)})`;
  }

  // The last-loaded time changes when the views finish loading, which doesn't
  // re-render this component — so update the tooltip just before it shows.
  private _updateRefreshTitle = (e: React.SyntheticEvent<unknown>): void => {
    (e.currentTarget as unknown as HTMLElement).title = this._refreshTitle();
  }

  private _refreshData = async (): Promise<void> => {
    const service = this.state.graphService;
    if (!service || this.state.isRefreshing) return;
    this.setState({ isRefreshing: true });
    try {
      await service.refresh();
    } catch {
      // ignore — the views surface any load error when they remount
    }
    if (!this._mounted) return;
    this.setState(s => ({ isRefreshing: false, serviceGen: s.serviceGen + 1 }));
  }

  private _openSettings = (): void => {
    this.setState({ isSettingsOpen: true });
  }

  private _closeSettings = (): void => {
    this.setState({ isSettingsOpen: false });
  }

  private _saveSettings = (settings: IUserSettings): void => {
    // Only persist the font size when the user chose something other than the
    // admin default, so later changes to that default still reach them
    const overridden = (settings.fontScale || 1) !== (this.props.defaultFontScale || 1);
    this._fontScaleOverridden = overridden;
    const toStore: { [key: string]: unknown } = { ...settings, [FONT_OVERRIDE_FLAG]: overridden };
    if (!overridden) delete toStore.fontScale;
    lsSet(LS_KEY, this._instanceId, JSON.stringify(toStore));
    this.setState({ userSettings: settings, isSettingsOpen: false });
  }

  public render(): React.ReactElement<ISmartOrgChartProps> {
    const { currentView, isSettingsOpen, graphService, userSettings, mockSize, serviceGen, isRefreshing } = this.state;
    const { theme, accentColor, defaultLayout, logoUrl, companyName } = this.props;
    const meta        = VIEW_META[currentView];
    const resolvedLogoUrl = (() => {
      if (!logoUrl) return '';
      if (logoUrl.startsWith('http://') || logoUrl.startsWith('https://')) return logoUrl;
      // Relative paths can't be served from the local dev server
      if (window.location.hostname === 'localhost') return '';
      if (logoUrl.startsWith('/')) return `${window.location.origin}${logoUrl}`;
      // Site-relative (no leading slash) — resolve against the SharePoint site URL
      const siteUrl = (this.props.context?.pageContext?.web?.absoluteUrl || '').replace(/\/$/, '');
      return siteUrl ? `${siteUrl}/${logoUrl}` : '';
    })();

    return (
      <div className={styles.container} style={{ '--soc-font-scale': userSettings.fontScale || 1 } as React.CSSProperties}>
        <div className={styles.header}>
          <div className={styles.brandArea}>
            {resolvedLogoUrl && (
              <img
                key={resolvedLogoUrl}
                src={resolvedLogoUrl}
                // Decorative when the company name is shown right beside it
                alt={companyName ? '' : 'Company logo'}
                className={styles.logo}
                onError={e => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }}
              />
            )}
            <div className={styles.viewTitle}>
              {companyName && (
                <>
                  <span className={styles.companyName}>{companyName}</span>
                  <span className={styles.brandDivider}>·</span>
                </>
              )}
              <Icon iconName={meta.icon} className={styles.viewIcon} />
              <span>{meta.label}</span>
            </div>
          </div>

          <div className={styles.headerActions}>
            <IconButton
              iconProps={{ iconName: 'Refresh' }}
              title={this._refreshTitle()}
              ariaLabel={isRefreshing ? 'Refreshing data' : 'Refresh data'}
              onClick={this._refreshData}
              onMouseEnter={this._updateRefreshTitle}
              onFocus={this._updateRefreshTitle}
              disabled={!graphService || isRefreshing}
              className={styles.actionBtn}
            />
            <IconButton
              iconProps={{ iconName: meta.toggleIcon }}
              title={meta.toggleTitle}
              ariaLabel={meta.toggleTitle}
              onClick={this._toggleView}
              className={styles.actionBtn}
            />
            <IconButton
              iconProps={{ iconName: 'Settings' }}
              title="Preferences"
              ariaLabel="Open preferences panel"
              onClick={this._openSettings}
              className={styles.actionBtn}
            />
          </div>
        </div>

        <div className={styles.content}>
          {currentView === 'directory' && graphService && (
            <EmployeeDirectory
              key={`dir-${serviceGen}`}
              graphService={graphService}
              instanceId={this._instanceId}
              alphabetFilterField={userSettings.alphabetFilterField}
              cardSize={userSettings.cardSize}
              showEmail={userSettings.showEmail}
              showPhone={userSettings.showPhone}
              showDepartment={userSettings.showDepartment}
              showOffice={userSettings.showOffice}
              pageSize={this.props.pageSize}
              theme={theme}
              accentColor={accentColor}
              customAttributes={this.props.customAttributes || []}
            />
          )}

          {currentView === 'orgchart' && graphService && (
            <OrgChart
              key={`org-${serviceGen}`}
              graphService={graphService}
              instanceId={this._instanceId}
              topLevelUser={this.props.topLevelUser}
              levelsBelow={this.props.levelsBelow}
              levelsAbove={userSettings.levelsAbove}
              showDepartment={userSettings.showDepartment}
              showOffice={userSettings.showOffice}
              theme={theme}
              accentColor={accentColor}
              customAttributes={this.props.customAttributes || []}
              currentUserEmail={this.props.context?.pageContext?.user?.email || ''}
              compactCards={userSettings.compactCards}
              defaultLayout={defaultLayout || 'drill'}
              enableFindMe={this.props.enableFindMe !== false}
              enableLayoutToggle={this.props.enableLayoutToggle !== false}
              enableStats={this.props.enableStats !== false}
              enableDeptFilter={this.props.enableDeptFilter !== false}
              enableUserFilter={this.props.enableUserFilter !== false}
              defaultZoom={this.props.defaultZoom ?? 0}
            />
          )}
        </div>

        <SettingsPanel
          isOpen={isSettingsOpen}
          settings={userSettings}
          onDismiss={this._closeSettings}
          onSave={this._saveSettings}
          mockSize={this._isDemoMode() ? mockSize : undefined}
          onMockSizeChange={this._isDemoMode() ? this._setMockSize : undefined}
        />
      </div>
    );
  }
}

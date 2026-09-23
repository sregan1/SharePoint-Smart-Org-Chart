import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { IGraphUser } from '../../../../services/GraphService';
import { getInitials } from '../personUtils';
import { IFilterCounts } from './orgTreeUtils';
import { CHART_LAYOUTS, ChartLayout } from './chartPersistence';
import { OrgChartTheme, getThemeTokens } from './orgTheme';
import { FilterPanel, UserFilterKey } from './ChartControls';
import styles from './OrgChart.module.scss';

const LAYOUT_ICON: Record<ChartLayout, string> = {
  drill:      'Org',
  vertical:   'Down',
  horizontal: 'Forward',
};

const LAYOUT_TITLE: Record<ChartLayout, string> = {
  drill:      'Drill-Down',
  vertical:   'Top Down',
  horizontal: 'Left to Right',
};

export interface IOrgChartToolbarProps {
  theme: OrgChartTheme;
  accentColor?: string;
  isDrillMode: boolean;

  // Search (combobox + listbox)
  searchRef: React.RefObject<HTMLDivElement>;
  /** Unique prefix for listbox/option ids */
  idPrefix: string;
  searchQuery: string;
  appliedQuery: string;
  matchCount: number;
  searchResults: IGraphUser[];
  showSearchResults: boolean;
  searchActiveIndex: number;
  onSearchChange: (e?: React.ChangeEvent<HTMLInputElement>, value?: string) => void;
  onSearchFocus: () => void;
  onSearchKeyDown: (e: React.KeyboardEvent<HTMLInputElement>) => void;
  onSearchEnter: (value?: string) => void;
  onSearchEscape: () => void;
  onSelectSearchResult: (user: IGraphUser) => void;

  // Expand / collapse and progress
  isExpandingAll: boolean;
  isRefocusing: boolean;
  onExpandAll: () => void;
  onCollapseAll: () => void;

  showFindMe: boolean;
  onFindMe: () => void;

  enableLayoutToggle: boolean;
  chartLayout: ChartLayout;
  showLayoutPicker: boolean;
  onToggleLayoutPicker: () => void;
  onSelectLayout: (layout: ChartLayout) => void;

  enableStats: boolean;
  showStats: boolean;
  onToggleStats: () => void;

  enableDeptFilter: boolean;
  uniqueDepts: Map<string, number>;
  filterDepartments: Set<string>;
  showDeptFilter: boolean;
  onToggleDeptFilter: () => void;
  onToggleDept: (dept: string) => void;
  onClearDeptFilter: () => void;

  enableUserFilter: boolean;
  filterMembers: boolean;
  filterGuests: boolean;
  filterDisabled: boolean;
  treeCounts: IFilterCounts;
  showFilters: boolean;
  onToggleUserFilterPanel: () => void;
  onToggleUserFilter: (key: UserFilterKey) => void;

  onExportPdf: () => void;
  onExportCsv: () => void;

  rootPickerRef: React.RefObject<HTMLDivElement>;
  runtimeRootUser: IGraphUser | null;
  rootPickerQuery: string;
  rootPickerResults: IGraphUser[];
  onRootPickerChange: (e: React.ChangeEvent<HTMLInputElement>) => void;
  onRootPickerSelect: (user: IGraphUser) => void;
  onResetRoot: () => void;

  zoomLevel: number;
  resetZoom: number;
  onZoomOut: () => void;
  onZoomIn: () => void;
  onZoomReset: () => void;
}

const preventMouseFocus = (e: React.MouseEvent): void => { e.preventDefault(); };

export const OrgChartToolbar: React.FC<IOrgChartToolbarProps> = (p) => {
  const t = getThemeTokens(p.theme, p.accentColor);
  const { isDrillMode, appliedQuery, matchCount, searchResults, searchActiveIndex } = p;
  const listOpen  = p.showSearchResults && searchResults.length > 0;
  const listboxId = `${p.idPrefix}-results`;
  const optionId  = (i: number): string => `${p.idPrefix}-opt-${i}`;
  const activeIdx = listOpen && searchActiveIndex < searchResults.length ? searchActiveIndex : -1;
  const activeFilters = (!p.filterMembers ? 1 : 0) + (!p.filterGuests ? 1 : 0) + (!p.filterDisabled ? 1 : 0);
  const deptCount = p.filterDepartments.size;
  const isBusy = p.isRefocusing || p.isExpandingAll;

  return (
    <div className={styles.chartToolbar}>

      {/* Search with results dropdown */}
      <div className={styles.searchWrapper} ref={p.searchRef}>
        <SearchBox
          placeholder="Search people..."
          ariaLabel="Search people"
          value={p.searchQuery}
          onChange={p.onSearchChange}
          onFocus={p.onSearchFocus}
          onKeyDown={p.onSearchKeyDown}
          onSearch={p.onSearchEnter}
          onEscape={p.onSearchEscape}
          className={styles.chartSearch}
          aria-autocomplete="list"
          aria-controls={listOpen ? listboxId : undefined}
          aria-activedescendant={activeIdx >= 0 ? optionId(activeIdx) : undefined}
          underlined
        />
        {appliedQuery && !listOpen && !isDrillMode && (
          <span className={styles.chartSearchHit} role="status">
            {matchCount} {matchCount === 1 ? 'match' : 'matches'} in tree
          </span>
        )}
        {listOpen && (
          <div className={styles.searchResults} role="listbox" id={listboxId} aria-label="People">
            {searchResults.map((u, i) => (
              <div
                key={u.id}
                id={optionId(i)}
                role="option"
                aria-selected={i === activeIdx}
                className={`${styles.searchResult} ${i === activeIdx ? styles.searchResultActive : ''}`}
                onMouseDown={preventMouseFocus}
                onClick={() => p.onSelectSearchResult(u)}
              >
                <span className={styles.searchResultInitials} style={{ background: t.accent, color: t.onAccent }}>
                  {getInitials(u.displayName)}
                </span>
                <span className={styles.searchResultInfo}>
                  <span className={styles.searchResultName}>{u.displayName}</span>
                  <span className={styles.searchResultMeta}>{[u.jobTitle, u.department].filter(Boolean).join(' · ')}</span>
                </span>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Expand/Collapse — only in full-tree mode */}
      {!isDrillMode && (
        <div className={styles.chartActions}>
          <button className={styles.chartActionBtn} onClick={p.onExpandAll} disabled={p.isExpandingAll}>
            <Icon iconName="ExploreContent" /> Expand All
          </button>
          <button className={styles.chartActionBtn} onClick={p.onCollapseAll}>
            <Icon iconName="CollapseContent" /> Collapse All
          </button>
        </div>
      )}

      {/* Inline progress for re-rooting / Expand All (the chart stays on screen) */}
      {isBusy && (
        <span className={styles.refocusIndicator}>
          <Spinner size={SpinnerSize.xSmall} ariaLive="polite" label={p.isExpandingAll ? 'Expanding…' : 'Loading…'} labelPosition="right" />
        </span>
      )}

      {/* Find Me */}
      {p.showFindMe && (
        <button
          className={styles.iconToolBtn}
          onClick={p.onFindMe}
          title="Find me in the org chart"
          aria-label="Find me in the org chart"
        >
          <Icon iconName="Contact" />
        </button>
      )}

      {/* Layout picker */}
      {p.enableLayoutToggle && (
        <div className={styles.toolbarPopupAnchor}>
          <button
            className={`${styles.chartActionBtn} ${p.showLayoutPicker ? styles.iconToolBtnActive : ''}`}
            onClick={p.onToggleLayoutPicker}
            title="Switch view layout"
            aria-haspopup="true"
            aria-expanded={p.showLayoutPicker}
          >
            <Icon iconName="ViewAll" />
            <span>View</span>
          </button>
          {p.showLayoutPicker && (
            <div className={styles.filterPanel} style={{ minWidth: 210 }} role="group" aria-label="View layout">
              <div className={styles.filterPanelTitle}>View layout</div>
              {CHART_LAYOUTS.map(layout => {
                const selected = p.chartLayout === layout;
                return (
                  <button
                    key={layout}
                    className={styles.filterItem}
                    aria-pressed={selected}
                    style={{
                      border: 'none',
                      background: selected ? t.selectedBg : 'transparent',
                      cursor: 'pointer',
                      width: '100%',
                      textAlign: 'left',
                      display: 'flex',
                      alignItems: 'center',
                      gap: 8,
                      padding: '8px 10px',
                      borderRadius: 4,
                      fontWeight: selected ? 600 : 400,
                      color: selected ? t.selectedText : 'inherit',
                    }}
                    onClick={() => p.onSelectLayout(layout)}
                  >
                    <Icon iconName={LAYOUT_ICON[layout]} />
                    <span style={{ flex: 1 }}>{LAYOUT_TITLE[layout]}</span>
                    {selected && <Icon iconName="CheckMark" />}
                  </button>
                );
              })}
            </div>
          )}
        </div>
      )}

      {/* Stats toggle */}
      {p.enableStats && (
        <button
          className={`${styles.iconToolBtn} ${p.showStats ? styles.iconToolBtnActive : ''}`}
          onClick={p.onToggleStats}
          title="Org stats summary"
          aria-label="Org stats summary"
          aria-pressed={p.showStats}
        >
          <Icon iconName="BarChartVertical" />
        </button>
      )}

      {/* Dept filter button */}
      {p.enableDeptFilter && <div className={styles.toolbarPopupAnchor}>
        <button
          className={`${styles.iconToolBtn} ${deptCount > 0 ? styles.iconToolBtnActive : ''}`}
          onClick={p.onToggleDeptFilter}
          title="Filter by department"
          aria-label={deptCount > 0 ? `Filter by department (${deptCount} selected)` : 'Filter by department'}
          aria-haspopup="true"
          aria-expanded={p.showDeptFilter}
        >
          <Icon iconName="DeveloperTools" />
          {deptCount > 0 && <span className={styles.toolBtnBadge}>{deptCount}</span>}
        </button>
        {p.showDeptFilter && (
          <div className={styles.filterPanel} style={{ minWidth: 220 }} role="group" aria-label="Filter by department">
            <div className={styles.filterPanelTitle}>Filter by department</div>
            {Array.from(p.uniqueDepts.entries()).sort((a, b) => a[0].localeCompare(b[0])).map(([dept, count]) => (
              <label key={dept} className={styles.filterItem}>
                <input
                  type="checkbox"
                  className={styles.filterCheckbox}
                  checked={p.filterDepartments.has(dept)}
                  onChange={() => p.onToggleDept(dept)}
                />
                <span className={styles.filterLabel}>{dept}</span>
                <span className={styles.filterCount}>{count}</span>
              </label>
            ))}
            {deptCount > 0 && (
              <button
                className={styles.filterItem}
                style={{ border: 'none', background: 'none', cursor: 'pointer', color: t.linkText, fontWeight: 600, fontSize: 12 }}
                onClick={p.onClearDeptFilter}
              >
                Clear all filters
              </button>
            )}
          </div>
        )}
      </div>}

      {/* Filter button */}
      {p.enableUserFilter && (
        <div className={styles.toolbarPopupAnchor}>
          <button
            className={`${styles.iconToolBtn} ${activeFilters > 0 ? styles.iconToolBtnActive : ''}`}
            onClick={p.onToggleUserFilterPanel}
            title="Filter user types"
            aria-label="Filter user types"
            aria-haspopup="true"
            aria-expanded={p.showFilters}
          >
            <Icon iconName="Filter" />
            {activeFilters > 0 && <span className={styles.toolBtnBadge}>{activeFilters}</span>}
          </button>
          {p.showFilters && (
            <FilterPanel
              filterMembers={p.filterMembers}
              filterGuests={p.filterGuests}
              filterDisabled={p.filterDisabled}
              counts={p.treeCounts}
              onToggle={p.onToggleUserFilter}
            />
          )}
        </div>
      )}

      {/* Export PDF / CSV */}
      <button className={styles.iconToolBtn} onClick={p.onExportPdf} title="Download as PDF" aria-label="Download as PDF">
        <Icon iconName="PDF" />
      </button>
      <button className={styles.iconToolBtn} onClick={p.onExportCsv} title="Download as CSV spreadsheet" aria-label="Download as CSV spreadsheet">
        <Icon iconName="ExcelDocument" />
      </button>

      {/* Root picker — "View from person…" */}
      <div className={styles.rootPickerWrapper} ref={p.rootPickerRef}>
        {p.runtimeRootUser ? (
          <div className={styles.rootPickerActive}>
            <Icon iconName="Org" className={styles.rootPickerIcon} />
            <span className={styles.rootPickerActiveName}>{p.runtimeRootUser.displayName}</span>
            <button
              className={styles.rootPickerReset}
              onClick={p.onResetRoot}
              title="Reset to default root"
              aria-label="Reset to default root"
            >
              <Icon iconName="Cancel" />
            </button>
          </div>
        ) : (
          <>
            <div className={styles.rootPickerInputWrap}>
              <Icon iconName="Org" className={styles.rootPickerIcon} />
              <input
                type="text"
                className={styles.rootPickerInput}
                placeholder="View from person…"
                aria-label="View the org chart from a person"
                value={p.rootPickerQuery}
                onChange={p.onRootPickerChange}
              />
            </div>
            {p.rootPickerResults.length > 0 && (
              <div className={styles.rootPickerDropdown}>
                {p.rootPickerResults.map(u => (
                  <button
                    key={u.id}
                    className={styles.rootPickerOption}
                    onClick={() => p.onRootPickerSelect(u)}
                  >
                    <span className={styles.rootPickerOptionInitials} style={{ background: t.accent, color: t.onAccent }}>
                      {getInitials(u.displayName)}
                    </span>
                    <span className={styles.rootPickerOptionInfo}>
                      <span className={styles.rootPickerOptionName}>{u.displayName}</span>
                      {u.jobTitle && <span className={styles.rootPickerOptionMeta}>{u.jobTitle}</span>}
                    </span>
                  </button>
                ))}
              </div>
            )}
          </>
        )}
      </div>

      {/* Zoom — only in full-tree mode */}
      {!isDrillMode && (
        <div className={styles.zoomControls}>
          <button className={styles.zoomBtn} onClick={p.onZoomOut} title="Zoom out" aria-label="Zoom out" disabled={p.zoomLevel <= 0.25}>
            <Icon iconName="Remove" />
          </button>
          <span className={styles.zoomLabel} aria-live="polite">{Math.round(p.zoomLevel * 100)}%</span>
          <button className={styles.zoomBtn} onClick={p.onZoomIn} title="Zoom in" aria-label="Zoom in" disabled={p.zoomLevel >= 1.5}>
            <Icon iconName="Add" />
          </button>
          <button className={styles.zoomBtn} onClick={p.onZoomReset} title="Reset zoom" aria-label="Reset zoom" disabled={p.zoomLevel === p.resetZoom}>
            <Icon iconName="Refresh" />
          </button>
        </div>
      )}
    </div>
  );
};

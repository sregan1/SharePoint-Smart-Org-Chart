// Chart state persistence (localStorage) and ?socFocus= deep links.

export type ChartLayout = 'drill' | 'vertical' | 'horizontal';

export const CHART_LAYOUTS: ChartLayout[] = ['drill', 'vertical', 'horizontal'];

export function isChartLayout(v: unknown): v is ChartLayout {
  return typeof v === 'string' && CHART_LAYOUTS.indexOf(v as ChartLayout) !== -1;
}

export interface IChartStoredState {
  chartLayout?: ChartLayout;
  showStats?: boolean;
  filterMembers?: boolean;
  filterGuests?: boolean;
  filterDisabled?: boolean;
  filterDepartments?: string[];
  focusEmail?: string | null;
}

const LS_CHART_KEY = 'smartOrgChart_chartState';

// Storage is scoped per web part instance — all SharePoint sites share one
// origin, so a bare key would leak state between instances on different pages.
function chartStateKey(instanceId: string): string {
  return instanceId ? `${LS_CHART_KEY}_${instanceId}` : LS_CHART_KEY;
}

// Stored JSON is untrusted (older versions, manual edits, other code on the
// origin) — keep only fields with the expected shape.
function sanitizeStoredState(raw: unknown): IChartStoredState {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return {};
  const r = raw as { [key: string]: unknown };
  const out: IChartStoredState = {};
  if (isChartLayout(r.chartLayout)) out.chartLayout = r.chartLayout;
  if (typeof r.showStats === 'boolean') out.showStats = r.showStats;
  if (typeof r.filterMembers === 'boolean') out.filterMembers = r.filterMembers;
  if (typeof r.filterGuests === 'boolean') out.filterGuests = r.filterGuests;
  if (typeof r.filterDisabled === 'boolean') out.filterDisabled = r.filterDisabled;
  if (Array.isArray(r.filterDepartments)) {
    out.filterDepartments = (r.filterDepartments as unknown[])
      .filter((d): d is string => typeof d === 'string' && d.length > 0);
  }
  if (typeof r.focusEmail === 'string' && r.focusEmail) out.focusEmail = r.focusEmail;
  return out;
}

function parseStored(s: string | null): IChartStoredState | null {
  if (!s) return null;
  try { return sanitizeStoredState(JSON.parse(s)); } catch { return null; }
}

export function loadChartState(instanceId: string): IChartStoredState {
  try {
    const key = chartStateKey(instanceId);
    const scoped = parseStored(localStorage.getItem(key));
    if (scoped) return scoped;
    // One-time migration from the legacy un-scoped key: copy it to this
    // instance's key, then delete it so it can't bleed into other instances.
    if (key !== LS_CHART_KEY) {
      const legacyRaw = localStorage.getItem(LS_CHART_KEY);
      if (legacyRaw !== null) {
        const legacy = parseStored(legacyRaw) || {};
        try {
          localStorage.setItem(key, JSON.stringify(legacy));
          localStorage.removeItem(LS_CHART_KEY);
        } catch { /* ignore */ }
        return legacy;
      }
    }
  } catch { /* storage unavailable */ }
  return {};
}

export function saveChartState(instanceId: string, s: IChartStoredState): void {
  try { localStorage.setItem(chartStateKey(instanceId), JSON.stringify(s)); } catch { /* ignore */ }
}

/* ── Deep links (?socFocus=email) ────────── */

const FOCUS_URL_PARAM = 'socFocus';

export function readUrlFocus(): string | null {
  try { return new URLSearchParams(window.location.search).get(FOCUS_URL_PARAM); } catch { return null; }
}

/**
 * Writes (or clears) the ?socFocus= param. `ownValue` is the value this
 * instance last wrote; clearing only removes the param when it still holds
 * that value, so a shared link (or another instance's focus) isn't stripped.
 * Returns the value this instance now owns in the URL.
 */
export function syncUrlFocus(email: string | null, ownValue: string | null): string | null {
  try {
    const url = new URL(window.location.href);
    const current = url.searchParams.get(FOCUS_URL_PARAM);
    if (email) {
      if (current !== email) {
        url.searchParams.set(FOCUS_URL_PARAM, email);
        window.history.replaceState(window.history.state, '', url.toString());
      }
      return email;
    }
    if (current !== null && ownValue !== null && current === ownValue) {
      url.searchParams.delete(FOCUS_URL_PARAM);
      window.history.replaceState(window.history.state, '', url.toString());
    }
    return null;
  } catch {
    return ownValue;
  }
}

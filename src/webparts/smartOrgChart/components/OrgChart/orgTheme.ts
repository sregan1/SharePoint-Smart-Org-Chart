import * as React from 'react';
import { OrgChartTheme } from '../ISmartOrgChartProps';
import { getAccentCssVars, getAccentPalette, mixWithWhite } from '../colorUtils';
import styles from './OrgChart.module.scss';

export type { OrgChartTheme };

export function getSiteColor(theme: OrgChartTheme, accentColor?: string): string {
  if (theme === 'custom')    return getAccentPalette(accentColor).accent;
  if (theme === 'corporate') return '#0052a5';
  try {
    const t = (window as any).__themeState__?.theme;
    if (!t) return theme === 'dark' ? '#71afe5' : '#0078d4';
    return theme === 'dark' ? (t.themeTertiary ?? '#71afe5') : (t.themePrimary ?? '#0078d4');
  } catch {
    return theme === 'dark' ? '#71afe5' : '#0078d4';
  }
}

export const THEME_CONTAINER_CLASS: Record<OrgChartTheme, string> = {
  modern:    '',
  minimal:   styles.themeMinimal,
  corporate: styles.themeCorporate,
  dark:      styles.themeDark,
  custom:    styles.themeCustom,
};

/**
 * Inline style for the chart container. The custom theme exposes its palette
 * as CSS variables (read by .themeCustom rules in the SCSS); other themes
 * need nothing inline.
 */
export function getThemeContainerStyle(theme: OrgChartTheme, accentColor?: string): React.CSSProperties | undefined {
  if (theme !== 'custom') return undefined;
  const vars = getAccentCssVars(accentColor) as { [key: string]: string };
  return {
    ...vars,
    // Connector lines: a mid tint so they stay visible on the light canvas
    '--soc-custom-accent-line': mixWithWhite(getAccentPalette(accentColor).accent, 35),
  } as React.CSSProperties;
}

/** Resolved colors for inline styles — replaces scattered `isDark ? … : …` pairs. */
export interface IThemeTokens {
  isDark: boolean;
  isCustom: boolean;
  /** Accent for fills and borders (card top border, avatar initials, header band) */
  accent: string;
  /** Text drawn on top of an accent fill */
  onAccent: string;
  /** Accent used as text on the card/panel background */
  accentText: string;
  /** Subtle accent tint for badge backgrounds */
  accentSoft: string;
  /** Selected-row styling in toolbar popups (layout picker) */
  selectedBg: string;
  selectedText: string;
  /** Text-style action links in popups ("Clear all filters") */
  linkText: string;
  cardBg: string;
  text: string;
  subText: string;
  fieldBg: string;
  border: string;
  chainBg: string;
  headerBg: string;
  nameColor: string;
  deptText: string;
  managerLine: string;
  neutralBtnBg: string;
  neutralBtnText: string;
}

let _cacheKey = '';
let _cache: IThemeTokens | null = null;

export function getThemeTokens(theme: OrgChartTheme, accentColor?: string): IThemeTokens {
  const accent = getSiteColor(theme, accentColor);
  const key = `${theme}|${accentColor || ''}|${accent}`;
  if (_cache && key === _cacheKey) return _cache;

  const isDark   = theme === 'dark';
  const isCustom = theme === 'custom';
  const palette  = isCustom ? getAccentPalette(accentColor) : null;

  const tokens: IThemeTokens = {
    isDark,
    isCustom,
    accent,
    // Non-custom themes keep their original white-on-accent look
    onAccent:     palette ? palette.onAccent : '#ffffff',
    accentText:   palette ? palette.text : accent,
    accentSoft:   palette ? palette.soft : `${accent}1a`,
    selectedBg:   palette ? palette.soft : isDark ? '#2d3a5a' : '#e8f4fd',
    selectedText: palette ? palette.text : isDark ? '#80b0ff' : accent,
    linkText:     palette ? palette.text : isDark ? '#80b0ff' : accent,
    cardBg:       isDark ? '#242740' : '#ffffff',
    text:         isDark ? '#f0f0f0' : '#1a1a2e',
    subText:      isDark ? '#a0a8c0' : '#555',
    fieldBg:      isDark ? '#1e2138' : '#f8f9fb',
    border:       isDark ? '#3a3d5c' : '#e8ecf0',
    chainBg:      isDark ? '#1a1c2e' : '#f2f4f8',
    headerBg:     isDark ? '#1e2138' : '#ffffff',
    nameColor:    isDark ? '#e8ecff' : '#1a1a2e',
    deptText:     isDark ? '#9098b8' : '#5a6472',
    managerLine:  isDark ? '#8090b0' : '#999',
    neutralBtnBg:   isDark ? '#3a3d5c' : '#eef0f4',
    neutralBtnText: isDark ? '#e0e0f0' : '#333',
  };
  _cacheKey = key;
  _cache = tokens;
  return tokens;
}

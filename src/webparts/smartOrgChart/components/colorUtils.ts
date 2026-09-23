import * as React from 'react';

export const DEFAULT_ACCENT = '#0078d4';

export interface IRgb { r: number; g: number; b: number; }

/**
 * Normalizes user input to a lowercase 6-digit hex string (#rrggbb).
 * Accepts surrounding whitespace, a missing '#', and 3-digit shorthand.
 * Returns null when the input isn't a valid hex color.
 */
export function normalizeHex(input: string | undefined | null): string | null {
  if (!input) return null;
  let v = input.trim().toLowerCase();
  if (v.charAt(0) !== '#') v = '#' + v;
  if (/^#[0-9a-f]{3}$/.test(v)) {
    v = '#' + v.charAt(1) + v.charAt(1) + v.charAt(2) + v.charAt(2) + v.charAt(3) + v.charAt(3);
  }
  return /^#[0-9a-f]{6}$/.test(v) ? v : null;
}

export function hexToRgb(hex: string): IRgb {
  const h = normalizeHex(hex) || DEFAULT_ACCENT;
  return {
    r: parseInt(h.substring(1, 3), 16),
    g: parseInt(h.substring(3, 5), 16),
    b: parseInt(h.substring(5, 7), 16),
  };
}

function toHexByte(n: number): string {
  const s = Math.max(0, Math.min(255, Math.round(n))).toString(16);
  return s.length === 1 ? '0' + s : s;
}

export function rgbToHex(r: number, g: number, b: number): string {
  return '#' + toHexByte(r) + toHexByte(g) + toHexByte(b);
}

function channelLuminance(c: number): number {
  const s = c / 255;
  return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
}

/** WCAG relative luminance (0 = black, 1 = white). */
export function relativeLuminance(hex: string): number {
  const { r, g, b } = hexToRgb(hex);
  return 0.2126 * channelLuminance(r) + 0.7152 * channelLuminance(g) + 0.0722 * channelLuminance(b);
}

/** WCAG contrast ratio between two colors (1–21). */
export function contrastRatio(a: string, b: string): number {
  const la = relativeLuminance(a);
  const lb = relativeLuminance(b);
  const hi = Math.max(la, lb);
  const lo = Math.min(la, lb);
  return (hi + 0.05) / (lo + 0.05);
}

/** Text color to draw on top of the given background: white or near-black. */
export function getContrastText(bg: string): string {
  return contrastRatio(bg, '#ffffff') >= contrastRatio(bg, '#1b1b1b') ? '#ffffff' : '#1b1b1b';
}

/** Blends the color toward white. `pct` is the share of the original color (0–100). */
export function mixWithWhite(hex: string, pct: number): string {
  const { r, g, b } = hexToRgb(hex);
  const p = Math.max(0, Math.min(100, pct)) / 100;
  return rgbToHex(r * p + 255 * (1 - p), g * p + 255 * (1 - p), b * p + 255 * (1 - p));
}

/**
 * Returns a version of the color that reaches at least `minRatio` contrast
 * against `bg` (default white), darkening (or lightening, for dark backgrounds)
 * in small steps. Used when the accent is drawn as text.
 */
export function ensureReadable(hex: string, bg = '#ffffff', minRatio = 4.5): string {
  const base = normalizeHex(hex) || DEFAULT_ACCENT;
  if (contrastRatio(base, bg) >= minRatio) return base;
  const darken = relativeLuminance(bg) > 0.5;
  let { r, g, b } = hexToRgb(base);
  for (let i = 0; i < 40; i++) {
    const f = darken ? 0.92 : 1.08;
    r = darken ? r * f : r + (255 - r) * (f - 1) * 2;
    g = darken ? g * f : g + (255 - g) * (f - 1) * 2;
    b = darken ? b * f : b + (255 - b) * (f - 1) * 2;
    const candidate = rgbToHex(r, g, b);
    if (contrastRatio(candidate, bg) >= minRatio) return candidate;
  }
  return darken ? '#1b1b1b' : '#ffffff';
}

export interface IAccentPalette {
  /** The accent itself, for fills and borders. */
  accent: string;
  /** Text color to place on top of an accent fill. */
  onAccent: string;
  /** Accent adjusted to be readable as text on white. */
  text: string;
  /** Light tint of the accent for subtle backgrounds (badges, hovers). */
  soft: string;
}

export function getAccentPalette(accentColor: string | undefined): IAccentPalette {
  const accent = normalizeHex(accentColor) || DEFAULT_ACCENT;
  return {
    accent,
    onAccent: getContrastText(accent),
    text: ensureReadable(accent),
    soft: mixWithWhite(accent, 12),
  };
}

/**
 * CSS custom properties for the custom theme. Set these inline on a theme
 * container; SCSS reads them with var(--soc-custom-accent, #0078d4) etc.
 */
export function getAccentCssVars(accentColor: string | undefined): React.CSSProperties {
  const p = getAccentPalette(accentColor);
  return {
    '--soc-custom-accent': p.accent,
    '--soc-custom-accent-contrast': p.onAccent,
    '--soc-custom-accent-text': p.text,
    '--soc-custom-accent-soft': p.soft,
  } as React.CSSProperties;
}

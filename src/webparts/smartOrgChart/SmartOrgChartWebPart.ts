import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneCustomFieldProps,
  IPropertyPaneField,
  PropertyPaneFieldType,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneChoiceGroup,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';

import { SmartOrgChart } from './components/SmartOrgChart';
import { ISmartOrgChartProps, OrgChartTheme } from './components/ISmartOrgChartProps';
import { DEMO_CEO_EMAIL } from '../../services/MockGraphService';
import { ICustomAttributeConfig, isValidCustomAttributeField } from '../../services/GraphService';
import { DEFAULT_ACCENT, normalizeHex, hexToRgb, rgbToHex, contrastRatio, getContrastText, ensureReadable } from './components/colorUtils';

type ChangeCallback = ((targetProperty?: string, newValue?: any) => void) | undefined;

/**
 * State kept for a custom property pane field whose DOM is built once.
 * SPFx calls onRender again on every pane refresh; rebuilding the inputs then
 * would destroy focus and any in-progress drag, so later calls only re-sync.
 */
interface ICustomFieldHandle {
  changeCallback: ChangeCallback;
  sync: () => void;
  dispose: () => void;
}

// Property pane fields whose value changes the pane's structure (conditional
// fields) or the value of another field. Only these need a pane refresh.
const PANE_STRUCTURE_PROPS = ['theme', 'useDemoData'];

// Cap on admin-configured custom attributes — keeps the pane usable and bounds
// the extra columns/fields the Directory and Org Chart would otherwise render.
const MAX_CUSTOM_ATTRIBUTES = 10;

export interface ISmartOrgChartWebPartProps {
  defaultView: 'directory' | 'orgchart';
  topLevelUser: string;
  levelsBelow: number;
  pageSize: number;
  useDemoData: boolean;
  // Branding
  companyName: string;
  logoUrl: string;
  // Visual style (admin-controlled)
  theme: OrgChartTheme;
  accentColor: string;
  defaultLayout: 'drill' | 'vertical' | 'horizontal';
  // Feature flags
  enableFindMe: boolean;
  enableLayoutToggle: boolean;
  enableStats: boolean;
  enableDeptFilter: boolean;
  enableUserFilter: boolean;
  defaultZoom: number;
  defaultFontScale: number;
  // Data
  dataSource: 'auto' | 'graph' | 'search';
  dottedLineAttribute: string;
  // Admin-configured Entra ID attributes to surface in the Directory / Org Chart.
  // Graph only — SharePoint Search has no generic way to read them.
  customAttributes: ICustomAttributeConfig[];
  // User filters
  excludedAccounts: string;
  restrictToTenantDomain: boolean;
  hideGuestUsers: boolean;
  hideDisabledAccounts: boolean;
  hideNoJobTitle: boolean;
  hideNoDepartment: boolean;
}

export default class SmartOrgChartWebPart extends BaseClientSideWebPart<ISmartOrgChartWebPartProps> {
  // SPFx can call render (and onThemeChanged, if it is ever overridden) before
  // onInit has run, while this.properties/this.context are not yet usable.
  // Any future onThemeChanged override must also return early until this is set.
  private _isInitialized = false;
  private _customFields: WeakMap<HTMLElement, ICustomFieldHandle> = new WeakMap();

  protected async onInit(): Promise<void> {
    const p = this.properties;
    // Feature flag defaults
    if (p.enableFindMe       === undefined) p.enableFindMe       = true;
    if (p.enableLayoutToggle === undefined) p.enableLayoutToggle = true;
    if (p.enableStats        === undefined) p.enableStats        = true;
    if (p.enableDeptFilter   === undefined) p.enableDeptFilter   = true;
    if (p.enableUserFilter   === undefined) p.enableUserFilter   = true;
    // Visual style defaults
    if (p.theme         === undefined) p.theme         = 'modern';
    if (p.accentColor   === undefined) p.accentColor   = '#0078d4';
    if (p.defaultLayout === undefined) p.defaultLayout = 'drill';
    if (p.defaultZoom      === undefined) p.defaultZoom      = 0;
    if (p.defaultFontScale === undefined) p.defaultFontScale = 1;
    // Branding defaults
    if (p.companyName === undefined) p.companyName = '';
    if (p.logoUrl     === undefined) p.logoUrl     = '';
    // Data source default
    if (p.dataSource  === undefined) p.dataSource  = 'auto';
    if (p.dottedLineAttribute === undefined) p.dottedLineAttribute = '';
    if (p.customAttributes    === undefined) p.customAttributes    = [];
    // User filter defaults
    if (p.excludedAccounts       === undefined) p.excludedAccounts       = '';
    if (p.restrictToTenantDomain === undefined) p.restrictToTenantDomain = false;
    if (p.hideGuestUsers         === undefined) p.hideGuestUsers         = true;
    if (p.hideDisabledAccounts   === undefined) p.hideDisabledAccounts   = true;
    if (p.hideNoJobTitle         === undefined) p.hideNoJobTitle         = false;
    if (p.hideNoDepartment       === undefined) p.hideNoDepartment       = false;
    await super.onInit();
    this._isInitialized = true;
  }

  // The pane is reactive, so SPFx re-renders the web part after every change
  // on its own — no explicit render() here. The pane itself is only refreshed
  // when the change affects which fields are shown or another field's value.
  protected onPropertyPaneFieldChanged(propertyPath: string, _oldValue: any, newValue: any): void {
    // Give first-time editors a working example immediately: turning on Use
    // Demo Data fills in the Top-Level User field with the demo CEO, unless
    // the admin has already configured a real one.
    if (propertyPath === 'useDemoData' && newValue === true && !this.properties.topLevelUser) {
      this.properties.topLevelUser = DEMO_CEO_EMAIL;
    }
    if (PANE_STRUCTURE_PROPS.indexOf(propertyPath) >= 0) {
      this.context.propertyPane.refresh();
    }
  }

  public render(): void {
    // Must be the first statement: nothing may read this.properties before onInit.
    if (!this._isInitialized) return;
    const element: React.ReactElement<ISmartOrgChartProps> = React.createElement(SmartOrgChart, {
      context: this.context,
      defaultView: this.properties.defaultView || 'directory',
      topLevelUser: this.properties.topLevelUser || '',
      levelsBelow: this.properties.levelsBelow || 3,
      pageSize: this.properties.pageSize || 50,
      useDemoData: this.properties.useDemoData || false,
      companyName: this.properties.companyName || '',
      logoUrl: this.properties.logoUrl || '',
      theme: this.properties.theme || 'modern',
      accentColor: this.properties.accentColor || '#0078d4',
      defaultLayout: this.properties.defaultLayout || 'drill',
      defaultZoom: this.properties.defaultZoom ?? 0,
      defaultFontScale: this.properties.defaultFontScale || 1,
      enableFindMe: this.properties.enableFindMe !== false,
      enableLayoutToggle: this.properties.enableLayoutToggle !== false,
      enableStats: this.properties.enableStats !== false,
      enableDeptFilter: this.properties.enableDeptFilter !== false,
      enableUserFilter: this.properties.enableUserFilter !== false,
      dataSource: this.properties.dataSource || 'auto',
      dottedLineAttribute: this.properties.dottedLineAttribute || '',
      customAttributes: this.properties.customAttributes || [],
      excludedAccounts:       this.properties.excludedAccounts       || '',
      restrictToTenantDomain: this.properties.restrictToTenantDomain || false,
      hideGuestUsers:         this.properties.hideGuestUsers         ?? true,
      hideDisabledAccounts:   this.properties.hideDisabledAccounts   ?? true,
      hideNoJobTitle:         this.properties.hideNoJobTitle         || false,
      hideNoDepartment:       this.properties.hideNoDepartment       || false,
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /* ── Custom property pane fields ───────────────────────────── */

  private _disposeCustomField(elem: HTMLElement): void {
    const handle = this._customFields.get(elem);
    if (handle) handle.dispose();
    this._customFields.delete(elem);
    elem.innerHTML = '';
  }

  /** Returns true when the field was already built (its values were re-synced). */
  private _resyncCustomField(elem: HTMLElement, changeCallback: ChangeCallback): boolean {
    const handle = this._customFields.get(elem);
    if (!handle) return false;
    handle.changeCallback = changeCallback;
    handle.sync();
    return true;
  }

  private _fieldId(name: string): string {
    return `soc-${name}-${this.context.instanceId}`;
  }

  private _renderAccentPicker(elem: HTMLElement, changeCallback: ChangeCallback): void {
    if (this._resyncCustomField(elem, changeCallback)) return;
    elem.innerHTML = '';

    const mainId = this._fieldId('accent-hex');
    const errorId = this._fieldId('accent-error');
    const contrastId = this._fieldId('accent-contrast');
    const storedHex = (): string => normalizeHex(this.properties.accentColor) || DEFAULT_ACCENT;

    const label = document.createElement('label');
    label.textContent = 'Custom Accent Color';
    label.htmlFor = mainId;
    label.style.cssText = 'display:block;font-weight:600;font-size:14px;color:#323130;margin-bottom:8px;margin-top:4px;';
    elem.appendChild(label);

    const row1 = document.createElement('div');
    row1.style.cssText = 'display:flex;align-items:center;gap:8px;margin-bottom:8px;';

    const colorInput = document.createElement('input');
    colorInput.type = 'color';
    colorInput.id = this._fieldId('accent-swatch');
    colorInput.setAttribute('aria-label', 'Accent color picker');
    colorInput.style.cssText = 'width:40px;height:32px;border:1px solid #c8c6c4;border-radius:2px;padding:1px;cursor:pointer;';

    const hexLabel = document.createElement('span');
    hexLabel.textContent = 'Hex';
    hexLabel.setAttribute('aria-hidden', 'true');
    hexLabel.style.cssText = 'font-size:12px;color:#605e5c;font-weight:600;';

    const hexInput = document.createElement('input');
    hexInput.type = 'text';
    hexInput.id = mainId;
    hexInput.maxLength = 9;
    hexInput.placeholder = DEFAULT_ACCENT;
    hexInput.setAttribute('aria-label', 'Accent color hex value');
    hexInput.setAttribute('aria-describedby', `${errorId} ${contrastId}`);
    hexInput.spellcheck = false;
    hexInput.style.cssText = 'width:90px;padding:4px 8px;border:1px solid #c8c6c4;border-radius:2px;font-size:14px;font-family:monospace;letter-spacing:0.5px;';

    row1.appendChild(colorInput);
    row1.appendChild(hexLabel);
    row1.appendChild(hexInput);
    elem.appendChild(row1);

    const errorMsg = document.createElement('div');
    errorMsg.id = errorId;
    errorMsg.setAttribute('role', 'alert');
    errorMsg.style.cssText = 'font-size:12px;color:#a4262c;margin:-4px 0 8px;display:none;';
    elem.appendChild(errorMsg);

    const row2 = document.createElement('div');
    row2.style.cssText = 'display:flex;align-items:center;gap:6px;margin-bottom:4px;';

    const rgbSpan = document.createElement('span');
    rgbSpan.textContent = 'RGB';
    rgbSpan.setAttribute('aria-hidden', 'true');
    rgbSpan.style.cssText = 'font-size:12px;color:#605e5c;font-weight:600;min-width:28px;';
    row2.appendChild(rgbSpan);

    const makeChannel = (ch: string, name: string): HTMLInputElement => {
      const inp = document.createElement('input');
      inp.id = this._fieldId(`accent-${ch.toLowerCase()}`);
      const lbl = document.createElement('label');
      lbl.textContent = ch;
      lbl.htmlFor = inp.id;
      lbl.style.cssText = 'font-size:12px;color:#605e5c;';
      inp.type = 'number'; inp.min = '0'; inp.max = '255';
      inp.setAttribute('aria-label', `${name} channel (0 to 255)`);
      inp.style.cssText = 'width:52px;padding:3px 4px;border:1px solid #c8c6c4;border-radius:2px;font-size:13px;text-align:center;';
      row2.appendChild(lbl);
      row2.appendChild(inp);
      return inp;
    };

    const rInput = makeChannel('R', 'Red');
    const gInput = makeChannel('G', 'Green');
    const bInput = makeChannel('B', 'Blue');
    elem.appendChild(row2);

    // Live contrast indicator — small and unobtrusive
    const contrast = document.createElement('div');
    contrast.id = contrastId;
    contrast.setAttribute('aria-live', 'polite');
    contrast.style.cssText = 'font-size:12px;color:#605e5c;margin-top:6px;line-height:1.4;';
    elem.appendChild(contrast);

    const showError = (msg: string): void => {
      hexInput.setAttribute('aria-invalid', 'true');
      hexInput.style.borderColor = '#a4262c';
      errorMsg.textContent = msg;
      errorMsg.style.display = 'block';
    };
    const clearError = (): void => {
      hexInput.removeAttribute('aria-invalid');
      hexInput.style.borderColor = '#c8c6c4';
      errorMsg.textContent = '';
      errorMsg.style.display = 'none';
    };

    const updateContrast = (hex: string): void => {
      // Contrast is symmetric: white text on the accent and the accent as
      // text on white have the same ratio, so one figure covers both.
      const ratio = contrastRatio('#ffffff', hex);
      const passes = ratio >= 4.5 ? 'passes AA' : ratio >= 3 ? 'large text only' : 'low contrast';
      const notes: string[] = [`Contrast with white: ${ratio.toFixed(2)}:1 (${passes})`];
      if (getContrastText(hex) !== '#ffffff') {
        notes.push('Text on this color will use dark text for readability.');
      }
      const readable = ensureReadable(hex);
      if (readable !== hex) {
        notes.push(`Titles and links will use a darker shade (${readable}) for readability.`);
      }
      contrast.textContent = notes.join(' ');
    };

    const showValues = (hex: string, skip?: HTMLInputElement): void => {
      const rgb = hexToRgb(hex);
      if (skip !== colorInput) colorInput.value = hex;
      if (skip !== hexInput) hexInput.value = hex;
      if (skip !== rInput) rInput.value = String(rgb.r);
      if (skip !== gInput) gInput.value = String(rgb.g);
      if (skip !== bInput) bInput.value = String(rgb.b);
      updateContrast(hex);
    };

    // Commit to the web part — rAF-throttled so dragging the native color
    // picker re-renders the web part at most once per frame
    let rafId = 0;
    let pending: string | null = null;

    const handle: ICustomFieldHandle = {
      changeCallback,
      // Re-sync from the stored value without clobbering an input that has focus
      sync: () => {
        const active = document.activeElement as HTMLInputElement | null;
        const focused = active && elem.contains(active) ? active : undefined;
        showValues(storedHex(), focused);
      },
      dispose: () => { if (rafId !== 0) window.cancelAnimationFrame(rafId); rafId = 0; },
    };

    const flush = (): void => {
      rafId = 0;
      const value = pending;
      pending = null;
      if (value !== null && value !== this.properties.accentColor && handle.changeCallback) {
        // changeCallback sets this.properties.accentColor itself, so oldValue stays accurate
        handle.changeCallback('accentColor', value);
      }
    };
    const commit = (hex: string, immediate: boolean): void => {
      pending = hex;
      if (immediate) {
        if (rafId !== 0) window.cancelAnimationFrame(rafId);
        flush();
      } else if (rafId === 0) {
        rafId = window.requestAnimationFrame(flush);
      }
    };

    colorInput.addEventListener('input', () => {
      const hex = normalizeHex(colorInput.value);
      if (!hex) return;
      clearError();
      showValues(hex, colorInput);
      commit(hex, false);
    });
    colorInput.addEventListener('change', () => {
      const hex = normalizeHex(colorInput.value);
      if (hex) commit(hex, true);
    });

    // Set after an invalid entry is reverted, so the follow-up blur/change for
    // the same edit doesn't immediately hide the error message
    let reverted = false;
    const applyHexText = (): void => {
      if (reverted) return;
      const hex = normalizeHex(hexInput.value);
      if (!hex) {
        // Revert to the stored value rather than leaving an unusable entry behind
        const typed = hexInput.value.trim();
        showValues(storedHex());
        if (typed) {
          showError(`"${typed}" is not a valid hex color. Use a value like #0078d4 or #07d.`);
          reverted = true;
        } else {
          clearError();
        }
        return;
      }
      clearError();
      showValues(hex);
      commit(hex, true);
    };
    hexInput.addEventListener('input', () => { reverted = false; });
    hexInput.addEventListener('change', applyHexText);
    hexInput.addEventListener('blur', applyHexText);
    hexInput.addEventListener('keydown', (e: KeyboardEvent) => { if (e.key === 'Enter') applyHexText(); });

    const syncFromRgb = (): void => {
      const r = parseInt(rInput.value, 10);
      const g = parseInt(gInput.value, 10);
      const b = parseInt(bInput.value, 10);
      if (isNaN(r) || isNaN(g) || isNaN(b)) { showValues(storedHex()); return; }
      const hex = rgbToHex(r, g, b);
      clearError();
      showValues(hex);
      commit(hex, true);
    };
    rInput.addEventListener('change', syncFromRgb);
    gInput.addEventListener('change', syncFromRgb);
    bInput.addEventListener('change', syncFromRgb);

    this._customFields.set(elem, handle);
    showValues(storedHex());
  }

  private _renderPageSizeField(elem: HTMLElement, changeCallback: ChangeCallback): void {
    if (this._resyncCustomField(elem, changeCallback)) return;
    elem.innerHTML = '';

    const sliderId = this._fieldId('pagesize-slider');
    const clamp = (val: number): number => Math.max(10, Math.min(200, isNaN(val) ? 50 : val));
    const stored = (): number => clamp(this.properties.pageSize || 50);

    const label = document.createElement('label');
    label.textContent = 'Max employees per page';
    label.htmlFor = sliderId;
    label.style.cssText = 'display:block;font-weight:600;font-size:14px;color:#323130;margin-bottom:8px;';

    const row = document.createElement('div');
    row.style.cssText = 'display:flex;align-items:center;gap:12px;';

    const slider = document.createElement('input');
    slider.type = 'range';
    slider.id = sliderId;
    slider.min = '10';
    slider.max = '200';
    slider.step = '1';
    slider.style.cssText = 'flex:1;accent-color:#0078d4;';

    const numInput = document.createElement('input');
    numInput.type = 'number';
    numInput.id = this._fieldId('pagesize-number');
    numInput.min = '10';
    numInput.max = '200';
    numInput.setAttribute('aria-label', 'Max employees per page (10 to 200)');
    numInput.style.cssText = 'width:64px;padding:4px 6px;border:1px solid #c8c6c4;border-radius:2px;font-size:14px;text-align:center;';

    let rafId = 0;
    let pending: number | null = null;

    const handle: ICustomFieldHandle = {
      changeCallback,
      sync: () => {
        const v = String(stored());
        if (document.activeElement !== slider) slider.value = v;
        if (document.activeElement !== numInput) numInput.value = v;
      },
      dispose: () => { if (rafId !== 0) window.cancelAnimationFrame(rafId); rafId = 0; },
    };

    const flush = (): void => {
      rafId = 0;
      const v = pending;
      pending = null;
      if (v !== null && v !== this.properties.pageSize && handle.changeCallback) {
        handle.changeCallback('pageSize', v);
      }
    };

    const set = (val: number, immediate: boolean): void => {
      const v = clamp(val);
      slider.value = String(v);
      numInput.value = String(v);
      pending = v;
      if (immediate) {
        if (rafId !== 0) window.cancelAnimationFrame(rafId);
        flush();
      } else if (rafId === 0) {
        rafId = window.requestAnimationFrame(flush);
      }
    };

    slider.addEventListener('input', () => set(parseInt(slider.value, 10), false));
    slider.addEventListener('change', () => set(parseInt(slider.value, 10), true));
    numInput.addEventListener('change', () => set(parseInt(numInput.value, 10), true));

    row.appendChild(slider);
    row.appendChild(numInput);
    elem.appendChild(label);
    elem.appendChild(row);

    this._customFields.set(elem, handle);
    handle.sync();
  }

  private _renderCustomAttributesField(elem: HTMLElement, changeCallback: ChangeCallback): void {
    if (this._resyncCustomField(elem, changeCallback)) return;
    elem.innerHTML = '';

    interface IRowRefs {
      row: HTMLElement;
      fieldInput: HTMLInputElement;
      fieldError: HTMLDivElement;
      labelInput: HTMLInputElement;
      dirCheckbox: HTMLInputElement;
      chartCheckbox: HTMLInputElement;
    }

    const heading = document.createElement('div');
    heading.textContent = 'Manage custom attributes';
    heading.style.cssText = 'display:block;font-weight:600;font-size:14px;color:#323130;margin-bottom:8px;margin-top:4px;';
    elem.appendChild(heading);

    const rowsContainer = document.createElement('div');
    rowsContainer.style.cssText = 'display:flex;flex-direction:column;gap:8px;';
    elem.appendChild(rowsContainer);

    const emptyMsg = document.createElement('div');
    emptyMsg.textContent = 'No custom attributes configured.';
    emptyMsg.style.cssText = 'font-size:12px;color:#605e5c;font-style:italic;margin-bottom:8px;display:none;';
    elem.appendChild(emptyMsg);

    const addBtn = document.createElement('button');
    addBtn.type = 'button';
    addBtn.textContent = '+ Add attribute';
    addBtn.style.cssText = 'padding:6px 14px;border:1px solid #c8c6c4;border-radius:2px;background:#fff;cursor:pointer;font-size:13px;color:#323130;';
    elem.appendChild(addBtn);

    const rowRefs: Map<string, IRowRefs> = new Map();
    const getList = (): ICustomAttributeConfig[] => (this.properties.customAttributes || []).slice();
    const findIndex = (list: ICustomAttributeConfig[], id: string): number => list.findIndex(c => c.id === id);

    const updateEmptyState = (): void => {
      emptyMsg.style.display = rowsContainer.children.length === 0 ? 'block' : 'none';
    };
    const updateAddButtonState = (): void => {
      const atCap = getList().length >= MAX_CUSTOM_ATTRIBUTES;
      addBtn.disabled = atCap;
      addBtn.style.opacity = atCap ? '0.5' : '1';
      addBtn.style.cursor = atCap ? 'default' : 'pointer';
      addBtn.title = atCap ? `Maximum of ${MAX_CUSTOM_ATTRIBUTES} custom attributes` : '';
    };

    const applyRowValues = (refs: IRowRefs, cfg: ICustomAttributeConfig, skipFocused?: boolean): void => {
      const active = document.activeElement;
      if (!(skipFocused && active === refs.fieldInput)) {
        refs.fieldInput.value = cfg.graphField;
        if (cfg.graphField && !isValidCustomAttributeField(cfg.graphField)) {
          refs.fieldInput.setAttribute('aria-invalid', 'true');
          refs.fieldInput.style.borderColor = '#a4262c';
          refs.fieldError.textContent = 'Use employeeId, companyName, city, ... or extensionAttribute1-15';
          refs.fieldError.style.display = 'block';
        } else {
          refs.fieldInput.removeAttribute('aria-invalid');
          refs.fieldInput.style.borderColor = '#c8c6c4';
          refs.fieldError.textContent = '';
          refs.fieldError.style.display = 'none';
        }
      }
      if (!(skipFocused && active === refs.labelInput)) {
        refs.labelInput.value = cfg.label;
      }
      refs.dirCheckbox.checked = cfg.showInDirectory;
      refs.chartCheckbox.checked = cfg.showInOrgChart;
    };

    const handle: ICustomFieldHandle = {
      changeCallback,
      sync: () => {
        const list = getList();
        rowRefs.forEach((refs, id) => {
          const cfg = list[findIndex(list, id)];
          if (cfg) applyRowValues(refs, cfg, true);
        });
        updateAddButtonState();
        updateEmptyState();
      },
      dispose: () => { /* no timers/rAF to clean up */ },
    };

    const commit = (list: ICustomAttributeConfig[]): void => {
      if (handle.changeCallback) handle.changeCallback('customAttributes', list);
      updateAddButtonState();
    };

    const buildRow = (cfg: ICustomAttributeConfig, ordinal: number): IRowRefs => {
      const row = document.createElement('div');
      row.style.cssText = 'border:1px solid #edebe9;border-radius:4px;padding:8px;display:flex;flex-direction:column;gap:6px;';

      const line1 = document.createElement('div');
      line1.style.cssText = 'display:flex;align-items:flex-start;gap:6px;';

      const fieldWrap = document.createElement('div');
      fieldWrap.style.cssText = 'flex:1;min-width:0;';
      const fieldInputId = this._fieldId(`customattr-field-${cfg.id}`);
      const fieldLabel = document.createElement('label');
      fieldLabel.htmlFor = fieldInputId;
      fieldLabel.textContent = 'Attribute name';
      fieldLabel.style.cssText = 'display:block;font-size:11px;color:#605e5c;margin-bottom:2px;';
      const fieldInput = document.createElement('input');
      fieldInput.type = 'text';
      fieldInput.id = fieldInputId;
      fieldInput.placeholder = 'e.g. companyName';
      fieldInput.spellcheck = false;
      fieldInput.setAttribute('aria-label', 'Attribute name');
      fieldInput.style.cssText = 'width:100%;box-sizing:border-box;padding:4px 6px;border:1px solid #c8c6c4;border-radius:2px;font-size:13px;';
      fieldWrap.appendChild(fieldLabel);
      fieldWrap.appendChild(fieldInput);
      const fieldError = document.createElement('div');
      fieldError.setAttribute('role', 'alert');
      fieldError.style.cssText = 'font-size:11px;color:#a4262c;margin-top:2px;display:none;';
      fieldWrap.appendChild(fieldError);

      const labelWrap = document.createElement('div');
      labelWrap.style.cssText = 'flex:1;min-width:0;';
      const labelInputId = this._fieldId(`customattr-label-${cfg.id}`);
      const labelLabel = document.createElement('label');
      labelLabel.htmlFor = labelInputId;
      labelLabel.textContent = 'Display label';
      labelLabel.style.cssText = 'display:block;font-size:11px;color:#605e5c;margin-bottom:2px;';
      const labelInput = document.createElement('input');
      labelInput.type = 'text';
      labelInput.id = labelInputId;
      labelInput.placeholder = 'e.g. Cost Center';
      labelInput.setAttribute('aria-label', 'Display label');
      labelInput.style.cssText = 'width:100%;box-sizing:border-box;padding:4px 6px;border:1px solid #c8c6c4;border-radius:2px;font-size:13px;';
      labelWrap.appendChild(labelLabel);
      labelWrap.appendChild(labelInput);

      const removeBtn = document.createElement('button');
      removeBtn.type = 'button';
      removeBtn.textContent = '×';
      removeBtn.title = 'Remove attribute';
      removeBtn.setAttribute('aria-label', `Remove attribute ${ordinal + 1}`);
      removeBtn.style.cssText = 'flex-shrink:0;width:26px;height:26px;border:1px solid #c8c6c4;border-radius:2px;background:#fff;cursor:pointer;font-size:14px;line-height:1;color:#605e5c;margin-top:16px;';

      line1.appendChild(fieldWrap);
      line1.appendChild(labelWrap);
      line1.appendChild(removeBtn);

      const line2 = document.createElement('div');
      line2.style.cssText = 'display:flex;align-items:center;gap:14px;';

      const makeCheckbox = (text: string, idSuffix: string): HTMLInputElement => {
        const wrap = document.createElement('label');
        wrap.style.cssText = 'display:flex;align-items:center;gap:4px;font-size:12px;color:#323130;cursor:pointer;';
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.id = this._fieldId(`customattr-${idSuffix}-${cfg.id}`);
        const span = document.createElement('span');
        span.textContent = text;
        wrap.appendChild(cb);
        wrap.appendChild(span);
        line2.appendChild(wrap);
        return cb;
      };
      const dirCheckbox = makeCheckbox('Show in Directory', 'dir');
      const chartCheckbox = makeCheckbox('Show in Org Chart', 'chart');

      row.appendChild(line1);
      row.appendChild(line2);

      const refs: IRowRefs = { row, fieldInput, fieldError, labelInput, dirCheckbox, chartCheckbox };

      const commitFieldChange = (): void => {
        const value = fieldInput.value.trim();
        const list = getList();
        const idx = findIndex(list, cfg.id);
        if (idx < 0) return;
        applyRowValues(refs, { ...list[idx], graphField: value });
        if (list[idx].graphField !== value) {
          list[idx] = { ...list[idx], graphField: value };
          commit(list);
        }
      };
      fieldInput.addEventListener('change', commitFieldChange);
      fieldInput.addEventListener('blur', commitFieldChange);

      const commitLabelChange = (): void => {
        const value = labelInput.value;
        const list = getList();
        const idx = findIndex(list, cfg.id);
        if (idx < 0 || list[idx].label === value) return;
        list[idx] = { ...list[idx], label: value };
        commit(list);
      };
      labelInput.addEventListener('change', commitLabelChange);
      labelInput.addEventListener('blur', commitLabelChange);

      dirCheckbox.addEventListener('change', () => {
        const list = getList();
        const idx = findIndex(list, cfg.id);
        if (idx < 0) return;
        list[idx] = { ...list[idx], showInDirectory: dirCheckbox.checked };
        commit(list);
      });
      chartCheckbox.addEventListener('change', () => {
        const list = getList();
        const idx = findIndex(list, cfg.id);
        if (idx < 0) return;
        list[idx] = { ...list[idx], showInOrgChart: chartCheckbox.checked };
        commit(list);
      });

      removeBtn.addEventListener('click', () => {
        const list = getList().filter(c => c.id !== cfg.id);
        rowRefs.delete(cfg.id);
        row.remove();
        commit(list);
        updateEmptyState();
      });

      return refs;
    };

    const initialList = getList();
    initialList.forEach((cfg, idx) => {
      const refs = buildRow(cfg, idx);
      rowRefs.set(cfg.id, refs);
      rowsContainer.appendChild(refs.row);
      applyRowValues(refs, cfg);
    });
    updateEmptyState();
    updateAddButtonState();

    addBtn.addEventListener('click', () => {
      const list = getList();
      if (list.length >= MAX_CUSTOM_ATTRIBUTES) return;
      const newCfg: ICustomAttributeConfig = {
        id: 'attr-' + Date.now() + '-' + Math.random().toString(36).slice(2),
        graphField: '',
        label: '',
        showInDirectory: true,
        showInOrgChart: true,
      };
      const newList = [...list, newCfg];
      const refs = buildRow(newCfg, newList.length - 1);
      rowRefs.set(newCfg.id, refs);
      rowsContainer.appendChild(refs.row);
      applyRowValues(refs, newCfg);
      updateEmptyState();
      commit(newList);
      refs.fieldInput.focus();
    });

    this._customFields.set(elem, handle);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Smart Org Chart Settings' },
          groups: [
            {
              groupName: 'General',
              groupFields: [
                PropertyPaneChoiceGroup('defaultView', {
                  label: 'Default View',
                  options: [
                    { key: 'directory', text: 'Employee Directory', iconProps: { officeFabricIconFontName: 'People' } },
                    { key: 'orgchart', text: 'Org Chart', iconProps: { officeFabricIconFontName: 'Org' } }
                  ]
                })
              ]
            },
            {
              groupName: 'Branding',
              groupFields: [
                PropertyPaneTextField('companyName', {
                  label: 'App Title',
                  placeholder: 'Contoso Org Chart',
                  description: 'Displayed in the header bar alongside the view name'
                }),
                PropertyPaneTextField('logoUrl', {
                  label: 'Logo URL',
                  placeholder: 'https://contoso.sharepoint.com/sites/yoursite/SiteAssets/logo.png',
                  description: 'Full URL to a PNG/SVG/JPG image. Tip: open the image file in your browser and copy the address bar URL. "Copy link" sharing URLs will not work.'
                }),
              ]
            },
            {
              groupName: 'Visual Style',
              groupFields: [
                PropertyPaneChoiceGroup('theme', {
                  label: 'Chart Theme',
                  options: [
                    { key: 'modern',    text: 'Modern — dept colors on white',  iconProps: { officeFabricIconFontName: 'Color' } },
                    { key: 'minimal',   text: 'Minimal — flat & low contrast',  iconProps: { officeFabricIconFontName: 'CollapseMenu' } },
                    { key: 'corporate', text: 'Corporate — unified blue',        iconProps: { officeFabricIconFontName: 'Work' } },
                    { key: 'dark',      text: 'Dark — dark navy background',     iconProps: { officeFabricIconFontName: 'ClearNight' } },
                    { key: 'custom',    text: 'Custom — pick your own color',    iconProps: { officeFabricIconFontName: 'Eyedropper' } },
                  ]
                }),
                ...( this.properties.theme === 'custom' ? [
                  ({
                    type: PropertyPaneFieldType.Custom,
                    targetProperty: 'accentColor',
                    properties: {
                      key: 'accentColorPicker',
                      onRender: (elem: HTMLElement, _ctx: any, changeCallback: ChangeCallback) => {
                        this._renderAccentPicker(elem, changeCallback);
                      },
                      onDispose: (elem: HTMLElement) => { this._disposeCustomField(elem); }
                    }
                  } as IPropertyPaneField<IPropertyPaneCustomFieldProps>)
                ] : []),
                PropertyPaneDropdown('defaultFontScale', {
                  label: 'Default Font Size',
                  options: [
                    { key: 0.75, text: '75% — Extra Small' },
                    { key: 0.85, text: '85% — Small' },
                    { key: 1,    text: '100% — Normal' },
                    { key: 1.15, text: '115% — Large' },
                    { key: 1.3,  text: '130% — Extra Large' },
                    { key: 1.5,  text: '150% — XXL' },
                    { key: 1.75, text: '175% — XXXL' },
                  ],
                  selectedKey: this.properties.defaultFontScale || 1,
                }),
                PropertyPaneChoiceGroup('defaultLayout', {
                  label: 'Default Org Chart Layout',
                  options: [
                    { key: 'drill',      text: 'Drill-Down',    iconProps: { officeFabricIconFontName: 'Org' } },
                    { key: 'vertical',   text: 'Top Down',      iconProps: { officeFabricIconFontName: 'Down' } },
                    { key: 'horizontal', text: 'Left to Right', iconProps: { officeFabricIconFontName: 'Forward' } },
                  ]
                }),
              ]
            },
            {
              groupName: 'Data Source',
              groupFields: [
                PropertyPaneChoiceGroup('dataSource', {
                  label: 'Where to load user & org data from',
                  options: [
                    {
                      key: 'auto',
                      text: 'Auto — Graph API, fall back to SharePoint Search',
                      iconProps: { officeFabricIconFontName: 'AutoEnhanceOn' }
                    },
                    {
                      key: 'graph',
                      text: 'Graph API — live Azure AD data (no indexing delay)',
                      iconProps: { officeFabricIconFontName: 'AzureLogo' }
                    },
                    {
                      key: 'search',
                      text: 'SharePoint Search — legacy behavior',
                      iconProps: { officeFabricIconFontName: 'Search' }
                    }
                  ]
                }),
                PropertyPaneLabel('dataSource', {
                  text: 'Graph API is recommended. It reads directly from Azure Active Directory so new users and manager changes appear immediately. Requires Microsoft Graph permissions to be approved in the SharePoint App Catalog.'
                })
              ]
            },
            {
              groupName: 'Custom Attributes',
              groupFields: [
                PropertyPaneLabel('customAttributes', {
                  text: 'Custom attributes require Microsoft Graph (Data source: Graph or Auto). SharePoint Search has no generic way to read Entra ID attributes.'
                }),
                ({
                  type: PropertyPaneFieldType.Custom,
                  targetProperty: 'customAttributes',
                  properties: {
                    key: 'customAttributesField',
                    onRender: (elem: HTMLElement, _ctx: any, changeCallback: ChangeCallback) => {
                      this._renderCustomAttributesField(elem, changeCallback);
                    },
                    onDispose: (elem: HTMLElement) => { this._disposeCustomField(elem); }
                  }
                } as IPropertyPaneField<IPropertyPaneCustomFieldProps>)
              ]
            },
            {
              groupName: 'User Filters',
              groupFields: [
                PropertyPaneTextField('excludedAccounts', {
                  label: 'Exclude accounts',
                  placeholder: 'conf-room, noreply, admin@, Service Account',
                  description: 'Comma-separated words or patterns (case-insensitive). Any user whose display name, email, or UPN contains one of these will be hidden from all views.',
                  multiline: true,
                  rows: 3
                }),
                PropertyPaneToggle('restrictToTenantDomain', {
                  label: 'Only show tenant users',
                  onText: 'On — hides accounts with external email domains (e.g. gmail.com, hotmail.com)',
                  offText: 'Off — all users shown regardless of email domain'
                }),
                PropertyPaneToggle('hideGuestUsers', {
                  label: 'Hide Azure AD guest accounts',
                  onText: 'On — guest accounts hidden',
                  offText: 'Off — guest accounts visible (shown with Guest badge)'
                }),
                PropertyPaneToggle('hideDisabledAccounts', {
                  label: 'Hide disabled accounts',
                  onText: 'On — blocked sign-in accounts hidden',
                  offText: 'Off — disabled accounts visible (shown with Disabled badge)'
                }),
                PropertyPaneLabel('hideDisabledAccounts', {
                  text: 'Note: guest and disabled account detection requires the Graph API data source. SharePoint Search does not return this information, so these two filters (and the Guest/Disabled badges) have no effect when data comes from Search.'
                }),
                PropertyPaneToggle('hideNoJobTitle', {
                  label: 'Hide accounts without a job title',
                  onText: 'On — accounts with no job title hidden',
                  offText: 'Off — all accounts shown regardless of job title'
                }),
                PropertyPaneToggle('hideNoDepartment', {
                  label: 'Hide accounts without a department',
                  onText: 'On — accounts with no department hidden',
                  offText: 'Off — all accounts shown regardless of department'
                }),
              ]
            },
            {
              groupName: 'Org Chart',
              groupFields: [
                PropertyPaneTextField('topLevelUser', {
                  label: 'Top-Level User (UPN or Email)',
                  placeholder: 'ceo@yourcompany.com',
                  description: 'The person shown at the root of the org chart'
                }),
                PropertyPaneSlider('levelsBelow', {
                  label: 'Levels to load below root',
                  min: 1,
                  max: 8,
                  value: 3,
                  showValue: true,
                  step: 1
                }),
                PropertyPaneTextField('dottedLineAttribute', {
                  label: 'Dotted-line manager attribute',
                  placeholder: 'extensionAttribute10',
                  description: 'Optional. Name of the Azure AD on-premises extension attribute (extensionAttribute1-15) that stores a secondary "dotted line" manager\'s email or UPN. Dotted-line relationships appear on profile cards. Requires the Graph API data source.'
                }),
                PropertyPaneDropdown('defaultZoom', {
                  label: 'Default Org Chart Zoom',
                  options: [
                    { key: 0,    text: 'Default (Auto-fit)' },
                    { key: 0.5,  text: '50%' },
                    { key: 0.75, text: '75%' },
                    { key: 1,    text: '100%' },
                    { key: 1.25, text: '125%' },
                    { key: 1.5,  text: '150%' },
                  ],
                  selectedKey: this.properties.defaultZoom ?? 0,
                })
              ]
            },
            {
              groupName: 'Org Chart Features',
              groupFields: [
                PropertyPaneLabel('enableFindMe', {
                  text: 'Show or hide toolbar buttons in the org chart view.'
                }),
                PropertyPaneToggle('enableFindMe', {
                  label: 'Find Me button',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('enableLayoutToggle', {
                  label: 'Layout toggle (drill / vertical / horizontal)',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('enableStats', {
                  label: 'Org stats bar',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('enableDeptFilter', {
                  label: 'Department filter',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('enableUserFilter', {
                  label: 'User type filter (members / guests)',
                  onText: 'Visible',
                  offText: 'Hidden'
                })
              ]
            },
            {
              groupName: 'Directory',
              groupFields: [
                ({
                  type: PropertyPaneFieldType.Custom,
                  targetProperty: 'pageSize',
                  properties: {
                    key: 'pageSizeField',
                    onRender: (elem: HTMLElement, _ctx: any, changeCallback: ChangeCallback) => {
                      this._renderPageSizeField(elem, changeCallback);
                    },
                    onDispose: (elem: HTMLElement) => { this._disposeCustomField(elem); }
                  }
                } as IPropertyPaneField<IPropertyPaneCustomFieldProps>),
                PropertyPaneLabel('pageSize', {
                  text: 'Card size, visible fields, and other display preferences are set per-user via the Settings gear in the app.'
                })
              ]
            },
            {
              groupName: 'Demo',
              groupFields: [
                PropertyPaneToggle('useDemoData', {
                  label: 'Use Demo Data',
                  onText: 'On — showing sample employees',
                  offText: 'Off — live Microsoft 365 data',
                  checked: false
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}

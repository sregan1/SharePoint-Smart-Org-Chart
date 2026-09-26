import { IGraphUser, IOrgNode } from './GraphService';
import { formatString } from '../webparts/smartOrgChart/components/localeUtils';
import * as strings from 'SmartOrgChartWebPartStrings';

/* ── Shared download / CSV helpers ── */

function downloadBlob(content: string, filename: string, mime: string): void {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  // Revoking synchronously can cancel the download in some browsers
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// International phone numbers ("+1 206 555 0100") start with '+' but can't
// carry a formula, so they're exempt from the injection guard below
const PHONE_PATTERN = /^\+[\d\s().-]+$/;

// Prefix cells that Excel would interpret as formulas (CSV injection mitigation)
function csvCell(c: string): string {
  const safe = /^[=+\-@\t\r]/.test(c) && !PHONE_PATTERN.test(c) ? `'${c}` : c;
  return `"${safe.replace(/"/g, '""')}"`;
}

function csvRow(cells: string[]): string {
  return cells.map(csvCell).join(',');
}

// UTF-8 BOM ensures Excel opens the file with correct character encoding
// (explicit escape — a literal BOM character is invisible and easily stripped)
const CSV_BOM = '\uFEFF';

export interface IDirectoryExportOptions {
  showEmail: boolean;
  showPhone: boolean;
  showDepartment: boolean;
  showOffice: boolean;
}

const BASE_STYLES = `
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Segoe UI, Arial, sans-serif; font-size: 11pt; color: #333; }
  h1 { font-size: 18pt; color: #0078d4; margin-bottom: 4px; }
  .subtitle { font-size: 9pt; color: #666; margin-bottom: 20px; }
  @page { margin: 15mm; }
  @media print {
    body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  }
`;

function openPrintWindow(html: string): void {
  const win = window.open('', '_blank', 'width=900,height=700');
  if (!win) { alert(strings.Chart_AllowPopupsForPdf); return; }
  win.document.write(html);
  win.document.close();
  win.focus();
  // After document.write the document may already be "complete", in which
  // case onload never fires — print directly with a beat for layout/fonts
  const doPrint = (): void => { try { win.print(); } catch { /* window closed */ } };
  if (win.document.readyState === 'complete') setTimeout(doPrint, 100);
  else win.onload = doPrint;
}

export function exportDirectoryToExcel(users: IGraphUser[], opts: IDirectoryExportOptions): void {
  if (users.length === 0) return;

  const cols: { label: string; get: (u: IGraphUser) => string }[] = [
    { label: strings.Export_ColumnName,     get: u => u.displayName || '' },
    { label: strings.Export_ColumnJobTitle, get: u => u.jobTitle || '' },
  ];
  if (opts.showDepartment) cols.push({ label: strings.Export_ColumnDepartment, get: u => u.department || '' });
  if (opts.showOffice)     cols.push({ label: strings.Export_ColumnOffice,     get: u => u.officeLocation || '' });
  if (opts.showEmail)      cols.push({ label: strings.Export_ColumnEmail,      get: u => u.mail || '' });
  if (opts.showPhone)      cols.push({ label: strings.Export_ColumnPhone,      get: u => u.mobilePhone || (u.businessPhones && u.businessPhones[0]) || '' });

  const lines = [
    csvRow(cols.map(c => c.label)),
    ...users.map(u => csvRow(cols.map(c => c.get(u)))),
  ];

  downloadBlob(
    CSV_BOM + lines.join('\r\n'),
    `employee-directory-${new Date().toISOString().slice(0, 10)}.csv`,
    'text/csv;charset=utf-8;'
  );
}

export function exportOrgChartToCsv(rootNode: IOrgNode): void {
  const rows: string[] = [
    csvRow([
      strings.Export_ColumnName, strings.Export_ColumnJobTitle, strings.Export_ColumnDepartment, strings.Export_ColumnOffice,
      strings.Export_ColumnEmail, strings.Export_ColumnPhone, strings.Export_ColumnManager, strings.Export_ColumnLevel,
    ]),
  ];
  const visit = (n: IOrgNode, managerName: string, depth: number): void => {
    const u = n.user;
    rows.push(csvRow([
      u.displayName || '',
      u.jobTitle || '',
      u.department || '',
      u.officeLocation || '',
      u.mail || '',
      u.mobilePhone || (u.businessPhones && u.businessPhones[0]) || '',
      managerName,
      String(depth),
    ]));
    n.directReports.forEach(c => visit(c, u.displayName || '', depth + 1));
  };
  visit(rootNode, '', 0);

  downloadBlob(
    CSV_BOM + rows.join('\r\n'),
    `org-chart-${new Date().toISOString().slice(0, 10)}.csv`,
    'text/csv;charset=utf-8;'
  );
}

export function exportOrgChartToPdf(rootNode: IOrgNode, note?: string, locale?: string): void {
  const lines: string[] = [];

  const renderNode = (node: IOrgNode, depth: number): void => {
    const indent = depth * 20;
    const isRoot = depth === 0;
    const weight = isRoot ? 'bold' : depth === 1 ? '600' : 'normal';
    const size = isRoot ? '13pt' : depth === 1 ? '11pt' : '10pt';
    const color = isRoot ? '#0078d4' : '#222';

    let meta = '';
    if (node.user.jobTitle || node.user.department) {
      const parts = [node.user.jobTitle, node.user.department].filter(Boolean);
      meta = `<div style="font-size:8.5pt;color:#666;margin-top:2px;">${escHtml(parts.join(' · '))}</div>`;
    }

    lines.push(
      `<div style="margin-left:${indent}px;margin-bottom:${isRoot ? 10 : 5}px;padding-left:${depth > 0 ? 10 : 0}px;border-left:${depth > 0 ? '2px solid #e0e0e0' : 'none'};">` +
      `<span style="font-weight:${weight};font-size:${size};color:${color};">${escHtml(node.user.displayName)}</span>` +
      meta +
      `</div>`
    );

    node.directReports.forEach(child => renderNode(child, depth + 1));
  };

  renderNode(rootNode, 0);

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${escHtml(strings.Export_PdfDocumentTitle)}</title>
<style>${BASE_STYLES}</style></head><body>
<h1>${escHtml(strings.Export_PdfHeading)}</h1>
<div class="subtitle">${escHtml(formatString(strings.Export_PdfRootLabel, { name: rootNode.user.displayName }))} &nbsp;·&nbsp; ${new Date().toLocaleDateString(locale)}${note ? ` &nbsp;·&nbsp; ${escHtml(note)}` : ''}</div>
${lines.join('')}
</body></html>`;

  openPrintWindow(html);
}

function escHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { SmartOrgChart } from '../src/webparts/smartOrgChart/components/SmartOrgChart';
import { MockCompanySize } from '../src/services/MockGraphService';
import { OrgChartTheme } from '../src/webparts/smartOrgChart/components/ISmartOrgChartProps';
import { ICustomAttributeConfig, IExternalPersonConfig } from '../src/services/GraphService';
import SmartOrgChartWebPart from '../src/webparts/smartOrgChart/SmartOrgChartWebPart';

// A representative admin configuration for the "Custom Attributes" feature —
// shown on directory cards/list and in the org chart's profile popup — so the
// demo harness actually demonstrates it instead of rendering an empty default.
const DEMO_CUSTOM_ATTRIBUTES: ICustomAttributeConfig[] = [
  { id: 'demo-employeeid', graphField: 'employeeId',          label: 'Employee ID', showInDirectory: true,  showInOrgChart: true  },
  { id: 'demo-city',       graphField: 'city',                label: 'City',        showInDirectory: true,  showInOrgChart: false },
  { id: 'demo-costcenter', graphField: 'extensionAttribute1',  label: 'Cost Center', showInDirectory: false, showInOrgChart: true  },
];

// SharePoint normally registers the Fabric icon font glyphs for us; the demo
// harness runs standalone, so without this every <Icon> renders as a blank
// square. Font files are served locally (see webpack.demo.config.js) since
// this harness has no network access.
initializeIcons('fonts/');

const p = new URLSearchParams(window.location.search);

const view     = (p.get('view')   || 'directory') as 'directory' | 'orgchart';
const layout   = (p.get('layout') || 'drill')     as 'drill' | 'vertical' | 'horizontal';
const theme    = (p.get('theme')  || 'modern')    as OrgChartTheme;
// Only meaningful when theme=custom; a distinct default so a custom-theme
// screenshot doesn't just look like the modern theme's accent color.
const accentColor = p.get('accentColor') || '#8764b8';
const sizeRaw  = p.get('mockSize');
const mockSize: MockCompanySize = sizeRaw === '500' ? 500 : sizeRaw === '1000' ? 1000 : 150;
const levelsBelow = p.has('levels') ? parseInt(p.get('levels') as string, 10) : 3;
const defaultZoom = p.has('zoom') ? parseFloat(p.get('zoom') as string) : 0;

// Persist the requested mock size so SmartOrgChart's readMockSize() picks it up.
try { localStorage.setItem('smartOrgChart_mockSize', String(mockSize)); } catch { /* ignore */ }
// Force the requested view so readCurrentView() doesn't restore a stale value.
try { localStorage.setItem('smartOrgChart_currentView', view); } catch { /* ignore */ }
// Force the requested layout so OrgChart doesn't restore a stale chartLayout from localStorage.
try {
  const chartState = JSON.parse(localStorage.getItem('smartOrgChart_chartState') || '{}');
  chartState.chartLayout = layout;
  // Clear saved navigation so tree layouts always start from the root.
  delete chartState.focusEmail;
  localStorage.setItem('smartOrgChart_chartState', JSON.stringify(chartState));
} catch { /* ignore */ }

const mockContext = {
  pageContext: {
    user: { email: 'demo@contoso.com' },
    web: { absoluteUrl: 'https://contoso.sharepoint.com/sites/demo' },
  },
  spHttpClient: {},
  msGraphClientFactory: {},
} as any;

// ?pane=1 shows a side panel driven by the web part's real property pane
// editor for External People, so it can be tried without a SharePoint
// workbench. Off by default so screenshots and navigation tests are unchanged.
const showPane = p.get('pane') === '1';
const LS_DEMO_EXTERNAL = 'smartOrgChart_demoExternalPeople';
const DEMO_EXTERNAL_PEOPLE: IExternalPersonConfig[] = [
  { id: 'demo-ext-1', name: 'Priya Natarajan', jobTitle: 'Executive Advisor (Contract)', managerEmail: 'a.chen@contoso.com' },
  { id: 'demo-ext-2', name: 'Marco Bellini',   jobTitle: 'Partner Account Manager',      managerEmail: 'd.williams@contoso.com' },
  { id: 'demo-ext-3', name: 'Open Position',   jobTitle: 'Senior Frontend Engineer',     managerEmail: 'e.rodriguez@contoso.com' },
];

function loadDemoExternalPeople(): IExternalPersonConfig[] {
  try {
    const stored = JSON.parse(localStorage.getItem(LS_DEMO_EXTERNAL) || 'null');
    if (Array.isArray(stored)) return stored;
  } catch { /* ignore */ }
  return DEMO_EXTERNAL_PEOPLE;
}

let externalPeople: IExternalPersonConfig[] = showPane ? loadDemoExternalPeople() : [];
let showExternalBadge = true;

function renderApp(): void {
  ReactDOM.render(
    <SmartOrgChart
      context={mockContext}
      defaultView={view}
      topLevelUser="a.chen@contoso.com"
      levelsBelow={levelsBelow}
      pageSize={50}
      useDemoData={true}
      hideDemoBanner={true}
      companyName="Contoso"
      logoUrl=""
      theme={theme}
      accentColor={accentColor}
      customAttributes={DEMO_CUSTOM_ATTRIBUTES}
      externalPeople={externalPeople}
      showExternalBadge={showExternalBadge}
      defaultLayout={layout}
      enableFindMe={true}
      enableLayoutToggle={true}
      enableStats={true}
      enableDeptFilter={true}
      enableUserFilter={true}
      defaultZoom={defaultZoom}
      onSettingsSaved={() => { /* no-op in demo */ }}
    />,
    document.getElementById('root'),
  );
}

if (showPane) {
  const pane = document.createElement('aside');
  pane.style.cssText = 'position:fixed;top:0;right:0;bottom:0;width:340px;overflow:auto;background:#fff;' +
    'border-left:1px solid #edebe9;padding:16px;box-sizing:border-box;font-family:"Segoe UI",sans-serif;z-index:10;';
  const title = document.createElement('div');
  title.textContent = 'Property pane (demo) — External People';
  title.style.cssText = 'font-size:16px;font-weight:600;margin-bottom:6px;';
  const note = document.createElement('div');
  note.textContent = 'Saved in this browser only. Try managers like a.chen@contoso.com, d.williams@contoso.com or e.rodriguez@contoso.com.';
  note.style.cssText = 'font-size:12px;color:#605e5c;margin-bottom:12px;';
  const field = document.createElement('div');
  // Stand-in for the pane's PropertyPaneToggle (SPFx controls aren't available in the demo)
  const badgeToggle = document.createElement('label');
  badgeToggle.style.cssText = 'display:flex;align-items:center;gap:6px;font-size:13px;margin-top:14px;cursor:pointer;';
  const badgeCheckbox = document.createElement('input');
  badgeCheckbox.type = 'checkbox';
  badgeCheckbox.checked = showExternalBadge;
  badgeCheckbox.addEventListener('change', () => { showExternalBadge = badgeCheckbox.checked; renderApp(); });
  badgeToggle.appendChild(badgeCheckbox);
  badgeToggle.appendChild(document.createTextNode('Show "External" badge'));
  pane.appendChild(title); pane.appendChild(note); pane.appendChild(field); pane.appendChild(badgeToggle);
  document.body.appendChild(pane);
  (document.getElementById('root') as HTMLElement).style.marginRight = '340px';

  // The real web part class (SPFx base class is stubbed in the demo build)
  const webPart = new SmartOrgChartWebPart() as any;
  webPart.properties = { externalPeople };
  webPart.context = { ...mockContext, instanceId: 'demo' };   // instanceId only makes the field's element ids unique
  webPart._renderExternalPeopleField(field, (prop: string, value: IExternalPersonConfig[]) => {
    if (prop !== 'externalPeople') return;
    webPart.properties.externalPeople = value;
    externalPeople = value;
    try { localStorage.setItem(LS_DEMO_EXTERNAL, JSON.stringify(value)); } catch { /* ignore */ }
    renderApp();
  });
}

renderApp();

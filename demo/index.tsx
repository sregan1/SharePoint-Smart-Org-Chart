import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { SmartOrgChart } from '../src/webparts/smartOrgChart/components/SmartOrgChart';
import { MockCompanySize } from '../src/services/MockGraphService';
import { OrgChartTheme } from '../src/webparts/smartOrgChart/components/ISmartOrgChartProps';
import { ICustomAttributeConfig } from '../src/services/GraphService';

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

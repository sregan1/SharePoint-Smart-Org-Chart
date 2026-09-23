import { GraphService, ICustomAttributeConfig, IOrgNode } from '../../../../services/GraphService';
import { OrgChartTheme } from '../ISmartOrgChartProps';

export interface IOrgChartProps {
  graphService: GraphService | null;
  topLevelUser: string;
  levelsBelow: number;
  levelsAbove: number;
  showDepartment: boolean;
  showOffice: boolean;
  theme: OrgChartTheme;
  accentColor?: string;
  /** Admin-configured Entra ID attributes; only entries with showInOrgChart apply here */
  customAttributes: ICustomAttributeConfig[];
  currentUserEmail: string;
  compactCards: boolean;
  defaultLayout: 'drill' | 'vertical' | 'horizontal';
  enableFindMe: boolean;
  enableLayoutToggle: boolean;
  enableStats: boolean;
  enableDeptFilter: boolean;
  enableUserFilter: boolean;
  defaultZoom: number;
  /** Web part instance ID — scopes localStorage state so instances on different pages don't clash */
  instanceId: string;
}

export interface IOrgChartState {
  rootNode: IOrgNode | null;
  isLoading: boolean;
  error: string | null;
  photos: { [userId: string]: string | null };
  expandingNodes: Set<string>;
  searchQuery: string;
}

import { GraphService, ICustomAttributeConfig } from '../../../../services/GraphService';
import { OrgChartTheme } from '../ISmartOrgChartProps';

export interface IEmployeeDirectoryProps {
  graphService: GraphService;
  alphabetFilterField: 'firstName' | 'lastName';
  cardSize: 'small' | 'medium' | 'large';
  showEmail: boolean;
  showPhone: boolean;
  showDepartment: boolean;
  showOffice: boolean;
  pageSize: number;
  theme: OrgChartTheme;
  accentColor?: string;
  /** Admin-configured Entra ID attributes; only entries with showInDirectory apply here */
  customAttributes: ICustomAttributeConfig[];
  /** Web part instance ID — scopes localStorage state so instances on different pages don't clash */
  instanceId: string;
}

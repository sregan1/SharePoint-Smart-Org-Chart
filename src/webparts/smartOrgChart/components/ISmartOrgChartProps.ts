import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ISmartOrgChartWebPartProps } from '../SmartOrgChartWebPart';

export type OrgChartTheme = 'modern' | 'minimal' | 'corporate' | 'dark' | 'custom';

export interface ISmartOrgChartProps extends ISmartOrgChartWebPartProps {
  context: WebPartContext;
}

export interface IUserSettings {
  alphabetFilterField: 'firstName' | 'lastName';
  cardSize: 'small' | 'medium' | 'large';
  showEmail: boolean;
  showPhone: boolean;
  showDepartment: boolean;
  showOffice: boolean;
  levelsAbove: number;
  compactCards: boolean;
  fontScale: number;
}

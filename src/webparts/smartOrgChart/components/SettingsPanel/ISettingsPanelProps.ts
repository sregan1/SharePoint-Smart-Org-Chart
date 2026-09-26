import { IUserSettings } from '../ISmartOrgChartProps';

export interface ISettingsPanelProps {
  isOpen: boolean;
  settings: IUserSettings;
  locale: string;
  onDismiss: () => void;
  onSave: (settings: IUserSettings) => void;
  mockSize?: number;
  onMockSizeChange?: (size: number) => void;
}

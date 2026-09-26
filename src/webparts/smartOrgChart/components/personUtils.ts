// Shared person-rendering helpers used by both the Employee Directory and the Org Chart.
import { PresenceAvailability } from '../../../services/GraphService';
import * as strings from 'SmartOrgChartWebPartStrings';

export const PRESENCE_COLOR: Record<PresenceAvailability, string> = {
  Available:    '#6BB700',
  Busy:         '#C50F1F',
  DoNotDisturb: '#C50F1F',
  BeRightBack:  '#FFAA44',
  Away:         '#FFAA44',
  Offline:      '#8A8886',
  Unknown:      '#8A8886',
};

export const PRESENCE_LABEL: Record<PresenceAvailability, string> = {
  Available: strings.Presence_Available, Busy: strings.Presence_Busy, DoNotDisturb: strings.Presence_DoNotDisturb,
  BeRightBack: strings.Presence_BeRightBack, Away: strings.Presence_Away, Offline: strings.Presence_Offline, Unknown: '',
};

export function getInitials(displayName: string): string {
  const parts = (displayName || '').split(' ').filter(p => p.length > 0);
  if (parts.length === 0) return '?';
  if (parts.length === 1) return parts[0][0].toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}


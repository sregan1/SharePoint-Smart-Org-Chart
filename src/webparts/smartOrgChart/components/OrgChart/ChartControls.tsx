import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { IFilterCounts } from './orgTreeUtils';
import * as strings from 'SmartOrgChartWebPartStrings';
import styles from './OrgChart.module.scss';

/* ── User-type filter dropdown ───────────── */

export type UserFilterKey = 'members' | 'guests' | 'disabled';

export interface IFilterPanelProps {
  filterMembers: boolean;
  filterGuests: boolean;
  filterDisabled: boolean;
  counts: IFilterCounts;
  onToggle: (key: UserFilterKey) => void;
}

export const FilterPanel: React.FC<IFilterPanelProps> = ({ filterMembers, filterGuests, filterDisabled, counts, onToggle }) => {
  const items: Array<{ key: UserFilterKey; label: string; count: number; checked: boolean }> = [
    { key: 'members',  label: strings.Chart_FilterMembers,  count: counts.members,  checked: filterMembers  },
    { key: 'guests',   label: strings.Chart_FilterGuests,   count: counts.guests,   checked: filterGuests   },
    { key: 'disabled', label: strings.Chart_FilterDisabled, count: counts.disabled, checked: filterDisabled },
  ];
  return (
    <div className={styles.filterPanel} role="group" aria-label={strings.Chart_ShowInChartAria}>
      <div className={styles.filterPanelTitle}>{strings.Chart_ShowInChartTitle}</div>
      {items.map(({ key, label, count, checked }) => (
        <label key={key} className={styles.filterItem}>
          <input type="checkbox" checked={checked} onChange={() => onToggle(key)} className={styles.filterCheckbox} />
          <span className={styles.filterLabel}>{label}</span>
          <span className={styles.filterCount}>{count}</span>
        </label>
      ))}
    </div>
  );
};

/* ── No-config form ──────────────────────── */

export const NoConfigForm: React.FC<{ onLoad: (id: string) => void }> = ({ onLoad }) => {
  const [val, setVal] = React.useState('');
  const submit = (): void => { if (val.trim()) onLoad(val.trim()); };
  return (
    <div className={styles.noConfig}>
      <Icon iconName="Org" className={styles.noConfigIcon} />
      <div className={styles.noConfigTitle}>{strings.Chart_SetupTitle}</div>
      <div className={styles.noConfigSubtitle}>{strings.Chart_SetupSubtitle}</div>
      <div className={styles.noConfigForm}>
        <TextField
          placeholder={strings.Chart_SetupPlaceholder}
          ariaLabel={strings.Chart_SetupInputAria}
          value={val}
          onChange={(_, v) => setVal(v || '')}
          onKeyDown={e => { if (e.key === 'Enter') submit(); }}
          className={styles.noConfigInput}
        />
        <PrimaryButton text={strings.Chart_SetupLoadButton} onClick={submit} disabled={!val.trim()} />
      </div>
      <div className={styles.noConfigHint}>{strings.Chart_SetupHint}</div>
    </div>
  );
};

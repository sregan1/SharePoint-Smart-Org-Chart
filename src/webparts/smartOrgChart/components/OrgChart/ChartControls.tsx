import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { IFilterCounts } from './orgTreeUtils';
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
    { key: 'members',  label: 'Regular members',   count: counts.members,  checked: filterMembers  },
    { key: 'guests',   label: 'Guest users',       count: counts.guests,   checked: filterGuests   },
    { key: 'disabled', label: 'Disabled accounts', count: counts.disabled, checked: filterDisabled },
  ];
  return (
    <div className={styles.filterPanel} role="group" aria-label="Show in chart">
      <div className={styles.filterPanelTitle}>Show in chart</div>
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
      <div className={styles.noConfigTitle}>Set Up the Org Chart</div>
      <div className={styles.noConfigSubtitle}>Enter the top-level person&apos;s email or UPN to get started.</div>
      <div className={styles.noConfigForm}>
        <TextField
          placeholder="ceo@company.com"
          ariaLabel="Top-level person's email or UPN"
          value={val}
          onChange={(_, v) => setVal(v || '')}
          onKeyDown={e => { if (e.key === 'Enter') submit(); }}
          className={styles.noConfigInput}
        />
        <PrimaryButton text="Load" onClick={submit} disabled={!val.trim()} />
      </div>
      <div className={styles.noConfigHint}>You can also set this permanently in the web part settings (admin).</div>
    </div>
  );
};

import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Slider } from '@fluentui/react/lib/Slider';
import { Separator } from '@fluentui/react/lib/Separator';
import { Label } from '@fluentui/react/lib/Label';
import { IUserSettings } from '../ISmartOrgChartProps';
import { ISettingsPanelProps } from './ISettingsPanelProps';
import { formatString } from '../localeUtils';
import * as strings from 'SmartOrgChartWebPartStrings';
import styles from './SettingsPanel.module.scss';

interface ISettingsPanelState {
  draft: IUserSettings;
}

const alphabetOptions: IDropdownOption[] = [
  { key: 'firstName', text: strings.Settings_AlphabetFirstName },
  { key: 'lastName', text: strings.Settings_AlphabetLastName }
];

const cardSizeOptions: IDropdownOption[] = [
  { key: 'small', text: strings.Settings_CardSizeSmall },
  { key: 'medium', text: strings.Settings_CardSizeMedium },
  { key: 'large', text: strings.Settings_CardSizeLarge }
];

const fontScaleOptions: IDropdownOption[] = [
  { key: 0.75, text: strings.PropertyPane_FontSize_75 },
  { key: 0.85, text: strings.PropertyPane_FontSize_85 },
  { key: 1,    text: strings.PropertyPane_FontSize_100 },
  { key: 1.15, text: strings.PropertyPane_FontSize_115 },
  { key: 1.3,  text: strings.PropertyPane_FontSize_130 },
  { key: 1.5,  text: strings.PropertyPane_FontSize_150 },
  { key: 1.75, text: strings.PropertyPane_FontSize_175 },
];

export class SettingsPanel extends React.Component<ISettingsPanelProps, ISettingsPanelState> {
  constructor(props: ISettingsPanelProps) {
    super(props);
    this.state = { draft: { ...props.settings } };
  }

  public componentDidUpdate(prev: ISettingsPanelProps): void {
    if (!prev.isOpen && this.props.isOpen) {
      this.setState({ draft: { ...this.props.settings } });
    }
  }

  private _update = <K extends keyof IUserSettings>(key: K, value: IUserSettings[K]): void => {
    this.setState(prev => ({ draft: { ...prev.draft, [key]: value } }));
  }

  private _save = (): void => {
    this.props.onSave(this.state.draft);
  }

  private _reset = (): void => {
    this.setState({ draft: { ...this.props.settings } });
  }

  public render(): React.ReactElement {
    const { isOpen, onDismiss, mockSize, onMockSizeChange, locale } = this.props;
    const { draft } = this.state;

    return (
      <Panel
        isOpen={isOpen}
        onDismiss={onDismiss}
        type={PanelType.medium}
        headerText={strings.Settings_HeaderText}
        isFooterAtBottom
        onRenderFooterContent={() => (
          <div className={styles.footer}>
            <PrimaryButton text={strings.Settings_SaveButton} onClick={this._save} />
            <DefaultButton text={strings.Settings_CancelButton} onClick={onDismiss} style={{ marginLeft: 8 }} />
            <DefaultButton text={strings.Settings_DiscardButton} onClick={this._reset} style={{ marginLeft: 'auto' }} />
          </div>
        )}
      >
        <div className={styles.body}>

          {/* ── Directory display ── */}
          <Separator alignContent="start">
            <span className={styles.sectionLabel}>{strings.Settings_SectionDirectory}</span>
          </Separator>

          <Dropdown
            label={strings.Settings_AlphabetFilterByLabel}
            selectedKey={draft.alphabetFilterField}
            options={alphabetOptions}
            onChange={(_, o) => o && this._update('alphabetFilterField', o.key as 'firstName' | 'lastName')}
          />

          <Dropdown
            label={strings.Settings_CardSizeLabel}
            selectedKey={draft.cardSize}
            options={cardSizeOptions}
            onChange={(_, o) => o && this._update('cardSize', o.key as 'small' | 'medium' | 'large')}
            className={styles.field}
          />

          <Dropdown
            label={strings.Settings_FontSizeLabel}
            selectedKey={draft.fontScale || 1}
            options={fontScaleOptions}
            onChange={(_, o) => o && this._update('fontScale', o.key as number)}
            className={styles.field}
          />

          {/* ── Cards & Fields ── */}
          <Separator alignContent="start" className={styles.separator}>
            <span className={styles.sectionLabel}>{strings.Settings_SectionCardsAndFields}</span>
          </Separator>

          <div className={styles.toggleGroup}>
            <Label>{strings.Settings_ShowOnCardsLabel}</Label>
            <Toggle
              label={strings.Settings_EmailAddressLabel}
              checked={draft.showEmail}
              onChange={(_, v) => this._update('showEmail', !!v)}
              inlineLabel
            />
            <Toggle
              label={strings.Settings_PhoneNumberLabel}
              checked={draft.showPhone}
              onChange={(_, v) => this._update('showPhone', !!v)}
              inlineLabel
            />
            <Toggle
              label={strings.Settings_DepartmentLabel}
              checked={draft.showDepartment}
              onChange={(_, v) => this._update('showDepartment', !!v)}
              inlineLabel
            />
            <Toggle
              label={strings.Settings_OfficeLocationLabel}
              checked={draft.showOffice}
              onChange={(_, v) => this._update('showOffice', !!v)}
              inlineLabel
            />
          </div>

          {/* ── Org Chart ── */}
          <Separator alignContent="start" className={styles.separator}>
            <span className={styles.sectionLabel}>{strings.Settings_SectionOrgChart}</span>
          </Separator>

          <Slider
            label={formatString(strings.Settings_ManagerLevelsLabel, { levels: draft.levelsAbove })}
            min={0}
            max={5}
            step={1}
            value={draft.levelsAbove}
            onChange={v => this._update('levelsAbove', v)}
            className={styles.field}
            showValue={false}
          />

          <Toggle
            label={strings.Settings_CompactCardsLabel}
            checked={draft.compactCards}
            onChange={(_, v) => this._update('compactCards', !!v)}
            inlineLabel
            className={styles.field}
            onText={strings.Settings_CompactCardsOn}
            offText={strings.Settings_CompactCardsOff}
          />

          <div className={styles.hint}>
            {strings.Settings_Hint}
          </div>

          {/* ── Demo data ── */}
          {mockSize !== undefined && onMockSizeChange && (
            <>
              <Separator alignContent="start" className={styles.separator}>
                <span className={styles.sectionLabel}>{strings.Settings_SectionDemoData}</span>
              </Separator>
              <Label>{strings.Settings_DatasetSizeLabel}</Label>
              <div className={styles.mockSizeBtns}>
                {([150, 500, 1000] as const).map(s => (
                  <button
                    key={s}
                    className={`${styles.mockSizeBtn}${mockSize === s ? ` ${styles.mockSizeBtnActive}` : ''}`}
                    onClick={() => onMockSizeChange(s)}
                  >
                    {formatString(strings.Settings_PeopleCountSuffix, { count: s.toLocaleString(locale) })}
                  </button>
                ))}
              </div>
            </>
          )}
        </div>
      </Panel>
    );
  }
}

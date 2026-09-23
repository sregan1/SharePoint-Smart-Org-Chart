import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { FocusTrapZone } from '@fluentui/react/lib/FocusTrapZone';
import { ICustomAttributeConfig, IGraphUser, PresenceAvailability } from '../../../../services/GraphService';
import { PRESENCE_COLOR, PRESENCE_LABEL, getInitials } from '../personUtils';
import { OrgChartTheme, getThemeTokens } from './orgTheme';
import { dedupeUsers } from './orgTreeUtils';
import styles from './OrgChart.module.scss';

/* ── Person Card (Outlook-style popup) ───── */

export interface IPersonCardProps {
  user: IGraphUser;
  photo: string | null;
  presence: PresenceAvailability | undefined;
  theme: OrgChartTheme;
  accentColor?: string;
  managerChain: IGraphUser[];
  dottedManager: IGraphUser | null;
  dottedReports: IGraphUser[];
  directReportCount: number;
  totalReportCount: number;
  /** Admin-configured Entra ID attributes; only entries with showInOrgChart apply here */
  customAttributes: ICustomAttributeConfig[];
  onClose: () => void;
  onFocus: (user: IGraphUser) => void;
}

export const PersonCard: React.FC<IPersonCardProps> = ({
  user, photo, presence, theme, accentColor, managerChain, dottedManager, dottedReports,
  directReportCount, totalReportCount, customAttributes, onClose, onFocus
}) => {
  React.useEffect(() => {
    const onKey = (e: KeyboardEvent): void => { if (e.key === 'Escape') onClose(); };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [onClose]);

  // Transient "Copied!" feedback for the copy buttons
  const mountedRef = React.useRef(true);
  React.useEffect(() => () => { mountedRef.current = false; }, []);
  const [copied, setCopied] = React.useState('');
  const copy = (text: string, label: string): void => {
    const done = (): void => {
      if (!mountedRef.current) return;
      setCopied(label);
      window.setTimeout(() => { if (mountedRef.current) setCopied(''); }, 2000);
    };
    try { navigator.clipboard.writeText(text).then(done).catch(done); } catch { done(); }
  };

  const t = getThemeTokens(theme, accentColor);
  const initials   = getInitials(user.displayName);
  const isDisabled = user.accountEnabled === false;
  const isGuest    = user.userType === 'Guest';
  // Cycles in the directory can repeat a manager (or the person themself)
  const chain      = dedupeUsers(managerChain, user.id);
  const reports    = dedupeUsers(dottedReports, user.id);
  const onAccentStyle = t.isCustom ? { color: t.onAccent } : undefined;

  return (
    <div className={styles.personCardOverlay} onClick={onClose}>
      {/* Focus stays inside the dialog while open; the chart restores focus to the opener on close */}
      <FocusTrapZone
        className={styles.personCard}
        style={{ background: t.cardBg, borderColor: t.border }}
        onClick={e => e.stopPropagation()}
        role="dialog"
        aria-modal="true"
        aria-label={`Profile: ${user.displayName}`}
        firstFocusableTarget={`.${styles.personCardClose}`}
        isClickableOutsideFocusTrap={true}
        disableRestoreFocus={true}
      >
        {/* Colored header band */}
        <div className={styles.personCardHeader} style={{ background: t.accent }}>
          <button className={styles.personCardClose} onClick={onClose} title="Close" aria-label="Close profile" style={onAccentStyle}>
            <Icon iconName="Cancel" />
          </button>
          {photo
            ? <img src={photo} alt={user.displayName} className={styles.personCardPhoto} />
            : <div className={styles.personCardInitials} style={onAccentStyle}>{initials}</div>
          }
          {presence && presence !== 'Unknown' && (
            <div className={styles.personCardPresence} style={onAccentStyle}>
              <span className={styles.personCardPresenceDot} style={{ background: PRESENCE_COLOR[presence] }} />
              <span>{PRESENCE_LABEL[presence]}</span>
            </div>
          )}
        </div>

        {/* Body */}
        <div className={styles.personCardBody}>
          <div className={styles.personCardName} style={{ color: t.text }}>{user.displayName}</div>
          {user.jobTitle && (
            <div className={styles.personCardTitle} style={{ color: t.accentText }}>{user.jobTitle}</div>
          )}

          {/* Badges */}
          <div className={styles.personCardBadges}>
            {user.department && (
              <span className={styles.personCardDeptBadge} style={{ background: t.accentSoft, color: t.accentText }}>
                {user.department}
              </span>
            )}
            {isDisabled && <span className={styles.personCardStatusBadge} style={{ background: '#fde7e9', color: '#c50f1f' }}>Disabled</span>}
            {isGuest   && <span className={styles.personCardStatusBadge} style={{ background: '#fff4ce', color: '#835c00' }}>Guest</span>}
          </div>

          {/* Direct + total headcount */}
          {directReportCount > 0 && (
            <div className={styles.personCardReportsLine} style={{ color: t.subText }}>
              <Icon iconName="Group" />
              {directReportCount} direct report{directReportCount === 1 ? '' : 's'}
              {totalReportCount > directReportCount && ` · ${totalReportCount} total`}
            </div>
          )}

          {/* Info fields */}
          <div className={styles.personCardFields} style={{ background: t.fieldBg, borderColor: t.border }}>
            {user.mail && (
              <div className={styles.personCardField}>
                <Icon iconName="Mail" className={styles.personCardFieldIcon} style={{ color: t.accentText }} />
                <a href={`mailto:${user.mail}`} className={styles.personCardFieldLink} style={{ color: t.accentText }}>{user.mail}</a>
                <button
                  onClick={() => copy(user.mail, 'email')}
                  title="Copy email address"
                  aria-label={copied === 'email' ? 'Email address copied' : 'Copy email address'}
                  style={{ background: 'none', border: 'none', cursor: 'pointer', color: t.subText, padding: '2px 4px' }}
                >
                  <Icon iconName={copied === 'email' ? 'CheckMark' : 'Copy'} />
                </button>
              </div>
            )}
            {user.businessPhones && user.businessPhones[0] && (
              <div className={styles.personCardField}>
                <Icon iconName="Phone" className={styles.personCardFieldIcon} style={{ color: t.subText }} />
                <a href={`tel:${user.businessPhones[0]}`} className={styles.personCardFieldText} style={{ color: t.subText }}>{user.businessPhones[0]}</a>
              </div>
            )}
            {user.mobilePhone && (
              <div className={styles.personCardField}>
                <Icon iconName="CellPhone" className={styles.personCardFieldIcon} style={{ color: t.subText }} />
                <a href={`tel:${user.mobilePhone}`} className={styles.personCardFieldText} style={{ color: t.subText }}>{user.mobilePhone}</a>
              </div>
            )}
            {user.officeLocation && (
              <div className={styles.personCardField}>
                <Icon iconName="POI" className={styles.personCardFieldIcon} style={{ color: t.subText }} />
                <span className={styles.personCardFieldText} style={{ color: t.subText }}>{user.officeLocation}</span>
              </div>
            )}
            {customAttributes.filter(attr => attr.showInOrgChart).map(attr => {
              const value = user.customAttributes?.[attr.graphField];
              if (!value) return null;
              return (
                <div className={styles.personCardField} key={attr.id} title={attr.label}>
                  <Icon iconName="Tag" className={styles.personCardFieldIcon} style={{ color: t.subText }} />
                  <span className={styles.personCardFieldText} style={{ color: t.subText }}>{attr.label}: {value}</span>
                </div>
              );
            })}
          </div>

          {/* Reporting chain */}
          {chain.length > 0 && (
            <div className={styles.personCardChain} style={{ background: t.chainBg, borderColor: t.border }}>
              <div className={styles.personCardChainLabel} style={{ color: t.subText }}>Reports to</div>
              <div className={styles.personCardChainItems}>
                {chain.map((mgr, i) => (
                  <React.Fragment key={mgr.id}>
                    {i > 0 && <Icon iconName="ChevronRight" className={styles.personCardChainSep} style={{ color: t.subText }} />}
                    <button
                      className={styles.personCardChainChip}
                      onClick={() => { onClose(); onFocus(mgr); }}
                      title={`Focus on ${mgr.displayName}`}
                      aria-label={`Focus org chart on ${mgr.displayName}`}
                    >
                      <span className={styles.personCardChainInitials} style={{ background: t.accent, color: t.onAccent }}>
                        {getInitials(mgr.displayName)}
                      </span>
                      <span className={styles.personCardChainName} style={{ color: t.text }}>
                        {mgr.displayName.split(' ')[0]}
                      </span>
                    </button>
                  </React.Fragment>
                ))}
              </div>
            </div>
          )}

          {/* Dotted-line relationships */}
          {(dottedManager || reports.length > 0) && (
            <div className={styles.personCardChain} style={{ background: t.chainBg, borderColor: t.border }}>
              <div className={styles.personCardChainLabel} style={{ color: t.subText }}>Dotted line</div>
              <div className={styles.personCardChainItems}>
                {dottedManager && (
                  <button
                    className={styles.personCardChainChip}
                    onClick={() => { onClose(); onFocus(dottedManager); }}
                    title={`Dotted-line manager: ${dottedManager.displayName}`}
                  >
                    <span className={styles.personCardChainInitials} style={{ background: t.accent, color: t.onAccent }}>
                      {getInitials(dottedManager.displayName)}
                    </span>
                    <span className={styles.personCardChainName} style={{ color: t.text }}>
                      ↑ {dottedManager.displayName.split(' ')[0]}
                    </span>
                  </button>
                )}
                {reports.map(rep => (
                  <button
                    key={rep.id}
                    className={styles.personCardChainChip}
                    onClick={() => { onClose(); onFocus(rep); }}
                    title={`Dotted-line report: ${rep.displayName}`}
                  >
                    <span className={styles.personCardChainInitials} style={{ background: t.accent, color: t.onAccent }}>
                      {getInitials(rep.displayName)}
                    </span>
                    <span className={styles.personCardChainName} style={{ color: t.text }}>
                      ↓ {rep.displayName.split(' ')[0]}
                    </span>
                  </button>
                ))}
              </div>
            </div>
          )}

          {/* Action buttons */}
          {user.mail && (
            <div className={styles.personCardActions}>
              <a
                href={`https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(user.mail)}`}
                target="_blank" rel="noopener noreferrer"
                className={styles.personCardAction}
                style={{ background: t.accent, color: t.onAccent }}
              >
                <Icon iconName="Chat" />&nbsp;Chat
              </a>
              <a
                href={`mailto:${user.mail}`}
                className={styles.personCardAction}
                style={{ background: t.neutralBtnBg, color: t.neutralBtnText }}
              >
                <Icon iconName="Mail" />&nbsp;Email
              </a>
              <button
                className={styles.personCardAction}
                style={{ background: t.neutralBtnBg, color: t.neutralBtnText, border: 'none', cursor: 'pointer' }}
                onClick={() => { onClose(); onFocus(user); }}
                title="Focus org chart on this person"
              >
                <Icon iconName="Org" />&nbsp;Focus
              </button>
            </div>
          )}
        </div>
      </FocusTrapZone>
    </div>
  );
};

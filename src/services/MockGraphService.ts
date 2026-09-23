// MockGraphService — used automatically when the web part runs on localhost or demo mode.
// Supports three company sizes (150 / 500 / 1 000) selectable via the demo banner.
import { GraphService, IGraphUser, IRawUserData, IUserFilterOptions, PresenceAvailability } from './GraphService';

const delay = (ms: number): Promise<void> => new Promise(r => setTimeout(r, ms));

function hashCode(s: string): number {
  let h = 0;
  for (let i = 0; i < s.length; i++) h = (Math.imul(31, h) + s.charCodeAt(i)) | 0;
  return Math.abs(h);
}

// Seeded LCG — deterministic, same data every reload
function lcg(seed: number): () => number {
  let s = seed >>> 0;
  return (): number => { s = (Math.imul(1664525, s) + 1013904223) >>> 0; return s / 0x100000000; };
}

const PRESENCE_POOL: PresenceAvailability[] = [
  'Available', 'Available', 'Available', 'Available', // 40 %
  'Busy', 'Busy',                                     // 20 %
  'Away', 'BeRightBack',                              // 10 % each
  'DoNotDisturb',                                     // 10 %
  'Offline',                                          // 10 %
];

// [email, displayName, jobTitle, department, managerEmail|null, office, businessPhone, accountEnabled?, userType?]
type RawUser = [string, string, string, string, string | null, string, string, boolean?, string?];

// Demo values for the attributes admins commonly configure as "custom
// attributes" (see ICustomAttributeConfig) — always generated so a newly
// configured attribute shows real-looking data immediately in demo mode,
// without the mock needing to know which fields were actually configured.
const COST_CENTER: Record<string, string> = {
  'Executive':        'CC-1900',
  'Engineering':      'CC-1000',
  'Product':          'CC-1100',
  'Sales':            'CC-1200',
  'Marketing':        'CC-1300',
  'Finance':          'CC-1400',
  'Human Resources':  'CC-1500',
  'Operations':       'CC-1600',
  'Legal':            'CC-1700',
};

function buildDemoCustomAttributes(
  email: string, department: string, officeLocation: string, isGuest: boolean
): { [field: string]: string } {
  const n = hashCode(email);
  const city = (officeLocation || '').split(',')[0].trim() || 'Remote';
  return {
    employeeId:          'E' + (10000 + (n % 90000)),
    employeeType:        isGuest ? 'Contractor' : 'Employee',
    companyName:         'Contoso Ltd',
    city,
    extensionAttribute1: COST_CENTER[department] || 'CC-0000',
    extensionAttribute2: `Building ${String.fromCharCode(65 + (n % 4))}-${1 + (n % 3)}`,
  };
}

// ── Name pools for generated users ────────────────────────────────────────────
const FIRST_NAMES = [
  'Alex','Blake','Cameron','Dana','Elliott','Finley','Grace','Harper','Iris','Jordan',
  'Kennedy','Lane','Morgan','Nova','Oakley','Parker','Quinn','Reese','Sage','Taylor',
  'Uma','Val','Wesley','Xara','Yael','Zuri','Aiden','Brianna','Carlos','Divya',
  'Erica','Finn','Gianna','Hiroshi','Isla','Javier','Kaito','Lena','Matteo','Nina',
  'Omar','Priya','Raj','Sara','Tomas','Vera','Will','Xia','Yuki','Zara',
  'Adeola','Beatriz','Coen','Fatou','Gideon','Hana','Ivan','Juno','Kenji','Laila',
];

const LAST_NAMES = [
  'Adams','Baker','Chen','Davis','Evans','Foster','Garcia','Harris','Ibrahim','Jones',
  'Khan','Lee','Martinez','Nguyen','Osei','Patel','Quinn','Reed','Santos','Torres',
  'Ueda','Vance','Wang','Xavier','Yip','Zhang','Anderson','Brown','Clark','Diaz',
  'Ellis','Ford','Green','Hill','Iyer','Jackson','Kim','Lopez','Moore','Nash',
  'Owen','Price','Rivera','Singh','Thompson','Upton','Villa','White','Xiao','Young',
  'Zhou','Adeyemi','Blum','Chakraborty','Delgado','Ferreira','Goldstein','Hendrix','Ingram','Johansson',
];

// ── Generator infrastructure ──────────────────────────────────────────────────
interface MgrSlot { email: string; dept: string; office: string; area: string; }

// Weighted: more slots for Engineering & Sales to give them proportionally larger teams
const BASE_SLOTS: MgrSlot[] = [
  // Engineering — 40 %
  { email: 'e.rodriguez@contoso.com', dept: 'Engineering',     office: 'Seattle, WA',  area: '206' },
  { email: 'e.rodriguez@contoso.com', dept: 'Engineering',     office: 'Seattle, WA',  area: '206' },
  { email: 'k.park@contoso.com',      dept: 'Engineering',     office: 'Seattle, WA',  area: '206' },
  { email: 'k.park@contoso.com',      dept: 'Engineering',     office: 'Seattle, WA',  area: '206' },
  { email: 'an.gupta@contoso.com',    dept: 'Engineering',     office: 'Seattle, WA',  area: '206' },
  { email: 'p.sharma@contoso.com',    dept: 'Engineering',     office: 'Seattle, WA',  area: '206' },
  { email: 'o.brown@contoso.com',     dept: 'Engineering',     office: 'Seattle, WA',  area: '206' },
  // Sales — 20 %
  { email: 'r.chen@contoso.com',      dept: 'Sales',           office: 'New York, NY', area: '212' },
  { email: 'ev.price@contoso.com',    dept: 'Sales',           office: 'New York, NY', area: '212' },
  { email: 'a.thompson@contoso.com',  dept: 'Sales',           office: 'New York, NY', area: '212' },
  { email: 'c.garcia@contoso.com',    dept: 'Sales',           office: 'New York, NY', area: '212' },
  { email: 'ka.ibrahim@contoso.com',  dept: 'Sales',           office: 'New York, NY', area: '212' },
  // Marketing — 10 %
  { email: 'am.clark@contoso.com',    dept: 'Marketing',       office: 'Chicago, IL',  area: '312' },
  { email: 'b.wilson@contoso.com',    dept: 'Marketing',       office: 'Chicago, IL',  area: '312' },
  { email: 'n.davis@contoso.com',     dept: 'Marketing',       office: 'Chicago, IL',  area: '312' },
  // Product — 8 %
  { email: 'mi.davis@contoso.com',    dept: 'Product',         office: 'Seattle, WA',  area: '206' },
  { email: 'je.lee@contoso.com',      dept: 'Product',         office: 'Seattle, WA',  area: '206' },
  // Finance — 8 %
  { email: 'mi.wright@contoso.com',   dept: 'Finance',         office: 'Chicago, IL',  area: '312' },
  { email: 'ch.anderson@contoso.com', dept: 'Finance',         office: 'Chicago, IL',  area: '312' },
  // HR — 7 %
  { email: 'je.thomas@contoso.com',   dept: 'Human Resources', office: 'Chicago, IL',  area: '312' },
  { email: 'ry.jackson@contoso.com',  dept: 'Human Resources', office: 'Chicago, IL',  area: '312' },
  // Operations — 7 %
  { email: 'sa.moore@contoso.com',    dept: 'Operations',      office: 'Austin, TX',   area: '512' },
  { email: 'da.harris@contoso.com',   dept: 'Operations',      office: 'Austin, TX',   area: '512' },
];

const LEAD_TITLE: Record<string, string> = {
  'Engineering':     'Engineering Manager',
  'Sales':           'Sales Manager',
  'Marketing':       'Marketing Manager',
  'Product':         'Senior Product Manager',
  'Finance':         'Finance Manager',
  'Human Resources': 'HR Manager',
  'Operations':      'Operations Manager',
};

const IC_TITLES: Record<string, string[]> = {
  'Engineering':     ['Software Engineer I', 'Software Engineer II', 'Software Engineer III', 'QA Engineer', 'DevOps Engineer', 'Site Reliability Engineer', 'Security Engineer'],
  'Sales':           ['Account Executive', 'Sales Development Representative', 'Sales Engineer', 'Account Manager', 'Solution Consultant'],
  'Marketing':       ['Marketing Specialist', 'Content Writer', 'Digital Marketing Specialist', 'Brand Designer', 'Marketing Coordinator', 'Social Media Manager'],
  'Product':         ['Associate Product Manager', 'Product Manager', 'Product Analyst', 'UX Researcher', 'UX Designer'],
  'Finance':         ['Financial Analyst', 'Senior Financial Analyst', 'Accountant', 'Business Analyst', 'Treasury Analyst'],
  'Human Resources': ['HR Business Partner', 'Recruiter', 'HR Coordinator', 'L&D Specialist', 'HR Analyst', 'Benefits Specialist'],
  'Operations':      ['IT Engineer', 'Customer Success Manager', 'Customer Success Associate', 'IT Support Specialist', 'Operations Analyst'],
};

// ── Core hand-authored dataset (~170 users, 5–8 levels deep throughout) ────────
// Depth legend (hops from CEO, 0-indexed):
//   Engineering Backend:  8 hops (deepest — Morgan Adeyemi)
//   Engineering Frontend: 7 hops (Kim Lane)
//   Sales Enterprise:     7 hops (Dana Bell / Kim Moss / Jo Wade / Cy Ford)
//   Most other branches:  5–6 hops
const RAW_BASE: RawUser[] = [
  // ── L0: CEO ───────────────────────────────────────────────────────────────────
  ['a.chen@contoso.com',        'Alexandra Chen',     'Chief Executive Officer',                  'Executive',       null,                      'Chicago, IL',      '+1 312 555 0100'],

  // ── L1: VPs (→ CEO) ───────────────────────────────────────────────────────────
  ['m.johnson@contoso.com',     'Marcus Johnson',     'VP Engineering',                           'Engineering',     'a.chen@contoso.com',      'Seattle, WA',      '+1 206 555 0101'],
  ['s.patel@contoso.com',       'Sofia Patel',        'VP Product',                               'Product',         'a.chen@contoso.com',      'Seattle, WA',      '+1 206 555 0102'],
  ['d.williams@contoso.com',    'David Williams',     'VP Sales',                                 'Sales',           'a.chen@contoso.com',      'New York, NY',     '+1 212 555 0103'],
  ['r.kim@contoso.com',         'Rachel Kim',         'VP Marketing',                             'Marketing',       'a.chen@contoso.com',      'Chicago, IL',      '+1 312 555 0104'],
  ['t.brown@contoso.com',       'Thomas Brown',       'VP Finance',                               'Finance',         'a.chen@contoso.com',      'Chicago, IL',      '+1 312 555 0105'],
  ['l.martinez@contoso.com',    'Lisa Martinez',      'VP Human Resources',                       'Human Resources', 'a.chen@contoso.com',      'Chicago, IL',      '+1 312 555 0106'],
  ['j.taylor@contoso.com',      'James Taylor',       'VP Operations',                            'Operations',      'a.chen@contoso.com',      'Austin, TX',       '+1 512 555 0107'],
  ['vl.santos@contoso.com',     'Victoria Santos',    'VP Legal & Compliance',                    'Legal',           'a.chen@contoso.com',      'Chicago, IL',      '+1 312 555 0900'],

  // ══════════════════════════════════════════════════════════════════════════════
  // ENGINEERING  (Marcus Johnson, VP)
  // ══════════════════════════════════════════════════════════════════════════════

  // L2: Directors → Marcus Johnson
  ['e.rodriguez@contoso.com',   'Emily Rodriguez',    'Director, Frontend Engineering',           'Engineering',     'm.johnson@contoso.com',   'Seattle, WA',      '+1 206 555 0110'],
  ['k.park@contoso.com',        'Kevin Park',         'Director, Backend Engineering',            'Engineering',     'm.johnson@contoso.com',   'Seattle, WA',      '+1 206 555 0111'],
  ['p.sharma@contoso.com',      'Priya Sharma',       'Director, Platform Engineering',           'Engineering',     'm.johnson@contoso.com',   'Seattle, WA',      '+1 206 555 0112'],
  ['o.brown@contoso.com',       'Oliver Brown',       'Director, Quality Engineering',            'Engineering',     'm.johnson@contoso.com',   'Seattle, WA',      '+1 206 555 0113'],

  // ── Frontend branch (max depth 7: CEO→VP→Dir→SrMgr→TL→SrEng→Eng→JrEng) ────
  // L3: Senior Manager → Emily Rodriguez
  ['zh.wang@contoso.com',       'Zhao Wang',          'Senior Manager, Frontend Engineering',     'Engineering',     'e.rodriguez@contoso.com', 'Seattle, WA',      '+1 206 555 0200'],
  // L4: Tech Lead → Zhao Wang
  ['in.patel@contoso.com',      'Indira Patel',       'Tech Lead, Frontend',                      'Engineering',     'zh.wang@contoso.com',     'Seattle, WA',      '+1 206 555 0201'],
  // L5: Senior Engineers → Indira Patel
  ['no.kim@contoso.com',        'Noah Kim',           'Senior Frontend Engineer',                 'Engineering',     'in.patel@contoso.com',    'Seattle, WA',      '+1 206 555 0202'],
  // L6: Engineer → Noah Kim
  ['za.ali@contoso.com',        'Zara Ali',           'Frontend Engineer',                        'Engineering',     'no.kim@contoso.com',      'Seattle, WA',      '+1 206 555 0203'],
  // L7: Junior → Zara Ali
  ['ki.lane@contoso.com',       'Kim Lane',           'Junior Frontend Engineer',                 'Engineering',     'za.ali@contoso.com',      'Seattle, WA',      '+1 206 555 0204'],

  // ── Backend branch (max depth 8: …→SrMgr→EngMgr→TL→SrBE→BE→JrBE) ──────────
  // L3: Senior Manager → Kevin Park
  ['an.gupta@contoso.com',      'Ananya Gupta',       'Senior Manager, Backend Engineering',      'Engineering',     'k.park@contoso.com',      'Seattle, WA',      '+1 206 555 0210'],
  // L4: Engineering Manager + direct ICs → Ananya Gupta
  ['be.chang@contoso.com',      'Benjamin Chang',     'Engineering Manager, Backend Platform',    'Engineering',     'an.gupta@contoso.com',    'Seattle, WA',      '+1 206 555 0211'],
  ['ty.scott@contoso.com',      'Tyler Scott',        'Backend Engineer',                         'Engineering',     'an.gupta@contoso.com',    'Seattle, WA',      '+1 206 555 0212'],
  ['me.okonkwo@contoso.com',    'Mele Okonkwo',       'Backend Engineer',                         'Engineering',     'an.gupta@contoso.com',    'Seattle, WA',      '+1 206 555 0213'],
  ['hu.zhang@contoso.com',      'Hui Zhang',          'Backend Engineer',                         'Engineering',     'an.gupta@contoso.com',    'Seattle, WA',      '+1 206 555 0214'],
  // L5: Tech Lead + direct ICs → Benjamin Chang
  ['sa.yip@contoso.com',        'Sam Yip',            'Tech Lead, Backend',                       'Engineering',     'be.chang@contoso.com',    'Seattle, WA',      '+1 206 555 0215'],
  ['al.stone@contoso.com',      'Alexander Stone',    'Junior Backend Engineer',                  'Engineering',     'be.chang@contoso.com',    'Seattle, WA',      '+1 206 555 0216'],
  ['ma.osei@contoso.com',       'Maya Osei',          'Junior Backend Engineer',                  'Engineering',     'be.chang@contoso.com',    'Seattle, WA',      '+1 206 555 0217'],
  // L6: Senior Engineers → Sam Yip
  ['fi.nakamura@contoso.com',   'Finn Nakamura',      'Senior Backend Engineer',                  'Engineering',     'sa.yip@contoso.com',      'Seattle, WA',      '+1 206 555 0218'],
  ['le.vega@contoso.com',       'Leila Vega',         'Senior Backend Engineer',                  'Engineering',     'sa.yip@contoso.com',      'Seattle, WA',      '+1 206 555 0219'],
  // L7: Engineers → Finn / Leila
  ['qi.zhou@contoso.com',       'Qi Zhou',            'Backend Engineer',                         'Engineering',     'fi.nakamura@contoso.com', 'Seattle, WA',      '+1 206 555 0220'],
  ['jo.ashford@contoso.com',    'Jordan Ashford',     'Backend Engineer',                         'Engineering',     'le.vega@contoso.com',     'Seattle, WA',      '+1 206 555 0221'],
  // L8: Junior → Qi Zhou (deepest node)
  ['mo.adeyemi@contoso.com',    'Morgan Adeyemi',     'Junior Backend Engineer',                  'Engineering',     'qi.zhou@contoso.com',     'Seattle, WA',      '+1 206 555 0222'],

  // ── Platform & DevOps branch (max depth 6) ───────────────────────────────────
  // L3: Senior Manager → Priya Sharma
  ['de.nguyen@contoso.com',     'Derek Nguyen',       'Senior Manager, Platform & Reliability',   'Engineering',     'p.sharma@contoso.com',    'Seattle, WA',      '+1 206 555 0230'],
  // L4: Manager + direct ICs → Derek Nguyen
  ['ro.diaz@contoso.com',       'Rosa Diaz',          'DevOps & SRE Manager',                     'Engineering',     'de.nguyen@contoso.com',   'Seattle, WA',      '+1 206 555 0231'],
  ['lu.fernandez@contoso.com',  'Luna Fernandez',     'Platform Engineer',                        'Engineering',     'de.nguyen@contoso.com',   'Seattle, WA',      '+1 206 555 0232'],
  ['ka.reed@contoso.com',       'Kai Reed',           'Platform Engineer',                        'Engineering',     'de.nguyen@contoso.com',   'Seattle, WA',      '+1 206 555 0233'],
  // L5: DevOps Lead → Rosa Diaz
  ['mi.nakamura@contoso.com',   'Mika Nakamura',      'DevOps Lead',                              'Engineering',     'ro.diaz@contoso.com',     'Seattle, WA',      '+1 206 555 0234'],
  // L6: Engineers → Mika Nakamura
  ['et.osei@contoso.com',       'Ethan Osei',         'Site Reliability Engineer',                'Engineering',     'mi.nakamura@contoso.com', 'Seattle, WA',      '+1 206 555 0235'],
  ['pa.wells@contoso.com',      'Pat Wells',          'DevOps Engineer',                          'Engineering',     'mi.nakamura@contoso.com', 'Seattle, WA',      '+1 206 555 0236'],

  // ── QA branch (max depth 5) ───────────────────────────────────────────────────
  // L3: QA Manager → Oliver Brown
  ['fi.adeyemi@contoso.com',    'Fiona Adeyemi',      'QA Manager',                               'Engineering',     'o.brown@contoso.com',     'Seattle, WA',      '+1 206 555 0240'],
  // L4: QA Lead → Fiona Adeyemi
  ['ha.tanaka@contoso.com',     'Haruto Tanaka',      'QA Lead',                                  'Engineering',     'fi.adeyemi@contoso.com',  'Seattle, WA',      '+1 206 555 0241'],
  // L5: QA Engineers → Haruto Tanaka
  ['ch.roberts@contoso.com',    'Charlotte Roberts',  'QA Engineer',                              'Engineering',     'ha.tanaka@contoso.com',   'Seattle, WA',      '+1 206 555 0242'],
  ['ca.ross@contoso.com',       'Casey Ross',         'QA Engineer',                              'Engineering',     'ha.tanaka@contoso.com',   'Seattle, WA',      '+1 206 555 0243'],

  // ══════════════════════════════════════════════════════════════════════════════
  // PRODUCT  (Sofia Patel, VP)
  // ══════════════════════════════════════════════════════════════════════════════

  // L2: Directors → Sofia Patel
  ['mi.davis@contoso.com',      'Michael Davis',      'Director, Consumer Products',              'Product',         's.patel@contoso.com',     'Seattle, WA',      '+1 206 555 0120'],
  ['je.lee@contoso.com',        'Jessica Lee',        'Director, Enterprise Products',            'Product',         's.patel@contoso.com',     'Seattle, WA',      '+1 206 555 0121'],
  ['oc.james@contoso.com',      'Octavia James',      'Director, UX Design',                      'Product',         's.patel@contoso.com',     'Seattle, WA',      '+1 206 555 0122'],

  // ── Consumer Products branch (max depth 6) ───────────────────────────────────
  // L3: Senior Product Lead → Michael Davis
  ['as.nair@contoso.com',       'Arjun Nair',         'Senior Product Lead, Consumer',            'Product',         'mi.davis@contoso.com',    'Seattle, WA',      '+1 206 555 0300'],
  // L4: Product Managers → Arjun Nair
  ['gr.foster@contoso.com',     'Grace Foster',       'Product Manager',                          'Product',         'as.nair@contoso.com',     'Seattle, WA',      '+1 206 555 0301'],
  ['el.brooks@contoso.com',     'Elijah Brooks',      'Product Manager',                          'Product',         'as.nair@contoso.com',     'Seattle, WA',      '+1 206 555 0302'],
  // L5: Associate PMs → Grace / Elijah
  ['le.cho@contoso.com',        'Lena Cho',           'Associate Product Manager',                'Product',         'gr.foster@contoso.com',   'Seattle, WA',      '+1 206 555 0303'],
  ['ri.patel@contoso.com',      'Riley Patel',        'Associate Product Manager',                'Product',         'el.brooks@contoso.com',   'Seattle, WA',      '+1 206 555 0304'],

  // ── Enterprise Products branch (max depth 6) ─────────────────────────────────
  // L3: Senior PM → Jessica Lee
  ['is.santos@contoso.com',     'Isabella Santos',    'Senior Product Manager',                   'Product',         'je.lee@contoso.com',      'Seattle, WA',      '+1 206 555 0310'],
  // L4: PMs → Isabella Santos
  ['ow.campbell@contoso.com',   'Owen Campbell',      'Product Manager',                          'Product',         'is.santos@contoso.com',   'Seattle, WA',      '+1 206 555 0311'],
  ['ni.pham@contoso.com',       'Niamh Pham',         'Product Manager',                          'Product',         'is.santos@contoso.com',   'Seattle, WA',      '+1 206 555 0312'],
  // L5: Associate PMs → Owen / Niamh
  ['de.park@contoso.com',       'Devon Park',         'Associate Product Manager',                'Product',         'ow.campbell@contoso.com', 'Seattle, WA',      '+1 206 555 0313'],
  ['el.martin@contoso.com',     'Elise Martin',       'Associate Product Manager',                'Product',         'ni.pham@contoso.com',     'Seattle, WA',      '+1 206 555 0314'],

  // ── UX Design branch (max depth 6) ───────────────────────────────────────────
  // L3: UX Design Lead → Octavia James
  ['ah.tran@contoso.com',       'Anh Tran',           'UX Design Lead',                           'Product',         'oc.james@contoso.com',    'Seattle, WA',      '+1 206 555 0320'],
  // L4: UX Designer → Anh Tran
  ['li.stone@contoso.com',      'Liam Stone',         'UX Designer',                              'Product',         'ah.tran@contoso.com',     'Seattle, WA',      '+1 206 555 0321'],
  // L5: UX Researcher → Liam Stone
  ['no.ross@contoso.com',       'Noah Ross',          'UX Researcher',                            'Product',         'li.stone@contoso.com',    'Seattle, WA',      '+1 206 555 0322'],

  // ══════════════════════════════════════════════════════════════════════════════
  // SALES  (David Williams, VP)
  // ══════════════════════════════════════════════════════════════════════════════

  // L2: Directors → David Williams
  ['r.chen@contoso.com',        'Robert Chen',        'Director, Enterprise Sales',               'Sales',           'd.williams@contoso.com',  'New York, NY',     '+1 212 555 0130'],
  ['a.thompson@contoso.com',    'Aisha Thompson',     'Director, SMB Sales',                      'Sales',           'd.williams@contoso.com',  'New York, NY',     '+1 212 555 0131'],
  ['c.garcia@contoso.com',      'Carlos Garcia',      'Director, Partnerships',                   'Sales',           'd.williams@contoso.com',  'New York, NY',     '+1 212 555 0132'],
  ['ka.ibrahim@contoso.com',    'Karim Ibrahim',      'Director, International Sales',            'Sales',           'd.williams@contoso.com',  'New York, NY',     '+1 212 555 0133'],

  // ── Enterprise Sales — East branch (max depth 7) ─────────────────────────────
  // L3: Senior Sales Manager → Robert Chen
  ['ev.price@contoso.com',      'Evan Price',         'Senior Sales Manager, Enterprise East',    'Sales',           'r.chen@contoso.com',      'New York, NY',     '+1 212 555 0400'],
  // L4: Sales Manager → Evan Price
  ['ni.foster@contoso.com',     'Nia Foster',         'Sales Manager',                            'Sales',           'ev.price@contoso.com',    'New York, NY',     '+1 212 555 0401'],
  // L5: Senior AE + AEs → Nia Foster
  ['ma.reyes@contoso.com',      'Marcus Reyes',       'Senior Account Executive',                 'Sales',           'ni.foster@contoso.com',   'New York, NY',     '+1 212 555 0402'],
  ['bj.porter@contoso.com',     'Billie Porter',      'Account Executive',                        'Sales',           'ni.foster@contoso.com',   'New York, NY',     '+1 212 555 0403'],
  ['ha.cole@contoso.com',       'Hannah Cole',        'Junior Account Executive',                 'Sales',           'ni.foster@contoso.com',   'New York, NY',     '+1 212 555 0404'],
  // L6: SDRs → Marcus Reyes / Billie Porter
  ['da.bell@contoso.com',       'Dana Bell',          'Sales Development Representative',         'Sales',           'ma.reyes@contoso.com',    'New York, NY',     '+1 212 555 0440'],
  ['ki.moss@contoso.com',       'Kim Moss',           'Sales Development Representative',         'Sales',           'ma.reyes@contoso.com',    'New York, NY',     '+1 212 555 0441'],

  // ── Enterprise Sales — West branch (max depth 7) ─────────────────────────────
  // L3: Senior Sales Manager → Robert Chen
  ['st.hayes@contoso.com',      'Stella Hayes',       'Senior Sales Manager, Enterprise West',    'Sales',           'r.chen@contoso.com',      'New York, NY',     '+1 212 555 0405'],
  // L4: Senior AEs → Stella Hayes
  ['da.ford@contoso.com',       'Damien Ford',        'Senior Account Executive',                 'Sales',           'st.hayes@contoso.com',    'Boston, MA',       '+1 617 555 0406'],
  ['ol.butler@contoso.com',     'Olivia Butler',      'Senior Account Executive',                 'Sales',           'st.hayes@contoso.com',    'New York, NY',     '+1 212 555 0407'],
  // L5: AEs → Damien / Olivia
  ['xa.long@contoso.com',       'Xavier Long',        'Account Executive',                        'Sales',           'da.ford@contoso.com',     'Chicago, IL',      '+1 312 555 0408'],
  ['cy.ford@contoso.com',       'Cy Ford',            'Account Executive',                        'Sales',           'ol.butler@contoso.com',   'New York, NY',     '+1 212 555 0409'],
  // L6: SDRs → Xavier / Cy
  ['jo.wade@contoso.com',       'Jo Wade',            'Sales Development Representative',         'Sales',           'xa.long@contoso.com',     'Chicago, IL',      '+1 312 555 0442'],

  // ── SMB Sales branch (max depth 7) ───────────────────────────────────────────
  // L3: Senior Sales Manager → Aisha Thompson
  ['co.bell@contoso.com',       'Connor Bell',        'Senior Sales Manager, SMB',                'Sales',           'a.thompson@contoso.com',  'New York, NY',     '+1 212 555 0410'],
  // L4: Sales Manager → Connor Bell
  ['au.ross@contoso.com',       'Aurora Ross',        'Sales Manager, SMB',                       'Sales',           'co.bell@contoso.com',     'New York, NY',     '+1 212 555 0411'],
  // L5: AEs → Aurora Ross
  ['ma.powell@contoso.com',     'Marcus Powell',      'Account Executive',                        'Sales',           'au.ross@contoso.com',     'New York, NY',     '+1 212 555 0412'],
  ['le.jenkins@contoso.com',    'Leila Jenkins',      'Account Executive',                        'Sales',           'au.ross@contoso.com',     'Boston, MA',       '+1 617 555 0413'],
  // L6: SDRs → Marcus / Leila
  ['ry.cox@contoso.com',        'Ryan Cox',           'Sales Development Representative',         'Sales',           'ma.powell@contoso.com',   'New York, NY',     '+1 212 555 0443'],
  ['bi.reed@contoso.com',       'Billie Reed',        'Sales Development Representative',         'Sales',           'le.jenkins@contoso.com',  'Boston, MA',       '+1 617 555 0444'],

  // ── Partnerships branch (max depth 6) ────────────────────────────────────────
  // L3: Senior BD Manager → Carlos Garcia
  ['sa.cox@contoso.com',        'Santiago Cox',       'Senior BD Manager',                        'Sales',           'c.garcia@contoso.com',    'New York, NY',     '+1 212 555 0420'],
  // L4: BD Managers → Santiago Cox
  ['na.diaz@contoso.com',       'Natalie Diaz',       'Business Development Manager',             'Sales',           'sa.cox@contoso.com',      'New York, NY',     '+1 212 555 0421'],
  ['ad.murray@contoso.com',     'Adrian Murray',      'Business Development Manager',             'Sales',           'sa.cox@contoso.com',      'Los Angeles, CA',  '+1 310 555 0422'],
  // L5: BD Associates → Natalie / Adrian
  ['fr.luna@contoso.com',       'Frances Luna',       'Business Development Associate',           'Sales',           'na.diaz@contoso.com',     'New York, NY',     '+1 212 555 0445'],
  ['ke.barnes@contoso.com',     'Kenji Barnes',       'Business Development Associate',           'Sales',           'ad.murray@contoso.com',   'Los Angeles, CA',  '+1 310 555 0446'],

  // ── International branch (max depth 6) ───────────────────────────────────────
  // L3: Regional Manager → Karim Ibrahim
  ['su.park@contoso.com',       'Suki Park',          'Regional Manager, APAC',                   'Sales',           'ka.ibrahim@contoso.com',  'New York, NY',     '+1 212 555 0430'],
  // L4: AEs → Suki Park
  ['ro.chan@contoso.com',       'Roland Chan',        'Account Executive, APAC',                  'Sales',           'su.park@contoso.com',     'New York, NY',     '+1 212 555 0431'],
  ['em.khalil@contoso.com',     'Emre Khalil',        'Account Executive, EMEA',                  'Sales',           'su.park@contoso.com',     'New York, NY',     '+1 212 555 0432'],
  // L5: Coordinator → Roland Chan
  ['so.kwon@contoso.com',       'Sofia Kwon',         'Sales Coordinator, APAC',                  'Sales',           'ro.chan@contoso.com',      'New York, NY',     '+1 212 555 0447'],

  // ══════════════════════════════════════════════════════════════════════════════
  // MARKETING  (Rachel Kim, VP)
  // ══════════════════════════════════════════════════════════════════════════════

  // L2: Directors → Rachel Kim
  ['am.clark@contoso.com',      'Amanda Clark',       'Director, Brand',                          'Marketing',       'r.kim@contoso.com',       'Chicago, IL',      '+1 312 555 0140'],
  ['b.wilson@contoso.com',      'Brian Wilson',       'Director, Digital Marketing',              'Marketing',       'r.kim@contoso.com',       'Chicago, IL',      '+1 312 555 0141'],
  ['n.davis@contoso.com',       'Nicole Davis',       'Director, Content',                        'Marketing',       'r.kim@contoso.com',       'Chicago, IL',      '+1 312 555 0142'],

  // ── Brand branch (max depth 6) ────────────────────────────────────────────────
  // L3: Brand Marketing Manager → Amanda Clark
  ['re.ward@contoso.com',       'Rebecca Ward',       'Brand Marketing Manager',                  'Marketing',       'am.clark@contoso.com',    'Chicago, IL',      '+1 312 555 0500'],
  // L4: Senior Designers / Specialists → Rebecca Ward
  ['ja.riley@contoso.com',      'Jackson Riley',      'Senior Brand Designer',                    'Marketing',       're.ward@contoso.com',     'Chicago, IL',      '+1 312 555 0501'],
  ['lu.morgan@contoso.com',     'Lucas Morgan',       'Marketing Specialist',                     'Marketing',       're.ward@contoso.com',     'Chicago, IL',      '+1 312 555 0502'],
  // L5: ICs → Jackson Riley / Lucas Morgan
  ['ai.cooper@contoso.com',     'Aisha Cooper',       'Brand Strategist',                         'Marketing',       'ja.riley@contoso.com',    'Chicago, IL',      '+1 312 555 0503'],
  ['ca.gray@contoso.com',       'Carmen Gray',        'Marketing Coordinator',                    'Marketing',       'lu.morgan@contoso.com',   'Chicago, IL',      '+1 312 555 0504'],

  // ── Digital Marketing branch (max depth 6) ───────────────────────────────────
  // L3: Digital Marketing Manager → Brian Wilson
  ['im.turner@contoso.com',     'Imani Turner',       'Digital Marketing Manager',                'Marketing',       'b.wilson@contoso.com',    'Chicago, IL',      '+1 312 555 0510'],
  // L4: SEO Lead → Imani Turner
  ['et.parker@contoso.com',     'Ethan Parker',       'SEO Lead',                                 'Marketing',       'im.turner@contoso.com',   'Chicago, IL',      '+1 312 555 0511'],
  // L5: Specialists → Ethan Parker
  ['pe.evans@contoso.com',      'Penelope Evans',     'Paid Media Specialist',                    'Marketing',       'et.parker@contoso.com',   'Chicago, IL',      '+1 312 555 0512'],
  ['yu.nakamura@contoso.com',   'Yuki Nakamura',      'SEO Specialist',                           'Marketing',       'et.parker@contoso.com',   'Chicago, IL',      '+1 312 555 0513'],

  // ── Content branch (max depth 6) ─────────────────────────────────────────────
  // L3: Content Lead → Nicole Davis
  ['mi.edwards@contoso.com',    'Miles Edwards',      'Content Lead',                             'Marketing',       'n.davis@contoso.com',     'Chicago, IL',      '+1 312 555 0520'],
  // L4: Senior Writers → Miles Edwards
  ['so.collins@contoso.com',    'Sofia Collins',      'Senior Content Writer',                    'Marketing',       'mi.edwards@contoso.com',  'Chicago, IL',      '+1 312 555 0521'],
  ['be.stewart@contoso.com',    'Benjamin Stewart',   'Content Writer',                           'Marketing',       'mi.edwards@contoso.com',  'Chicago, IL',      '+1 312 555 0522'],
  // L5: Junior Writers → Sofia / Benjamin
  ['as.morgan@contoso.com',     'Ashley Morgan',      'Junior Content Writer',                    'Marketing',       'so.collins@contoso.com',  'Chicago, IL',      '+1 312 555 0523'],
  ['co.hayes@contoso.com',      'Connor Hayes',       'Junior Content Writer',                    'Marketing',       'be.stewart@contoso.com',  'Chicago, IL',      '+1 312 555 0524'],

  // ══════════════════════════════════════════════════════════════════════════════
  // FINANCE  (Thomas Brown, VP)
  // ══════════════════════════════════════════════════════════════════════════════

  // L2: Directors → Thomas Brown
  ['mi.wright@contoso.com',     'Michelle Wright',    'Director, Financial Planning & Analysis',  'Finance',         't.brown@contoso.com',     'Chicago, IL',      '+1 312 555 0150'],
  ['ch.anderson@contoso.com',   'Chris Anderson',     'Director, Accounting',                     'Finance',         't.brown@contoso.com',     'Chicago, IL',      '+1 312 555 0151'],

  // ── FP&A branch (max depth 7) ─────────────────────────────────────────────────
  // L3: Finance Manager → Michelle Wright
  ['ga.sanchez@contoso.com',    'Gabriel Sanchez',    'Finance Manager, FP&A',                    'Finance',         'mi.wright@contoso.com',   'Chicago, IL',      '+1 312 555 0600'],
  // L4: Senior Analysts → Gabriel Sanchez
  ['ev.morris@contoso.com',     'Evelyn Morris',      'Senior Financial Analyst',                 'Finance',         'ga.sanchez@contoso.com',  'Chicago, IL',      '+1 312 555 0601'],
  ['el.rogers@contoso.com',     'Elliot Rogers',      'Senior Financial Analyst',                 'Finance',         'ga.sanchez@contoso.com',  'Chicago, IL',      '+1 312 555 0602'],
  // L5: Analysts → Evelyn / Elliot
  ['fa.cook@contoso.com',       'Fatima Cook',        'Financial Analyst',                        'Finance',         'ev.morris@contoso.com',   'Chicago, IL',      '+1 312 555 0603'],
  ['jo.hayashi@contoso.com',    'Jordan Hayashi',     'Financial Analyst',                        'Finance',         'ev.morris@contoso.com',   'Chicago, IL',      '+1 312 555 0604'],
  ['lc.mendoza@contoso.com',    'Luca Mendoza',       'Junior Financial Analyst',                 'Finance',         'el.rogers@contoso.com',   'Chicago, IL',      '+1 312 555 0605'],
  // L6: Junior Analyst → Luca Mendoza
  ['te.liu@contoso.com',        'Terry Liu',          'Financial Analyst Intern',                 'Finance',         'lc.mendoza@contoso.com',  'Chicago, IL',      '+1 312 555 0606'],

  // ── Accounting branch (max depth 7) ──────────────────────────────────────────
  // L3: Accounting Manager → Chris Anderson
  ['vi.flores@contoso.com',     'Victoria Flores',    'Accounting Manager',                       'Finance',         'ch.anderson@contoso.com', 'Chicago, IL',      '+1 312 555 0610'],
  // L4: Senior Accountant + Treasury Lead → Victoria Flores
  ['se.washington@contoso.com', 'Sebastian Washington','Senior Accountant',                       'Finance',         'vi.flores@contoso.com',   'Chicago, IL',      '+1 312 555 0611'],
  ['cl.nwosu@contoso.com',      'Clarity Nwosu',      'Treasury Analyst Lead',                    'Finance',         'vi.flores@contoso.com',   'Chicago, IL',      '+1 312 555 0612'],
  // L5: Accountants → Sebastian / Clarity
  ['mo.hill@contoso.com',       'Moana Hill',         'Accountant',                               'Finance',         'se.washington@contoso.com','Chicago, IL',     '+1 312 555 0613'],
  ['fe.solomon@contoso.com',    'Felix Solomon',      'Accountant',                               'Finance',         'cl.nwosu@contoso.com',    'Chicago, IL',      '+1 312 555 0614'],
  // L6: Junior → Felix Solomon
  ['ma.lin@contoso.com',        'Mae Lin',            'Junior Accountant',                        'Finance',         'fe.solomon@contoso.com',  'Chicago, IL',      '+1 312 555 0615'],

  // ══════════════════════════════════════════════════════════════════════════════
  // HUMAN RESOURCES  (Lisa Martinez, VP)
  // ══════════════════════════════════════════════════════════════════════════════

  // L2: Directors → Lisa Martinez
  ['je.thomas@contoso.com',     'Jennifer Thomas',    'Director, Talent Acquisition',             'Human Resources', 'l.martinez@contoso.com',  'Chicago, IL',      '+1 312 555 0160'],
  ['ry.jackson@contoso.com',    'Ryan Jackson',       'Director, People Operations',              'Human Resources', 'l.martinez@contoso.com',  'Chicago, IL',      '+1 312 555 0161'],
  ['pa.white@contoso.com',      'Patricia White',     'Director, Learning & Development',         'Human Resources', 'l.martinez@contoso.com',  'Chicago, IL',      '+1 312 555 0162'],

  // ── Talent Acquisition branch (max depth 6) ───────────────────────────────────
  // L3: TA Manager → Jennifer Thomas
  ['ab.green@contoso.com',      'Abigail Green',      'Talent Acquisition Manager',               'Human Resources', 'je.thomas@contoso.com',   'Chicago, IL',      '+1 312 555 0700'],
  // L4: Senior Recruiters → Abigail Green
  ['br.adams@contoso.com',      'Brandon Adams',      'Senior Technical Recruiter',               'Human Resources', 'ab.green@contoso.com',    'Chicago, IL',      '+1 312 555 0701'],
  ['ty.chen@contoso.com',       'Tyler Chen',         'Technical Recruiter',                      'Human Resources', 'ab.green@contoso.com',    'Chicago, IL',      '+1 312 555 0702'],
  // L5: Recruiters / Coordinators → Brandon / Tyler
  ['pr.nelson@contoso.com',     'Priya Nelson',       'Recruiter',                                'Human Resources', 'br.adams@contoso.com',    'Chicago, IL',      '+1 312 555 0703'],
  ['wi.baker@contoso.com',      'William Baker',      'Recruiting Coordinator',                   'Human Resources', 'ty.chen@contoso.com',     'Chicago, IL',      '+1 312 555 0704'],

  // ── People Operations branch (max depth 6) ────────────────────────────────────
  // L3: HRBP Lead → Ryan Jackson
  ['ha.mitchell@contoso.com',   'Harrison Mitchell',  'HRBP Lead',                                'Human Resources', 'ry.jackson@contoso.com',  'Chicago, IL',      '+1 312 555 0710'],
  // L4: Senior HRBPs → Harrison Mitchell
  ['so.carter@contoso.com',     'Sophia Carter',      'Senior HR Business Partner',               'Human Resources', 'ha.mitchell@contoso.com', 'Chicago, IL',      '+1 312 555 0711'],
  ['ma.griffith@contoso.com',   'Maya Griffith',      'HR Business Partner',                      'Human Resources', 'ha.mitchell@contoso.com', 'Chicago, IL',      '+1 312 555 0712'],
  // L5: Coordinators → Sophia / Maya
  ['am.perez@contoso.com',      'Amara Perez',        'HR Coordinator',                           'Human Resources', 'so.carter@contoso.com',   'Chicago, IL',      '+1 312 555 0713'],
  ['si.oduya@contoso.com',      'Simon Oduya',        'HR Coordinator',                           'Human Resources', 'ma.griffith@contoso.com', 'Chicago, IL',      '+1 312 555 0714'],

  // ── L&D branch (max depth 7) ──────────────────────────────────────────────────
  // L3: L&D Lead → Patricia White
  ['ca.roberts@contoso.com',    'Caleb Roberts',      'L&D Lead',                                 'Human Resources', 'pa.white@contoso.com',    'Chicago, IL',      '+1 312 555 0720'],
  // L4: L&D Coordinator → Caleb Roberts
  ['vi.walker@contoso.com',     'Violet Walker',      'L&D Coordinator',                          'Human Resources', 'ca.roberts@contoso.com',  'Chicago, IL',      '+1 312 555 0721'],
  // L5: L&D Specialist → Violet Walker
  ['no.walsh@contoso.com',      'Nora Walsh',         'L&D Specialist',                           'Human Resources', 'vi.walker@contoso.com',   'Chicago, IL',      '+1 312 555 0722'],
  // L6: L&D Assistant → Nora Walsh
  ['av.brooks@contoso.com',     'Avery Brooks',       'L&D Assistant',                            'Human Resources', 'no.walsh@contoso.com',    'Chicago, IL',      '+1 312 555 0723'],

  // ══════════════════════════════════════════════════════════════════════════════
  // OPERATIONS  (James Taylor, VP)
  // ══════════════════════════════════════════════════════════════════════════════

  // L2: Directors → James Taylor
  ['sa.moore@contoso.com',      'Sarah Moore',        'Director, Information Technology',         'Operations',      'j.taylor@contoso.com',    'Austin, TX',       '+1 512 555 0170'],
  ['da.harris@contoso.com',     'Daniel Harris',      'Director, Customer Success',               'Operations',      'j.taylor@contoso.com',    'Austin, TX',       '+1 512 555 0171'],
  ['ge.wilson@contoso.com',     'George Wilson',      'Director, Facilities',                     'Operations',      'j.taylor@contoso.com',    'Austin, TX',       '+1 512 555 0172'],

  // ── IT branch (max depth 7) ───────────────────────────────────────────────────
  // L3: IT Manager → Sarah Moore
  ['de.young@contoso.com',      'Derek Young',        'IT Manager',                               'Operations',      'sa.moore@contoso.com',    'Austin, TX',       '+1 512 555 0800'],
  // L4: IT Lead → Derek Young
  ['to.allen@contoso.com',      'Tomás Allen',        'IT Lead',                                  'Operations',      'de.young@contoso.com',    'Austin, TX',       '+1 512 555 0801'],
  // L5: Engineers → Tomás Allen
  ['ch.king@contoso.com',       'Chelsea King',       'Senior IT Engineer',                       'Operations',      'to.allen@contoso.com',    'Austin, TX',       '+1 512 555 0802'],
  ['ja.wade@contoso.com',       'Jada Wade',          'IT Engineer',                              'Operations',      'to.allen@contoso.com',    'Austin, TX',       '+1 512 555 0803'],
  // L6: IT Support Technician → Chelsea King
  ['ta.williams@contoso.com',   'Talia Williams',     'IT Support Technician',                    'Operations',      'ch.king@contoso.com',     'Austin, TX',       '+1 512 555 0804'],

  // ── Customer Success branch (max depth 7) ─────────────────────────────────────
  // L3: CS Senior Manager → Daniel Harris
  ['ze.wright@contoso.com',     'Zena Wright',        'CS Senior Manager',                        'Operations',      'da.harris@contoso.com',   'Austin, TX',       '+1 512 555 0810'],
  // L4: CS Managers → Zena Wright
  ['co.scott@contoso.com',      'Colin Scott',        'CS Manager, West',                         'Operations',      'ze.wright@contoso.com',   'Austin, TX',       '+1 512 555 0811'],
  ['ph.brooks@contoso.com',     'Phoenix Brooks',     'CS Manager, East',                         'Operations',      'ze.wright@contoso.com',   'Austin, TX',       '+1 512 555 0812'],
  // L5: CSMs / CS Associates → Colin / Phoenix
  ['si.adams@contoso.com',      'Simone Adams',       'Senior Customer Success Manager',          'Operations',      'co.scott@contoso.com',    'Austin, TX',       '+1 512 555 0813'],
  ['ro.turner@contoso.com',     'Roland Turner',      'Customer Success Associate',               'Operations',      'co.scott@contoso.com',    'Austin, TX',       '+1 512 555 0814'],
  ['ik.adaora@contoso.com',     'Ikenna Adaora',      'Customer Success Associate',               'Operations',      'ph.brooks@contoso.com',   'Austin, TX',       '+1 512 555 0815'],
  ['te.nguyen@contoso.com',     'Teresa Nguyen',      'Customer Success Manager',                 'Operations',      'ph.brooks@contoso.com',   'Austin, TX',       '+1 512 555 0816'],
  // L6: CS Coordinator → Simone Adams
  ['pi.jones@contoso.com',      'Pilar Jones',        'Customer Success Coordinator',             'Operations',      'si.adams@contoso.com',    'Austin, TX',       '+1 512 555 0817'],

  // ── Facilities branch (max depth 6) ──────────────────────────────────────────
  // L3: Facilities Manager → George Wilson
  ['pa.james@contoso.com',      'Patricia James',     'Facilities Manager',                       'Operations',      'ge.wilson@contoso.com',   'Austin, TX',       '+1 512 555 0820'],
  // L4: Senior Facilities Coordinator → Patricia James
  ['sa.watson@contoso.com',     'Samuel Watson',      'Senior Facilities Coordinator',            'Operations',      'pa.james@contoso.com',    'Austin, TX',       '+1 512 555 0821'],
  // L5: Facilities Technician → Samuel Watson
  ['ed.gray@contoso.com',       'Eddie Gray',         'Facilities Technician',                    'Operations',      'sa.watson@contoso.com',   'Austin, TX',       '+1 512 555 0822'],

  // ══════════════════════════════════════════════════════════════════════════════
  // LEGAL  (Victoria Santos, VP)
  // ══════════════════════════════════════════════════════════════════════════════

  // L2: General Counsel → Victoria Santos
  ['ch.adeyemi@contoso.com',    'Christopher Adeyemi','General Counsel',                          'Legal',           'vl.santos@contoso.com',   'Chicago, IL',      '+1 312 555 0901'],
  // L3: Senior Associates / Compliance Manager → Christopher Adeyemi
  ['ra.holt@contoso.com',       'Rachel Holt',        'Senior Legal Associate',                   'Legal',           'ch.adeyemi@contoso.com',  'Chicago, IL',      '+1 312 555 0902'],
  ['mi.cross@contoso.com',      'Miguel Cross',       'Compliance Manager',                       'Legal',           'ch.adeyemi@contoso.com',  'Chicago, IL',      '+1 312 555 0903'],
  // L4: Legal Coordinator / Compliance Analyst → Rachel / Miguel
  ['sa.ogden@contoso.com',      'Samantha Ogden',     'Legal Coordinator',                        'Legal',           'ra.holt@contoso.com',     'Chicago, IL',      '+1 312 555 0904'],
  ['cu.wu@contoso.com',         'Curtis Wu',          'Compliance Analyst',                       'Legal',           'mi.cross@contoso.com',    'Chicago, IL',      '+1 312 555 0905'],
  // L5: Legal Assistant → Samantha Ogden
  ['ju.fox@contoso.com',        'Jules Fox',          'Legal Assistant',                          'Legal',           'sa.ogden@contoso.com',    'Chicago, IL',      '+1 312 555 0906'],

  // ── Special accounts for test coverage ───────────────────────────────────────
  ['di.former1@contoso.com',    'Diana Former',       'Former Account Executive',                 'Sales',           'd.williams@contoso.com',  'New York, NY',     '',               false],
  ['ex.employee2@contoso.com',  'Edward Ex',          'Former Engineer',                          'Engineering',     'm.johnson@contoso.com',   'Seattle, WA',      '',               false],
  ['guest.partner@fabrikam.com','Alex Partner',       'Partner Consultant',                       'Operations',      'j.taylor@contoso.com',    'Remote',           '',               true, 'Guest'],
  ['guest.vendor@woodgrove.com','Sam Vendor',         'Vendor Specialist',                        'Engineering',     'm.johnson@contoso.com',   'Remote',           '',               true, 'Guest'],
  ['guest.auditor@contoso.com', 'Priya Auditor',     'External Auditor',                         'Finance',         't.brown@contoso.com',     'Remote',           '',               true, 'Guest'],
];

// The demo CEO (always RAW_BASE's first, L0 entry, at every company size) —
// used to auto-fill the Top-Level User property when Use Demo Data is enabled
export const DEMO_CEO_EMAIL: string = RAW_BASE[0][0];

// ── Generator ─────────────────────────────────────────────────────────────────
// Builds a deterministic dataset of approximately targetCount users by adding
// generated team leads and ICs on top of the hand-authored base data.
function buildRaw(targetCount: number): RawUser[] {
  if (targetCount <= 150) return RAW_BASE;

  const rand   = lcg(0x5EED_BEEF);
  const result: RawUser[] = [...RAW_BASE];
  let   counter = result.length;

  function pick<T>(arr: T[]): T { return arr[Math.floor(rand() * arr.length)]; }
  function pad4(n: number): string { const s = String(n); return '0000'.slice(s.length) + s; }
  function next(): string { return pad4(++counter); }

  // Pass 1 — generate intermediate team-lead managers under existing directors
  const allSlots: MgrSlot[] = [...BASE_SLOTS];
  const leadCount = targetCount >= 1000 ? 80 : 30;

  for (let i = 0; i < leadCount; i++) {
    const parent = pick(BASE_SLOTS);
    const first  = pick(FIRST_NAMES);
    const last   = pick(LAST_NAMES);
    const n      = next();
    const email  = `tl${n}.${last.toLowerCase().slice(0, 8)}@contoso.com`;
    const title  = LEAD_TITLE[parent.dept] ?? 'Manager';
    result.push([email, `${first} ${last}`, title, parent.dept, parent.email, parent.office, `+1 ${parent.area} 555 ${n}`]);
    allSlots.push({ email, dept: parent.dept, office: parent.office, area: parent.area });
  }

  // Pass 2 — generate individual contributors under all managers
  while (result.length < targetCount) {
    const slot   = pick(allSlots);
    const first  = pick(FIRST_NAMES);
    const last   = pick(LAST_NAMES);
    const n      = next();
    const email  = `ic${n}.${last.toLowerCase().slice(0, 8)}@contoso.com`;
    const titles = IC_TITLES[slot.dept] ?? ['Specialist'];
    result.push([email, `${first} ${last}`, pick(titles), slot.dept, slot.email, slot.office, `+1 ${slot.area} 555 ${n}`]);
  }

  return result;
}

// ── Public type ───────────────────────────────────────────────────────────────
export type MockCompanySize = 150 | 500 | 1000;

// ── MockGraphService ──────────────────────────────────────────────────────────
// Supplies generated raw data through the base class's _fetchRawData hook, so
// filtering, manager bridging, cycle breaking, lookups, tree building and the
// updateFilterOptions / refresh / getLastLoaded API all reuse GraphService.
export class MockGraphService extends GraphService {
  private _size: MockCompanySize;

  constructor(size: MockCompanySize = 150, filterOptions: IUserFilterOptions = {}) {
    super(null, 'https://localhost', undefined, 'auto', filterOptions);
    this._size = size;
  }

  // Demo data is generated locally — never use the persistent cache
  protected _cacheKey(): string | null {
    return null;
  }

  protected async _fetchRawData(): Promise<IRawUserData> {
    await delay(500);   // simulated network latency
    const raw = buildRaw(this._size);
    const data: IRawUserData = {
      users: [], managerIds: new Map(), dottedIds: new Map(), objectIds: new Map(),
      noPhotoIds: new Set(), fetchedAt: Date.now(),
    };

    for (const [email, displayName, jobTitle, department, managerEmail, officeLocation, phone, accountEnabled, userType] of raw) {
      data.users.push({
        id:                email,
        displayName,
        mail:              email,
        jobTitle,
        department,
        officeLocation,
        mobilePhone:       '',
        businessPhones:    phone ? [phone] : [],
        userPrincipalName: email,
        accountEnabled:    accountEnabled !== false,
        userType:          userType || 'Member',
        customAttributes:  buildDemoCustomAttributes(email, department, officeLocation, userType === 'Guest'),
      });
      if (managerEmail) data.managerIds.set(email, managerEmail);
      data.noPhotoIds.add(email);   // demo users have no photos — UI shows initials
    }
    data.users.sort((a, b) => a.displayName.localeCompare(b.displayName));

    // Dotted-line (secondary manager) demo relationships
    const DOTTED: Array<[string, string]> = [
      ['sa.yip@contoso.com',      'mi.davis@contoso.com'],   // Backend tech lead ↔ Consumer Products director
      ['ah.tran@contoso.com',     'e.rodriguez@contoso.com'], // UX Design lead ↔ Frontend director
      ['ga.sanchez@contoso.com',  'd.williams@contoso.com'],  // FP&A manager ↔ VP Sales
    ];
    for (const [reportEmail, mgrEmail] of DOTTED) data.dottedIds.set(reportEmail, mgrEmail);

    return data;
  }

  public getUserPhoto(_userId: string): Promise<string | null> {
    return Promise.resolve(null);
  }

  public async findUser(identifier: string): Promise<IGraphUser | null> {
    if (!identifier || !identifier.trim()) return null;
    return super.findUser(identifier);
  }

  public async getPresence(userIds?: string[]): Promise<Map<string, PresenceAvailability>> {
    const ids = userIds
      ? userIds.map(id => id.toLowerCase())
      : (await this.getAllUsers()).map(u => u.id);
    await delay(200);
    const presence = new Map<string, PresenceAvailability>();
    for (const id of ids) presence.set(id, PRESENCE_POOL[hashCode(id) % PRESENCE_POOL.length]);
    return presence;
  }
}

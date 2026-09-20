import React, { useState, useEffect, useRef, useMemo, useContext, createContext } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, CartesianGrid, Legend,
} from "recharts";
import {
  Lock,
  Info,
  Clock,
  AlertCircle,
  CheckCircle2,
  FileSpreadsheet,
  ChevronLeft,
  ChevronRight,
  Check,
  Circle,
  Save,
  User,
  Users as UsersIcon,
  ChevronDown,
  ChevronUp,
  Plus,
  Minus,
  ShieldCheck,
  LogOut,
  RefreshCcw,
  ArrowLeft,
  Download,
  LayoutDashboard,
  UploadCloud,
  Pencil,
  Trash2,
  X,
  ClipboardList,
  FileClock,
  Building2,
  Search,
  Printer,
  TreeDeciduous,
  Sprout,
  Leaf,
  Mountain,
  Link2Off,
} from "lucide-react";

/* ========================================================================
   LOGIN & ROUTING
   Same login idea as the frontline app: no company email needed.
   Username = FirstName + first 3 letters of LastName. Password = DDMMYYYY.

   Each Business Unit is treated as its own site (own URL path in
   production, own hash-route here) so one BU's code can never query
   another BU's people or data — not even by accident. Master Admin has
   a separate site of its own that aggregates read-only across BUs.
========================================================================= */

const BUS = ["UTSE", "UTWC", "UTLC", "UTCT", "TPCTH"];

// slug used in the URL path/hash <-> the BU code stored on each user
const BU_SLUGS = {
  utse: "UTSE",
  utwc: "UTWC",
  utlc: "UTLC",
  utct: "UTCT",
  tpcth: "TPCTH",
};
const SLUG_BY_BU = Object.fromEntries(Object.entries(BU_SLUGS).map(([slug, bu]) => [bu, slug]));

// Full legal names — short codes (above) stay the internal identifier/slug;
// the full name is used wherever a formal company reference is shown.
const BU_FULL_NAME = {
  UTSE: "Unithai Shipyard and Engineering Limited",
  UTLC: "United Thai Logistics Company Limited",
  UTCT: "United Thai Shipping Corporation Limited (Branch)",
  UTWC: "United Thai Warehouse Company Limited",
  TPCTH: "TPC Group (Thailand) Limited",
};
// The Group entity isn't one of the 4 operating companies — its people are
// typically functional managers who assess someone at an operating company
// without needing an evaluation of their own through this system, and
// without sitting in the direct reporting line of anyone they assess.
const GROUP_ENTITY = "TPCTH";

const ROLE_LABEL = {
  staff: "Staff",
  po_admin: "P&O Admin",
  master_admin: "Master Admin",
};

/* ========================================================================
   SEED ROSTER — the starting People data. In production this table lives
   in a real database; here it's the fallback used the first time the
   prototype runs, then overtaken by whatever Master Admin imports/edits
   (persisted in shared storage as the live "roster").

   role: "staff" covers everyone doing frontline work — a "manager" is
   just a staff member who has other staff pointing managerId at them.
   That's computed dynamically (hasReports), not a fixed account type —
   which is what lets one person hold both an Employee view of their own
   journey and a Manager view of their team from a single login.
========================================================================= */

const SEED_ROSTER = [
  { id: "master1", firstName: "Master", lastName: "Admin01", username: "Master01", password: "cnb2026x", role: "master_admin", bu: null, buList: null, employeeId: "TH00-24000" },
  { id: "master2", firstName: "Master", lastName: "Admin02", username: "Master02", password: "cnb2026x", role: "master_admin", bu: null, buList: null, employeeId: "TH00-24001" },
  { id: "master3", firstName: "Master", lastName: "Admin03", username: "Master03", password: "cnb2026x", role: "master_admin", bu: null, buList: null, employeeId: "TH00-24002" },

  // P&O — same coverage as before: Punjamaporn on UTSE alone, Charinee across three
  { id: "po-punjamaorn", firstName: "Punjamaporn", lastName: "Srisuwan", username: "Punjamapornsri", password: "testonly", role: "po_admin", bu: "UTSE", buList: null, employeeId: "TH00-25001" },
  { id: "po-charinee", firstName: "Charinee", lastName: "Hemarajata", username: "Charineehem", password: "testonly", role: "po_admin", bu: null, buList: ["UTLC", "UTWC", "UTCT"], employeeId: "TH00-25002" },

  // UTSE — covers Pattern C (Username1), B (Username2), A (Username3/4/5),
  // D (Username6). Username5 is pre-assigned as the delegate on Username4's
  // review, so the delegate flow is testable immediately. Username3 has
  // Username13 (TPCTH) as a functional manager, to test the matrix-reporting
  // case too.
  { id: "utse-sr", firstName: "Username", lastName: "1", username: "Username1", password: "testonly", role: "staff", bu: "UTSE", department: "Operations", jobGrade: "JG8", designation: "Senior Director", managerId: null, employeeId: "TH08-25001", needsEvaluation: false },
  { id: "utse-mgr", firstName: "Username", lastName: "2", username: "Username2", password: "testonly", role: "staff", bu: "UTSE", department: "Operations", jobGrade: "M2", designation: "Operations Manager", managerId: "utse-sr", employeeId: "TH08-25002" },
  { id: "utse-e1", firstName: "Username", lastName: "3", username: "Username3", password: "testonly", role: "staff", bu: "UTSE", department: "Operations", jobGrade: "S2", designation: "Frontline Officer", managerId: "utse-mgr", employeeId: "TH08-25003", functionalManagerId: "tpcg-fm1" },
  { id: "utse-e2", firstName: "Username", lastName: "4", username: "Username4", password: "testonly", role: "staff", bu: "UTSE", department: "Operations", jobGrade: "S1", designation: "Frontline Officer", managerId: "utse-mgr", employeeId: "TH08-25004", reviewDelegateId: "utse-e3" },
  { id: "utse-e3", firstName: "Username", lastName: "5", username: "Username5", password: "testonly", role: "staff", bu: "UTSE", department: "Operations", jobGrade: "S3", designation: "Senior Officer", managerId: "utse-mgr", employeeId: "TH08-25005" },
  { id: "utse-dh", firstName: "Username", lastName: "6", username: "Username6", password: "testonly", role: "staff", bu: "UTSE", department: "Operations", jobGrade: "JG7", designation: "Principal Advisor", managerId: "utse-sr", employeeId: "TH08-25006", needsEvaluation: false },

  // UTLC — Pattern B (Username7) + A (Username8), for Charinee's scope
  { id: "utlc-mgr", firstName: "Username", lastName: "7", username: "Username7", password: "testonly", role: "staff", bu: "UTLC", department: "Operations", jobGrade: "M2", designation: "Operations Manager", managerId: null, employeeId: "TH09-25001" },
  { id: "utlc-e1", firstName: "Username", lastName: "8", username: "Username8", password: "testonly", role: "staff", bu: "UTLC", department: "Operations", jobGrade: "S2", designation: "Frontline Officer", managerId: "utlc-mgr", employeeId: "TH09-25002" },

  // UTWC — Pattern B (Username9) + A (Username10), for Charinee's scope
  { id: "utwc-mgr", firstName: "Username", lastName: "9", username: "Username9", password: "testonly", role: "staff", bu: "UTWC", department: "Operations", jobGrade: "M2", designation: "Operations Manager", managerId: null, employeeId: "TH10-25001" },
  { id: "utwc-e1", firstName: "Username", lastName: "10", username: "Username10", password: "testonly", role: "staff", bu: "UTWC", department: "Operations", jobGrade: "S2", designation: "Frontline Officer", managerId: "utwc-mgr", employeeId: "TH10-25002" },

  // UTCT — Pattern B (Username11) + A (Username12), for Charinee's scope
  { id: "utsc-mgr", firstName: "Username", lastName: "11", username: "Username11", password: "testonly", role: "staff", bu: "UTCT", department: "Operations", jobGrade: "M2", designation: "Operations Manager", managerId: null, employeeId: "TH11-25001" },
  { id: "utsc-e1", firstName: "Username", lastName: "12", username: "Username12", password: "testonly", role: "staff", bu: "UTCT", department: "Operations", jobGrade: "S2", designation: "Frontline Officer", managerId: "utsc-mgr", employeeId: "TH11-25002" },

  // TPC Group (Thailand) Limited — Username3's functional manager (matrix-reporting test)
  { id: "tpcg-fm1", firstName: "Username", lastName: "13", username: "Username13", password: "testonly", role: "staff", bu: "TPCTH", department: "Group Operations", jobGrade: "JG9", designation: "Group Functional Director", managerId: null, employeeId: "TPCTH-25001", needsEvaluation: false },
];

function hasReports(personId, roster) {
  // Only counts direct reports who actually need to be evaluated — someone
  // set up as a pure reviewer/delegate (needsEvaluation: false) isn't
  // "managed" in the evaluation sense, so doesn't trigger a My Team tab.
  return roster.some((p) => p.role === "staff" && p.managerId === personId && p.needsEvaluation !== false);
}
function directReports(personId, roster) {
  return roster.filter((p) => p.managerId === personId && p.needsEvaluation !== false);
}
function adminBUs(person) {
  if (!person) return [];
  if (person.buList && person.buList.length) return person.buList;
  if (person.bu) return [person.bu];
  return [];
}

const ROSTER_KEY = "gva-roster-v10"; // bumped: generic Username1-13 test roster, same pattern coverage as before
const LOG_KEY = "gva-export-log-v3"; // bumped alongside the v10 reset — clears the activity log too
const CYCLE_KEY = "gva-current-cycle-v2"; // bumped: v1 had a "2026 H1" naming that's now retired
const CYCLES_LIST_KEY = "gva-cycles-list-v2";
const CYCLE_WINDOW_KEY = "gva-cycle-window-v1";
const DEFAULT_CYCLE = "2026";

async function loadRoster() {
  try {
    const res = await window.storage.get(ROSTER_KEY, true);
    return res && res.value ? JSON.parse(res.value) : null;
  } catch {
    return null;
  }
}
async function saveRoster(roster) {
  try {
    await window.storage.set(ROSTER_KEY, JSON.stringify(roster), true);
  } catch (e) {
    console.error("roster save failed", e);
  }
}
async function loadLogs() {
  try {
    const res = await window.storage.get(LOG_KEY, true);
    return res && res.value ? JSON.parse(res.value) : [];
  } catch {
    return [];
  }
}
async function appendLog(actorLabel, action, detail) {
  try {
    const logs = await loadLogs();
    logs.unshift({ id: `log-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`, ts: new Date().toISOString(), actor: actorLabel, action, detail });
    await window.storage.set(LOG_KEY, JSON.stringify(logs.slice(0, 500)), true);
  } catch (e) {
    console.error("log append failed", e);
  }
}

async function loadCurrentCycle() {
  try {
    const res = await window.storage.get(CYCLE_KEY, true);
    return res && res.value ? JSON.parse(res.value) : null;
  } catch {
    return null;
  }
}
async function saveCurrentCycle(cycle) {
  try {
    await window.storage.set(CYCLE_KEY, JSON.stringify(cycle), true);
  } catch (e) {
    console.error("cycle save failed", e);
  }
}
async function loadCyclesList() {
  try {
    const res = await window.storage.get(CYCLES_LIST_KEY, true);
    return res && res.value ? JSON.parse(res.value) : null;
  } catch {
    return null;
  }
}
async function saveCyclesList(list) {
  try {
    await window.storage.set(CYCLES_LIST_KEY, JSON.stringify(list), true);
  } catch (e) {
    console.error("cycles list save failed", e);
  }
}

async function loadCycleWindow() {
  try {
    const res = await window.storage.get(CYCLE_WINDOW_KEY, true);
    return res && res.value ? JSON.parse(res.value) : null;
  } catch {
    return null;
  }
}
async function saveCycleWindow(window_) {
  try {
    await window.storage.set(CYCLE_WINDOW_KEY, JSON.stringify(window_), true);
  } catch (e) {
    console.error("cycle window save failed", e);
  }
}

const RosterCtx = createContext({ roster: SEED_ROSTER, setRoster: () => {}, cycle: DEFAULT_CYCLE, lang: "en", setLang: () => {} });
function useRosterCtx() {
  return useContext(RosterCtx);
}

/* ========================================================================
   TPC BRAND TOKENS (from Brand Guide)
========================================================================= */
const BRAND = {
  primary: "#007788", // Pantone 3145C
  teal: "#00A7B5", // Pantone 7710C
  mint: "#40C1AC", // Pantone 7465C
  deep: "#004F59", // Secondary 3165C
  gray: "#75787B", // Cool Gray 9C
  red: "#E10600", // Awakening Red — signature colour, used sparingly
  bg: "#F6FBFA",
  bgGradient: "radial-gradient(circle at 12% -10%, #EAF7F6 0%, #F6FBFA 45%, #F3F9F8 100%)",
  card: "#FFFFFF",
  line: "#E3EEEC",
};

const TPC_LOGO = "data:image/png;base64,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";

const FONT_IMPORT =
  "@import url('https://fonts.googleapis.com/css2?family=Quicksand:wght@400;500;600;700&family=Noto+Sans+SC:wght@400;500;700&family=Noto+Sans+Thai:wght@400;500;600;700&display=swap');";

// TPC brand type system: Quicksand (English), Sukhumvit Set (Thai — commercial,
// not on Google Fonts, so Noto Sans Thai stands in until it's licensed here),
// Noto Sans SC (Simplified Chinese). Each script falls through to its own
// matching family instead of collapsing to one font.
const FONT_STACK = "'Quicksand', 'Noto Sans Thai', 'Noto Sans SC', sans-serif";

const STORAGE_KEY = "gva-journey-demo-v3"; // bumped alongside the v10 reset — clears every test answer from the previous roster

/* ========================================================================
   CONTENT — pulled from the GVA form
========================================================================= */

const STAGE_COLUMNS = ["Stage", "What it looks like", "If this is you, you'd say…", "What a manager might see"];
const STAGE_COLUMNS_TH = ["ระดับ", "ลักษณะที่ปรากฏ", "ถ้าเป็นคุณ คุณอาจพูดว่า...", "สิ่งที่หัวหน้างานอาจสังเกตเห็น"];

const STAGES_A1 = [
  {
    stage: "Mature",
    looksLike: "Self-cultivation is the substrate from which work emerges. Others learn from your practice.",
    youdSay: "Others come to me asking how I keep showing up the way I do.",
    managerSees: "Practice consistent across 4 quarters; named as a learning anchor by peers.",
  },
  {
    stage: "Established",
    looksLike: "Regular reflective practice. Values consistent across contexts. Learning is sought, not assigned.",
    youdSay: "My habits are mostly there; the values feel like mine, not external rules.",
    managerSees: "8 Living Habits consistent; Stewardship of Life journal entries with named change; coaching shows applied learning.",
  },
  {
    stage: "Developing",
    looksLike: "Building reflective habits. Values awareness emerging. Some adaptation visible.",
    youdSay: "I'm starting to see my patterns; reflection isn't yet automatic.",
    managerSees: "8 Living Habits intermittent; journal entries present but generic; coaching responses positive but not always applied.",
  },
  {
    stage: "Emerging",
    looksLike: "Aware of cultivation as an orientation. Occasional reflection. Habits inconsistent.",
    youdSay: "I know reflection and habits matter; I haven't built the muscle yet.",
    managerSees: "8 Living Habits sporadic; few journal entries; awareness signals without applied pattern.",
  },
];

const STAGES_A2 = [
  {
    stage: "Mature",
    looksLike: "Builds relational systems others use. All 5 Relational Matrix orientations embodied across contexts.",
    youdSay: "Others come to me asking how I build trust across teams.",
    managerSees: "Internal & external networks are strong; all 5 orientations consistent; out-of-team peer aggregates converge.",
  },
  {
    stage: "Established",
    looksLike: "Strong network. Relational orientations consistent. Reciprocity practised.",
    youdSay: "I have working relationships across most teams I touch, and I invest in them.",
    managerSees: "Internal & external networks are balanced; ≥4 of 5 orientations signalled regularly.",
  },
  {
    stage: "Developing",
    looksLike: "Active in team relationships. Beginning cross-team. Some orientations embodied.",
    youdSay: "My network is mostly my own team; I'm starting to reach out beyond.",
    managerSees: "Internal network growing; external networks emerging; some orientations signalled.",
  },
  {
    stage: "Emerging",
    looksLike: "Focused on own role. Limited cross-team relationships. Awareness without active embodiment.",
    youdSay: "I focus mostly on my own work and my immediate team.",
    managerSees: "Network low on cross-team dimensions; few orientations beyond own team.",
  },
];

const STAGES_A3 = [
  {
    stage: "Mature",
    looksLike: "Shapes conditions for others' leadership. Cultivates new leaders. The system becomes smarter because of your contribution.",
    youdSay: "I see myself as building the next layer of leadership, not just leading the current work.",
    managerSees: "All 4 Mindsets evidenced; Mintzberg signals show proactive role dominance; direct evidence of cultivating other leaders.",
  },
  {
    stage: "Established",
    looksLike: "Leads without formal title. Makes the system smarter through influence. Initiative is reliable.",
    youdSay: "I take initiative reliably; my influence reaches beyond my role.",
    managerSees: "3-4 Mindsets consistent; proactive role balance in Mintzberg; coaching effectiveness at Value-Adding+ if applicable.",
  },
  {
    stage: "Developing",
    looksLike: "Takes initiative. Supports others' work. Beginning to influence beyond own role.",
    youdSay: "I'm starting to take ownership beyond just my work; some cycles more than others.",
    managerSees: "2-3 Mindsets evidenced; some initiative visible; not yet consistent across cycles.",
  },
  {
    stage: "Emerging",
    looksLike: "Contributes to own work. Follows direction. Limited initiative beyond role.",
    youdSay: "I focus on getting my own work done well.",
    managerSees: "1-2 Mindsets (typically Management); Individual contributor-focused Mintzberg signals.",
  },
];

const STAGES_A4 = [
  {
    stage: "Beyond Me",
    looksLike: "Self/system distinction has dissolved. Wisdom is generative; others draw from your presence.",
    youdSay: "I act without calculating whether the action serves me.",
    managerSees: "Multiple personal-cost decisions for systemic benefit; peer Think-Through items name your \"presence\".",
  },
  {
    stage: "As Me",
    looksLike: "Integrated self-and-purpose. Role expresses who you are; no tension between self and contribution.",
    youdSay: "My role expresses who I am; I don't feel a split between self and work.",
    managerSees: "IRTF reflections name role and purpose as integrated; decisions weigh both personal alignment and systemic impact.",
  },
  {
    stage: "Through Me",
    looksLike: "Instrument of larger purpose. You see purpose larger than your role; weigh systemic impact alongside personal interest.",
    youdSay: "I see purpose larger than my role; sometimes I set aside personal preference for it.",
    managerSees: "IRTF reflections name purpose larger than role; some decisions evidence personal preference set aside.",
  },
  {
    stage: "By Me",
    looksLike: "Self-as-agent. World responds to your action; you are the cause. Effective and ambitious.",
    youdSay: "I take ownership; I make things happen.",
    managerSees: "OKR achievement strong; Profit.co activity high; \"I made this happen\" framing observed.",
  },
  {
    stage: "To Me",
    looksLike: "Self-oriented. World happens to you; you respond to circumstance. Foundational stage; not a deficit.",
    youdSay: "I respond to what comes; agency is something I'm growing into.",
    managerSees: "IRTF frames situations as external; awareness statements present without yet agentic framing.",
  },
];

const A5_DIMENSIONS = [
  "Self-Cultivation – Individual Level",
  "Role & Relationship – Team Level",
  "Stewardship – Organizational Level",
  "Consciousness – Universe Level",
];

const SNAPSHOT_OPTIONS = [
  { key: "Flourishing", desc: "Strengthens the system by creating the conditions for others, and the whole, to thrive." },
  { key: "Generative", desc: "Creates real value that compounds for the role, the team, and the purpose it serves." },
  { key: "Sustaining", desc: "Maintains the system's functioning, delivering reliably on what is expected." },
  { key: "Depleting", desc: "Draws down the system's capacity, trust or coherence, leaving it weaker." },
];

/* ========================================================================
   SECTIONS (pages)
========================================================================= */

const SECTIONS = [
  { id: "cover", nav: "Overview", track: null },
  {
    id: "a1", nav: "A1 · Self-Cultivation", track: "A", code: "A1",
    title: "Self-Cultivation (Individual Level)",
    chinese: "修身 (Xiu Shen)",
    anchorMeaning: "Self-Cultivation — growth begins inside before it radiates outward.",
    intro: "Your daily habits, values, and approach to learning. How you are growing and developing yourself over time.",
    stages: STAGES_A1,
    confusion: "Just knowing TPC values is not self-cultivation. Self-cultivation is about whether values are LIVED in daily practice, and is visible in how you show up over time.",
    selfPrompts: [
      "What have you noticed about yourself in this dimension of self-cultivation — from feedback you've received, situations you've been in, or patterns you've spotted in your own habits and learning? What invites you to grow and develop to the next stage?",
    ],
    managerPrompts: [
      "What shift have you seen emerging in the dimension of self-cultivation in the Employee?",
      "How could the Employee build on this intentionally next year? And how can you support the Employee's Journey in this dimension?",
    ],
    fields: { self: "a1_self", manager: "a1_manager", radio: "a1_stage" },
  },
  {
    id: "a2", nav: "A2 · Roles & Relationships", track: "A", code: "A2",
    title: "Role & Relationship (Team Level)",
    chinese: "齐家 (Qi Jia)",
    anchorMeaning: "Roles & Relationships — the middle stage, where individual contribution expands into conscious partnership.",
    intro: "How you build relationships and collaborate with your team and colleagues. The quality of how you connect.",
    stages: STAGES_A2,
    confusion: "Simply having many relationships does not translate to a 'Mature' growth stage. Maturity in this dimension is more about depth and others adopting the relational pattern, not breadth alone.",
    selfPrompts: [
      "What have you noticed about yourself in this dimension of Role & Relationship with Teams — in feedback, in situations, or in specific interactions? What invites you to grow and develop to the next stage?",
    ],
    managerPrompts: [
      "What shift have you seen emerging in the dimension of Role & Relationship in the Employee?",
      "How could the Employee build on this intentionally next year? And how can you support the Employee's Journey in this dimension?",
    ],
    fields: { self: "a2_self", manager: "a2_manager", radio: "a2_stage" },
  },
  {
    id: "a3", nav: "A3 · Stewardship", track: "A", code: "A3",
    title: "Stewardship (Organisational Level)",
    chinese: "治国 (Zhi Guo)",
    anchorMeaning: "Leading not for oneself, but for the well-being of all entrusted to one's care.",
    intro: "How you take responsibility for your work and influence outcomes. How you lead, guide, and support others — whether or not you have a formal manager title.",
    stages: STAGES_A3,
    confusion: "Stewardship goes beyond just titles or formal authority. It is about leadership quality regardless of role. The progression is often catalysed by the manager intentionally creating space for the person to lead on a stretch initiative.",
    selfPrompts: [
      "What have you noticed about yourself in this dimension of Stewardship — in moments where you stepped forward, or in patterns of how you respond when something needs leading? What invites you to grow and develop to the next stage?",
    ],
    managerPrompts: [
      "What shift have you seen emerging in the dimension of Stewardship in the Employee?",
      "How could the Employee build on this intentionally next year? And how can you support the Employee's Journey in this dimension?",
    ],
    fields: { self: "a3_self", manager: "a3_manager", radio: "a3_stage" },
  },
  {
    id: "a4", nav: "A4 · Consciousness", track: "A", code: "A4",
    title: "Consciousness (Universe Level)",
    chinese: "天人合一 (Tian Ren He Yi)",
    anchorMeaning: "Unity of Self and the Whole — the outermost sphere; consciousness as lived responsibility for the living system.",
    intro: "How you see your impact beyond yourself and your immediate work. Decisions and actions that consider long-term outcomes for the team, organisation, and wider system.",
    stages: STAGES_A4,
    confusion: "Cycle-over-cycle movement is the signal, not single-cycle judgment. Don't expect a fast movement. Stages are NOT a hierarchy of \"good employees vs bad.\" Each stage is a legitimate place on a long developmental journey.",
    stageRadio: true,
    selfPrompts: [
      "What have you noticed about yourself in this dimension of Consciousness — in moments where you weighed broader impact, or where you noticed your perspective changing? What invites you to grow and develop to the next stage?",
    ],
    managerPrompts: [
      "What shift have you seen emerging in the dimension of Consciousness in the Employee?",
      "How could the Employee build on this intentionally next year? And how can you support the Employee's Journey in this dimension?",
    ],
    fields: { self: "a4_self", manager: "a4_manager", radio: "a4_stage" },
  },
  { id: "a5", nav: "A5 · Growth Plan", track: "A", code: "A5", title: "Growth Plan", starred: true },
  {
    id: "b1-learning", nav: "B1 · Learning", track: "B", pillar: "Person", code: "LEARNING & SELF-CULTIVATION",
    anchorEn: "The ability to learn, adapt from setbacks, and contribute to a learning culture within the organisation.",
    selfLabel: "By Employee",
    selfPrompts: [
      "Give an example that reflects your learning, adaptation, or growth in your assigned role recently, specifying:",
    ],
    managerPrompts: [
      "What development have you seen reflecting the Employee's awareness, learning, and adaptability?",
      "How did this event affect the way the Employee works, for the better?",
    ],
    fields: { self: "b1_learning_self", manager: "b1_learning_manager" },
  },
  {
    id: "b1-integrity", nav: "B1 · Integrity", track: "B", pillar: "Person (managers)", code: "MANAGE WITH INTEGRITY", managerOnly: true,
    anchorEn: "Integrity in management often begins with honesty and sincerity — to the organisation's mission, to colleagues, and to yourself.",
    selfLabel: "By Employee",
    selfPrompts: [
      "Give an example that reflects how you managed your team and made decisions with integrity and sincerity, specifying:",
    ],
    managerPrompts: [
      "What development have you seen reflecting the Employee's integrity and adherence to what's right?",
      "How did this event affect the way the Employee works, for the better?",
    ],
    fields: { self: "b1_integrity_self", manager: "b1_integrity_manager" },
  },
  {
    id: "b1-coaching", nav: "B1 · Coaching", track: "B", pillar: "Person (managers)", code: "MANAGE WITH COACHING", managerOnly: true,
    anchorEn: "Developing others through questions, guidance, mentorship, and creating opportunities for others to grow.",
    selfLabel: "By Employee",
    selfPrompts: [
      "Give an example that reflects how you prioritised people development, through coaching or guidance, specifying:",
    ],
    managerPrompts: [
      "What development have you seen reflecting the Employee's ability to grow into a managerial role?",
      "How did this event affect the way the Employee works, for the better?",
    ],
    fields: { self: "b1_coaching_self", manager: "b1_coaching_manager" },
  },
  {
    id: "b1-connection", nav: "B1 · Connection", track: "B", pillar: "Person (managers)", code: "MANAGE WITH CONNECTION", managerOnly: true,
    anchorEn: "Building genuine relationships and caring for well-being, including fostering a culture of collaboration within the team.",
    selfLabel: "By Employee",
    selfPrompts: [
      "Give an example that reflects how you strengthened bonds and unity within the team, specifying:",
    ],
    managerPrompts: [
      "What development have you seen reflecting the Employee building harmony within the team?",
      "How did this event affect the way the Employee works, for the better?",
    ],
    fields: { self: "b1_connection_self", manager: "b1_connection_manager" },
  },
  {
    id: "b2-collab", nav: "B2 · Collaboration", track: "B", pillar: "Practice & Presence", code: "COLLABORATION & PARTNERSHIP",
    anchorEn: "Supporting the success of others, building good relationships, and contributing beyond your own colleagues.",
    selfLabel: "By Employee",
    selfPrompts: [
      "Give an example that reflects how you supported a colleague's work — especially someone from another team — to a successful outcome, specifying:",
    ],
    managerPrompts: [
      "What development have you seen reflecting collaboration between the Employee and colleagues from other teams?",
      "How did this event affect the way the Employee works, for the better?",
    ],
    fields: { self: "b2_collab_self", manager: "b2_collab_manager" },
  },
  {
    id: "b2-steward", nav: "B2 · Accountability", track: "B", pillar: "Practice & Presence", code: "STEWARDSHIP & ACCOUNTABILITY",
    anchorEn: "Values, taking accountability, and constructive challenge in service of better outcomes.",
    selfLabel: "By Employee",
    selfPrompts: [
      "Give an example that reflects how you took accountability, voiced your view, and carried out an assigned task exceptionally well, specifying:",
    ],
    managerPrompts: ["Has the Employee proactively shared opinions or creative ideas at work? How?"],
    fields: { self: "b2_steward_self", manager: "b2_steward_manager" },
  },
  {
    id: "b2-innovation", nav: "B2 · Innovation", track: "B", pillar: "Practice & Presence", code: "INNOVATION & ENTREPRENEURSHIP",
    anchorEn: "Creating something new, adapting to change, and turning ideas into results that can actually be put into practice.",
    selfLabel: "By Employee",
    selfPrompts: [
      "Give an example that reflects an idea you created for the benefit of the wider organisation, specifying:",
    ],
    managerPrompts: ["Has the Employee created something new beyond their assigned work? How?"],
    fields: { self: "b2_innovation_self", manager: "b2_innovation_manager" },
  },
  { id: "b3", nav: "B3 · Performance", track: "B", pillar: "Performance", code: "DELIVERING RESULTS & IMPACT", fields: { self: "b3_results_self", manager: "b3_results_manager" } },
  { id: "b4", nav: "B4 · Overall", track: "B", pillar: "Overall", code: "OVERALL & SNAPSHOT" },
];

/* ========================================================================
   THAI TRANSLATIONS — sourced from the company's own translated GVA form.
   Keyed by section id; looked up alongside the English SECTIONS content
   rather than duplicating the whole array, so the English structure above
   stays the single source of truth for field names/ids/logic.
========================================================================= */
const SECTIONS_TH = {
  a1: {
    title: "ด้านการทบทวนตนเอง (Self-cultivation – Individual Level)",
    anchorMeaning: "ความเปลี่ยนแปลงที่ดีเริ่มต้นจากแนวคิดภายในของตนเอง ที่สะท้อนออกมาผ่านการทำงานและการปฏิบัติตนต่อผู้อื่น",
    intro: "มุมมองเกี่ยวกับพฤติกรรมการใช้ชีวิต และการเรียนรู้ รวมถึงการพัฒนาตนเองเมื่อเวลาผ่านไป",
    stages: [
      { stage: "ก้าวหน้าอย่างมั่นคง (Mature)", looksLike: "การพัฒนาและทบทวนตนเองเป็นพื้นฐานที่ทำให้งานทุกอย่างเกิดขึ้น และผู้อื่นสามารถเรียนรู้จากแนวทางปฏิบัติของฉันได้", youdSay: "คนอื่นมักถามฉันว่าทำอย่างไรถึงจะรักษาอุดมการณ์แบบนี้ได้", managerSees: "ปฏิบัติตาม 8 พฤติกรรมหลัก (8 Living Habits) ได้อย่างสม่ำเสมอ และผู้อื่นมองเป็นแบบอย่าง" },
      { stage: "พัฒนามั่นคง (Established)", looksLike: "ทบทวนตนเองอย่างสม่ำเสมอ ยึดมั่นในค่านิยม และแสวงหาการเรียนรู้ด้วยตนเองโดยสมัครใจ", youdSay: "ค่านิยมต่างๆกลายเป็นส่วนหนึ่งของฉัน ไม่ใช่กฎที่ถูกกำหนดโดยสังคมภายนอก", managerSees: "ปฏิบัติตาม 8 พฤติกรรมหลัก (8 Living Habits) มากขึ้น มีการสังเกตตนเอง แลกเปลี่ยนบทสนทนากับผู้อื่น และนำไปปฏิบัติในชีวิตจริงอย่างต่อเนื่อง" },
      { stage: "กำลังพัฒนา (Developing)", looksLike: "อยู่ระหว่างการฝึกทบทวนตนเอง เริ่มตระหนักถึงค่านิยมบางอย่าง และเริ่มเห็นแนวคิดการพัฒนาตนเอง", youdSay: "ฉันเริ่มมองเห็นการพัฒนาตนเอง แต่ยังต้องใช้ความพยายามในการทบทวนตนเองให้เป็นนิสัย", managerSees: "ปฏิบัติตาม 8 พฤติกรรมหลัก (8 Living Habits) บ้าง มีการสังเกตตนเอง แลกเปลี่ยนบทสนทนากับผู้อื่น และนำไปปฏิบัติในชีวิตจริงแต่ไม่สม่ำเสมอ" },
      { stage: "อยู่ในขั้นเริ่มต้น (Emerging)", looksLike: "ตระหนักได้ว่าการพัฒนาและทบทวนตนเองเป็นเรื่องสำคัญ มีการทบทวนตนเองเป็นครั้งคราว แต่ไม่สม่ำเสมอ", youdSay: "ฉันรู้ว่าการทบทวนพฤติกรรมและแนวคิดของตนเองเป็นเรื่องสำคัญ แต่ยังไม่ได้เริ่มฝึกฝน", managerSees: "ปฏิบัติตาม 8 พฤติกรรมหลัก (8 Living Habits) เป็นครั้งคราว มีสัญญาณของความตระหนักรู้ แต่ยังไม่เห็นรูปแบบการนำไปใช้ในชีวิตจริง" },
    ],
    confusion: "การตระหนักรู้และเข้าใจค่านิยมของบริษัทนั้นอาจยังไม่เพียงพอ การพัฒนาตนเองอย่างแท้จริงคือการนำค่านิยมต่างๆมาประยุกต์ใช้ในชีวิตประจำวัน โดยสะท้อนจากวิธีที่คุณปฏิบัติต่อตนเองและผู้อื่น",
    selfPrompts: [
      "คุณมองตนเองว่ามีการพัฒนาด้านการทบทวนตนเองอยู่ในระดับใด และคุณเห็นความเปลี่ยนแปลงในตนเองอย่างไรบ้าง? จงยกตัวอย่าง ทั้งนี้มีปัจจัยใดที่จะทำให้คุณพัฒนาตนเองในด้านดังกล่าวให้ดียิ่งขึ้นต่อไป?",
    ],
    managerPrompts: [
      "คุณมองเห็นความเปลี่ยนแปลงอะไรด้านการทบทวนตนเองของพนักงาน? จงยกตัวอย่าง",
      "คุณมองว่าพนักงานจะสามารถต่อยอดด้านดังกล่าวให้ดียิ่งขึ้นได้ในปีถัดไปได้อย่างไร? และในฐานะหัวหน้างาน คุณจะสามารถสนับสนุนพนักงานได้ในมุมใด?",
    ],
  },
  a2: {
    title: "ด้านบทบาทและความสัมพันธ์กับเพื่อนร่วมงาน (Role & Relationship – Team Level)",
    anchorMeaning: "การมีส่วนร่วมของพนักงานทุกคนนำไปสู่ความร่วมมือในการทำงานร่วมกันอย่างราบรื่น",
    intro: "มุมมองการสร้างความสัมพันธ์ในที่ทำงาน รวมถึงการทำงานกับเพื่อนร่วมงาน",
    stages: [
      { stage: "ก้าวหน้าอย่างมั่นคง (Mature)", looksLike: "ความสัมพันธ์กับเพื่อนร่วมงานมั่นคงมาก และสามารถเป็นแบบอย่างแก่ผู้อื่นได้", youdSay: "คนอื่นมักถามฉันว่าทำอย่างไรถึงจะสร้างความเชื่อมั่นระหว่างทีมตนเองกับทีมอื่นได้", managerSees: "มีปฏิสัมพันธ์กับเพื่อนร่วมงานในทางที่ดีมาก ทั้งภายในทีมและกับนอกทีม และแนวคิดจากเพื่อนร่วมงานทีมอื่นยังสอดคล้องกับแนวคิดของทีมตนเอง" },
      { stage: "พัฒนามั่นคง (Established)", looksLike: "ความสัมพันธ์กับเพื่อนร่วมงานค่อนข้างมั่นคง ต่างคนต่างปฏิบัติต่อกันอย่างถ้อยทีถ้อยอาศัย", youdSay: "ฉันมีความสัมพันธ์กับเพื่อนร่วมงานแทบทุกทีมที่เกี่ยวข้อง และให้ความสำคัญในการรักษาความสัมพันธ์เหล่านั้น", managerSees: "มีปฏิสัมพันธ์กับเพื่อนร่วมงานค่อนข้างสมดุล ทั้งภายในทีม และกับนอกทีม" },
      { stage: "กำลังพัฒนา (Developing)", looksLike: "มีความสัมพันธ์ที่ดีภายในทีม เริ่มมีความสัมพันธ์กับเพื่อนร่วมงานในทีมอื่น", youdSay: "ความสัมพันธ์กับเพื่อนร่วมงานส่วนใหญ่ของฉันยังอยู่ในทีมตนเอง แต่ฉันกำลังพยายามสร้างความสัมพันธ์กับทีมอื่นด้วย", managerSees: "มีปฏิสัมพันธ์กับเพื่อนร่วมงานภายในทีมค่อนข้างดี และเริ่มมีปฏิสัมพันธ์ที่ดีเพิ่มขึ้นกับเพื่อนร่วมงานนอกทีม" },
      { stage: "อยู่ในขั้นเริ่มต้น (Emerging)", looksLike: "มุ่งเน้นเฉพาะบทบาทของตนเอง มีความสัมพันธ์ข้ามทีมจำกัด", youdSay: "ฉันโฟกัสกับผลงานของตนเอง และเพื่อนร่วมงานในทีมของฉันเท่านั้น", managerSees: "มีปฏิสัมพันธ์กับเพื่อนร่วมงานอื่นนอกเหนือจากทีมของตนเองค่อนข้างน้อย" },
    ],
    confusion: "การรู้จักผู้คนจำนวนมากไม่ได้หมายความว่าคุณมีความสัมพันธ์ที่มั่นคงเสมอไป การพัฒนาในด้านบทบาทและความสัมพันธ์กับเพื่อนร่วมงานนั้นลงลึกไปถึงความไว้วางใจซึ่งกันและกัน เพื่อนำไปสู่การทำงานร่วมกันอย่างเกื้อกูล",
    selfPrompts: [
      "คุณมองตนเองว่ามีการพัฒนาด้านบทบาทและความสัมพันธ์กับเพื่อนร่วมงานอยู่ในระดับใด และคุณเห็นความเปลี่ยนแปลงในตนเองอย่างไรบ้าง? จงยกตัวอย่าง ทั้งนี้มีปัจจัยใดที่จะทำให้คุณพัฒนาตนเองในด้านดังกล่าวให้ดียิ่งขึ้นต่อไป?",
    ],
    managerPrompts: [
      "คุณมองเห็นความเปลี่ยนแปลงอะไรด้านบทบาทและความสัมพันธ์กับเพื่อนร่วมงานของพนักงาน? จงยกตัวอย่าง",
      "คุณมองว่าพนักงานจะสามารถต่อยอดด้านดังกล่าวให้ดียิ่งขึ้นได้ในปีถัดไปได้อย่างไร? และในฐานะหัวหน้างาน คุณจะสามารถสนับสนุนพนักงานได้ในมุมใด?",
    ],
  },
  a3: {
    title: "ด้านความรับผิดชอบและการมีส่วนร่วมต่อองค์กร (Stewardship – Organizational Level)",
    anchorMeaning: "ภาวะความเป็นผู้นำเพื่อประโยชน์ส่วนรวม และความเป็นอยู่ที่ดีของทุกคนที่ไว้วางใจในตัวคุณ",
    intro: "ความรับผิดชอบต่องานที่ได้รับมอบหมายและผลลัพธ์ รวมถึงภาวะการเป็นผู้นำ ให้คำแนะนำ และสนับสนุนผู้อื่น ไม่ว่าคุณจะเป็นหัวหน้างานหรือไม่ก็ตาม",
    stages: [
      { stage: "ก้าวหน้าอย่างมั่นคง (Mature)", looksLike: "สร้างสภาพแวดล้อมที่เอื้อให้ผู้อื่นก้าวขึ้นเป็นผู้นำ และมีความสามารถในการบ่มเพาะผู้นำรุ่นใหม่", youdSay: "ฉันมองว่าบทบาทของฉันคือการสร้างผู้นำรุ่นใหม่ๆต่อไป ไม่ใช่แค่เฉพาะผู้นำในงานปัจจุบันเท่านั้น", managerSees: "มีแนวคิดความรับผิดชอบ มีการแสดงบทบาทผู้นำเชิงรุกชัดเจน และมีการพัฒนาพนักงานเพื่อนร่วมงานอื่นเพื่อให้แสดงศักยภาพของความเป็นผู้นำอย่างเป็นรูปธรรม" },
      { stage: "พัฒนามั่นคง (Established)", looksLike: "แสดงความเป็นผู้นำแม้ไม่มีตำแหน่งอย่างเป็นทางการ ผ่านแนวคิดริเริ่มต่างๆ และสามารถสร้างแรงบันดาลใจในการทำงานให้กับผู้อื่นได้", youdSay: "ฉันมีแนวคิดริเริ่มงานต่างๆอยู่เสมอ และแนวคิดของฉันสามารถต่อยอดไปไกลได้เกินกว่าบทบาทที่ตนเองได้รับมอบหมาย", managerSees: "มีแนวคิดความรับผิดชอบ มีการแสดงความคิดริเริ่มด้วยตนเองอย่างสม่ำเสมอ และเริ่มแสดงบทบาทผู้นำเชิงรุก" },
      { stage: "กำลังพัฒนา (Developing)", looksLike: "มีความคิดริเริ่มในงานต่างๆ รวมถึงให้การสนับสนุน ผู้อื่นเริ่มมองคุณเป็นแรงบันดาลใจในการทำงาน", youdSay: "ฉันมีแนวคิดริเริ่มงานต่างๆที่นอกเหนือจากบทบาทที่ได้รับมอบหมายของตนเอง แม้ว่าอาจจะทำได้มากหรือน้อยต่างกันในแต่ละช่วงเวลา", managerSees: "มีแนวคิดความรับผิดชอบ มีการแสดงความคิดริเริ่มด้วยตนเองบางอย่างแต่ยังไม่สม่ำเสมอ" },
      { stage: "อยู่ในขั้นเริ่มต้น (Emerging)", looksLike: "มุ่งมั่นทำงานของตนเองให้สำเร็จ ปฏิบัติตามคำสั่ง อาจยังไม่มีความคิดริเริ่มในงานใหม่ๆ", youdSay: "ฉันโฟกัสกับการทำงานของตนเองให้สำเร็จและมีคุณภาพ", managerSees: "มีแนวคิดความรับผิดชอบ แต่เน้นบทบาทผู้รับคำสั่งเป็นหลัก" },
    ],
    confusion: "ความรับผิดชอบไม่ได้ถูกจำกัดอยู่เพียงตำแหน่งหัวหน้างานเท่านั้น แต่สะท้อนจากคุณภาพของการเป็นผู้นำไม่ว่าพนักงานจะอยู่ในบทบาทใดก็ตาม การพัฒนาในด้านนี้มักเกิดขึ้นเมื่อหัวหน้างานเปิดโอกาสให้พนักงานแสดงความคิดเห็น หรือรับผิดชอบงานต่างๆที่ท้าทายมากยิ่งขึ้น",
    selfPrompts: [
      "คุณมองว่าตนเองมีการพัฒนาตนเองด้านความรับผิดชอบและการมีส่วนร่วมกับองค์กรอยู่ในระดับใด และคุณเห็นความเปลี่ยนแปลงในตนเองอย่างไรบ้าง? จงยกตัวอย่าง ทั้งนี้มีปัจจัยใดที่จะทำให้คุณพัฒนาตนเองในด้านดังกล่าวให้ดียิ่งขึ้นต่อไป?",
    ],
    managerPrompts: [
      "คุณมองเห็นความเปลี่ยนแปลงอะไรด้านความรับผิดชอบและการมีส่วนร่วมกับองค์กรของพนักงาน? จงยกตัวอย่าง",
      "คุณมองว่าพนักงานจะสามารถต่อยอดด้านดังกล่าวให้ดียิ่งขึ้นได้ในปีถัดไปได้อย่างไร? และในฐานะหัวหน้างาน คุณจะสามารถสนับสนุนพนักงานได้ในมุมใด?",
    ],
  },
  a4: {
    title: "ด้านจิตใต้สำนึกต่อส่วนรวม (Consciousness – Universe Level)",
    anchorMeaning: "จิตใต้สำนึกรับผิดชอบต่อตนเองและส่วนรวม",
    intro: "มุมมองการตัดสินใจและการกระทำที่คำนึงถึงผลกระทบที่อาจเกิดขึ้นต่อตนเอง ผู้คน งานที่ได้รับมอบหมาย รวมถึงองค์กรในภาพรวม",
    stages: [
      { stage: "ก้าวข้ามผ่านตัวฉัน (Beyond Me)", looksLike: "ตัวตนและเป้าหมายเป็นอันหนึ่งอันเดียวกัน มีความสามารถในการสร้างสรรค์สิ่งใหม่ๆ เพื่อประโยชน์ส่วนรวม", youdSay: "ฉันลงมือทำโดยไม่ต้องมานั่งชั่งใจว่าสิ่งที่เป็นประโยชน์ต่อตนเองหรือไม่", managerSees: "ยอมเสียสละผลประโยชน์ส่วนตน เพื่อรักษาประโยชน์ส่วนรวมเป็นหลัก" },
      { stage: "เป็นตัวฉัน (As Me)", looksLike: "ตัวตนและเป้าหมายสอดประสานกัน บทบาทที่ได้รับมอบหมายสะท้อนความเป็นตนเอง", youdSay: "บทบาทของฉันสะท้อนตัวตนของฉัน และฉันไม่รู้สึกว่าตนเองกับงานเป็นคนละส่วนกัน", managerSees: "เป้าหมายที่ตั้งไว้และบทบาทที่ได้รับเป็นหนึ่งเดียวกัน มีการคำนึงถึงทั้งความต้องการส่วนตัวและส่วนรวม" },
      { stage: "ผ่านตัวฉัน (Through Me)", looksLike: "มองเห็นเป้าหมายที่กว้างกว่าบทบาทของตนเอง และพิจารณาผลกระทบต่อส่วนรวมควบคู่กับผลประโยชน์ส่วนตัว", youdSay: "ฉันมองเห็นเป้าหมายที่ยิ่งใหญ่กว่าบทบาทที่ได้รับมอบหมาย บางครั้งฉันยอมวางความต้องการส่วนตนไว้ เพื่อให้บรรลุเป้าหมายส่วนรวม", managerSees: "เป้าหมายที่ตั้งไว้กว้างกว่าบทบาทที่ได้รับมอบหมาย มีการตัดสินใจบางครั้งที่แสดงให้เห็นว่ายอมวางความต้องการส่วนตัวไว้เพื่อเดินหน้าไปกับส่วนรวม" },
      { stage: "ด้วยตัวฉัน (By Me)", looksLike: "มองเห็นตนเองเป็นผู้สร้างการเปลี่ยนแปลง มีประสิทธิภาพ และมีความมุ่งมั่นในการลงมือทำสูง", youdSay: "ฉันเป็นผู้รับผิดชอบงานที่ได้รับมอบหมาย และลงมือทำให้งานเหล่านั้นสำเร็จและเกิดขึ้นจริง", managerSees: "บรรลุเป้าหมาย (OKRs) ได้ดี มีการอัปเดตความคืบหน้าของผลงาน และเป็นผู้ลงมือทำให้งานสำเร็จ" },
      { stage: "ต่อตัวฉัน (To Me)", looksLike: "มุ่งเน้นและให้ความสนใจบทบาทของตนเองเป็นหลัก มีความรับผิดชอบต่อสถานการณ์ต่างๆ ในระดับพื้นฐาน", youdSay: "ฉันแสดงความรับผิดชอบต่องานที่ได้รับมอบหมาย และกำลังฝึกบทบาทการเป็นผู้ลงมือทำอยู่", managerSees: "มองสถานการณ์ต่างๆ ว่าเป็นปัจจัยภายนอก มีความตระหนักรู้ แต่ยังไม่ได้มองตนเองเป็นผู้สร้างความเปลี่ยนแปลง" },
    ],
    confusion: "การเปลี่ยนแปลงในด้านจิตใต้สำนึกของตนเองต่อส่วนรวมนั้นควรถูกพัฒนาอย่างค่อยเป็นค่อยไป ทั้งนี้ระดับทั้ง 5 ระดับข้างต้น ไม่ใช่เป็นการประเมินว่าบุคคลใดบุคคลหนึ่งเป็นพนักงานที่ดีหรือไม่ อย่างไร เพียงแต่เป็นส่วนหนึ่งในเส้นทางการพัฒนาตนเองของพนักงาน",
    selfPrompts: [
      "คุณมองเห็นตนเองด้านจิตใต้สำนึกของตนเองต่อส่วนรวมอยู่ในระดับใด หรือมีเหตุการณ์ใดในการทำงานที่ทำให้คุณเห็นมุมมองของตนเองที่เปลี่ยนไป? จงยกตัวอย่าง ทั้งนี้มีเรื่องใดที่คุณต้องการพัฒนาเพื่อให้เป็นประโยชน์ต่อส่วนรวมได้ดียิ่งขึ้นบ้าง?",
    ],
    managerPrompts: [
      "คุณมองเห็นความเปลี่ยนแปลงอะไรด้านจิตใต้สำนึกของตนเองต่อส่วนรวมของพนักงาน? จงยกตัวอย่าง",
      "คุณมองว่าพนักงานจะสามารถต่อยอดด้านดังกล่าวให้ดียิ่งขึ้นได้ในปีถัดไปได้อย่างไร? และในฐานะหัวหน้างาน คุณจะสามารถสนับสนุนพนักงานได้ในมุมใด?",
    ],
  },
  "b1-learning": {
    code: "ด้านการเรียนรู้และพัฒนาตนเอง (Learning & Self-cultivation)",
    pillar: "ตัวบุคคล",
    anchorEn: "ความสามารถในการเรียนรู้ ปรับตัวจากความล้มเหลว และมีส่วนร่วมในการสร้างวัฒนธรรมการเรียนรู้ภายในองค์กร",
    selfLabel: "โดยพนักงาน",
    selfPrompts: [
      "จงยกตัวอย่างเหตุการณ์ที่สะท้อนให้เห็นถึงการเรียนรู้ การปรับตัว หรือการเติบโตในบทบาทที่ได้รับมอบหมายในช่วงที่ผ่านมา โดยระบุ",
    ],
    managerPrompts: [
      "คุณเห็นการพัฒนาการที่สะท้อนถึงการตระหนักรู้ เรียนรู้ และการปรับตัวของพนักงาน หรือไม่ อย่างไร?",
      "คุณมองว่าเหตุการณ์ดังกล่าว ส่งผลต่อวิธีการทำงานของพนักงานในทางที่ดีขึ้นอย่างไร?",
    ],
  },
  "b1-integrity": {
    code: "ด้านการบริหารงานด้วยความซื่อตรง (Manage with Integrity)",
    pillar: "ตัวบุคคล (สำหรับหัวหน้างาน)",
    anchorEn: "ความซื่อตรงในการบริหารงานมักเริ่มจากความซื่อสัตย์และจริงใจ ต่อภารกิจขององค์กร ต่อเพื่อนร่วมงาน และต่อตัวของคุณเอง",
    selfLabel: "โดยพนักงาน",
    selfPrompts: [
      "จงยกตัวอย่างเหตุการณ์ที่สะท้อนให้เห็นว่าคุณบริหารจัดการทีม และตัดสินใจภายใต้ความซื่อตรงและจริงใจ โดยระบุ",
    ],
    managerPrompts: [
      "คุณเห็นการพัฒนาการที่สะท้อนถึงความซื่อตรงและยึดมั่นในความถูกต้องของพนักงาน หรือไม่ อย่างไร?",
      "คุณมองว่าเหตุการณ์ดังกล่าว ส่งผลต่อวิธีการทำงานของพนักงานในทางที่ดีขึ้นอย่างไร?",
    ],
  },
  "b1-coaching": {
    code: "ด้านการบริหารงานด้วยการให้คำแนะนำ (Manage with Coaching)",
    pillar: "ตัวบุคคล (สำหรับหัวหน้างาน)",
    anchorEn: "การพัฒนาผู้อื่นผ่านการตั้งคำถาม การให้คำแนะนำ เป็นที่ปรึกษา และการเปิดโอกาสให้ผู้อื่นเติบโต",
    selfLabel: "โดยพนักงาน",
    selfPrompts: [
      "จงยกตัวอย่างเหตุการณ์ที่สะท้อนให้เห็นว่าคุณให้ความสำคัญกับการพัฒนาคน ผ่านการให้คำปรึกษาหรือคำแนะนำ โดยระบุ",
    ],
    managerPrompts: [
      "คุณเห็นการพัฒนาการที่สะท้อนถึงความสามารถในการเติบโตในบทบาทหัวหน้างานของพนักงาน หรือไม่ อย่างไร?",
      "คุณมองว่าเหตุการณ์ดังกล่าว ส่งผลต่อวิธีการทำงานของพนักงานในทางที่ดีขึ้นอย่างไร?",
    ],
  },
  "b1-connection": {
    code: "ด้านการบริหารด้วยความผูกพัน (Manage with Connection)",
    pillar: "ตัวบุคคล (สำหรับหัวหน้างาน)",
    anchorEn: "การสร้างความสัมพันธ์ที่จริงใจ และดูแลความเป็นอยู่ที่ดี รวมถึงการส่งเสริมวัฒนธรรมการทำงานร่วมกันภายในทีม",
    selfLabel: "โดยพนักงาน",
    selfPrompts: [
      "จงยกตัวอย่างเหตุการณ์ที่สะท้อนให้เห็นว่าคุณเสริมสร้างความผูกพัน และความเป็นอันหนึ่งอันเดียวกันภายในทีม โดยระบุ",
    ],
    managerPrompts: [
      "คุณเห็นการพัฒนาการที่สะท้อนถึงการสร้างความกลมเกลียวภายในทีมของพนักงาน หรือไม่ อย่างไร?",
      "คุณมองว่าเหตุการณ์ดังกล่าว ส่งผลต่อวิธีการทำงานของพนักงานในทางที่ดีขึ้นอย่างไร?",
    ],
  },
  "b2-collab": {
    code: "ด้านความร่วมมือในการทำงาน (Collaboration & Partnership)",
    pillar: "แนวทางการทำงานและการปรากฏตัว",
    anchorEn: "การสนับสนุนความสำเร็จของผู้อื่น การสร้างความสัมพันธ์อันดี และการมีส่วนร่วมนอกเหนือจากเพื่อนร่วมงานของตนเอง",
    selfLabel: "โดยพนักงาน",
    selfPrompts: [
      "จงยกตัวอย่างเหตุการณ์ที่สะท้อนให้เห็นว่าคุณสนับสนุนการทำงานของเพื่อนร่วมงาน โดยเฉพาะอย่างยิ่งกับเพื่อนร่วมงานทีมอื่น จนสำเร็จลุล่วงไปได้ด้วยดี โดยระบุ",
    ],
    managerPrompts: [
      "คุณเห็นการพัฒนาการที่สะท้อนถึงการร่วมมือร่วมใจระหว่างพนักงานกับเพื่อนร่วมงานอื่น หรือไม่ อย่างไร?",
      "คุณมองว่าเหตุการณ์ดังกล่าว ส่งผลต่อวิธีการทำงานของพนักงานในทางที่ดีขึ้นอย่างไร?",
    ],
  },
  "b2-steward": {
    code: "ด้านภาวะความเป็นผู้นำและรับผิดชอบต่อผลลัพธ์ (Stewardship & Accountability)",
    pillar: "แนวทางการทำงานและการปรากฏตัว",
    anchorEn: "ค่านิยม การแสดงความรับผิดชอบ รวมถึงการตั้งคำถามอย่างสร้างสรรค์เพื่อให้ได้ผลลัพธ์ที่ดียิ่งขึ้น",
    selfLabel: "โดยพนักงาน",
    selfPrompts: [
      "จงยกตัวอย่างเหตุการณ์ที่สะท้อนให้เห็นว่าคุณมีความรับผิดชอบ แสดงความเห็น และลงมือปฏิบัติงานที่ได้รับมอบหมายได้อย่างดีเยี่ยม โดยระบุ",
    ],
    managerPrompts: ["พนักงานมีส่วนในการแสดงความคิดเห็น การเสนอความคิดสร้างสรรค์ในการทำงานหรือไม่ อย่างไร?"],
  },
  "b2-innovation": {
    code: "ด้านนวัตกรรม และความเป็นผู้ประกอบการ (Innovation & Entrepreneurship)",
    pillar: "แนวทางการทำงานและการปรากฏตัว",
    anchorEn: "สร้างสรรค์สิ่งใหม่ ปรับตัวต่อการเปลี่ยนแปลง และการเปลี่ยนความคิดเชิงรูปธรรมให้เป็นผลลัพธ์ที่สามารถปฏิบัติได้จริง",
    selfLabel: "โดยพนักงาน",
    selfPrompts: [
      "จงยกตัวอย่างเหตุการณ์ที่สะท้อนให้เห็นว่าคุณมีแนวคิดในการสร้างสรรค์สิ่งใหม่เพื่อประโยชน์ต่อส่วนรวม โดยระบุ",
    ],
    managerPrompts: ["พนักงานมีการสร้างสรรค์สิ่งใหม่นอกเหนือจากงานที่ได้รับมอบหมายหรือไม่ อย่างไร?"],
  },
  b3: {
    code: "การส่งมอบผลงานให้บรรลุเป้าหมายและมีความหมายต่อส่วนรวม (Delivering Results & Impact)",
    pillar: "ผลการปฏิบัติงาน",
    anchor: "คุณค่า และผลกระทบที่มีความหมายจากผลงานของพนักงานภายใต้บทบาทที่ได้รับมอบหมาย",
    selfLabel: "โดยพนักงาน",
    selfPrompts: [
      "จงยกตัวอย่างเหตุการณ์ที่สะท้อนให้เห็นถึงคุณค่าและผลกระทบที่มีความหมายจากผลงานที่คุณนำเสนอ หรือทำสำเร็จ โดยระบุสถานการณ์ ความคิดเห็น สิ่งที่คุณลงมือปฏิบัติ และสิ่งที่เกิดขึ้น",
    ],
    managerPrompts: ["จากผลงานของพนักงาน คุณมองว่าผลงานดังกล่าวเป็นประโยชน์ต่อส่วนรวมหรือไม่ อย่างไร?"],
  },
  b4: {
    selfOther: "ความคิดเห็นเพิ่มเติมต่อตนเอง",
    managerOther: "ข้อเสนอแนะ ความคิดเห็นเพิ่มเติมต่อพนักงาน",
    selfSnapshotLabel: "จากการประเมินตนเอง คุณมองว่าคุณอยู่ในระดับ",
    managerSnapshotLabel: "ในฐานะหัวหน้างาน คุณมองว่าพนักงานอยู่ในระดับ",
    waitingSnapshot: "รอการประเมินขั้นสุดท้ายจากหัวหน้างาน",
  },
};

const A5_TH = {
  title: "แผนพัฒนาตนเอง",
  ownedByYou: "พนักงานเป็นผู้จัดทำและรับผิดชอบแผนเท่านั้น หัวหน้างานสามารถแสดงความคิดเห็นได้หลังจากพนักงานจัดทำแผนแล้วเสร็จเท่านั้น",
  mobilityLabel: "การย้ายสถานที่ปฏิบัติงาน",
  mobilityDetailLabel: "หากพิจารณาได้ตามเงื่อนไข โปรดระบุรายละเอียด",
  dimensionLabel: "ด้านที่คุณอยากพัฒนาตนเองเป็นพิเศษ",
  dimensionHint: "กรุณาเลือก 1 ด้านที่ตรงกับความต้องการมากที่สุด",
  edgeLabel: "เหตุผลในการเลือกพัฒนาตนเองในหัวข้อด้านบน",
  edgeHint: "โปรดระบุเหตุผลในการเลือกพัฒนาตนเองในหัวข้อด้านบน และสิ่งที่ต้องการพัฒนาโดยสังเขป",
  expLabel: "แนวทางการพัฒนาตนเอง : ผ่านประสบการณ์จริง",
  expHint: "โปรดระบุทักษะหรือสิ่งที่ต้องการเรียนรู้จากประสบการณ์อย่างน้อย 1 อย่าง (สูงสุด 3 อย่าง) รวมถึงช่วงเวลาที่คาดว่าจะดำเนินการ",
  learnLabel: "แนวทางการพัฒนาตนเอง : ผ่านการเรียนรู้",
  learnHint: "โปรดระบุหลักสูตร โครงการ หรือหัวข้อการสัมมนาที่คุณสนใจอย่างน้อย 1 อย่าง (สูงสุด 3 อย่าง) เป้าหมายที่คาดว่าจะได้รับ รวมถึงช่วงเวลาที่คาดว่าจะดำเนินการ",
  supportLabel: "การสนับสนุนที่ต้องการจากหัวหน้างาน",
  supportHint: "หัวหน้างานสามารถช่วยสนับสนุนเพื่อให้คุณบรรลุเป้าหมายได้อย่างไรบ้าง? โปรดระบุ",
  managerLabel: "มุมมองของหัวหน้างาน",
  managerHint: "คุณมีมุมมองอย่างไรกับแผนพัฒนาตนเองของพนักงาน และมองว่าพนักงานจะสามารถต่อยอดทักษะดังกล่าวในอนาคตได้อย่างไรบ้าง? ในฐานะหัวหน้างานจะสามารถสนับสนุนพนักงานอย่างไรบ้างเพื่อให้บรรลุเป้าหมาย?",
};

const MOBILITY_TH = {
  "Yes, able to relocate anytime": "พร้อมเดินทางไปปฏิบัติที่สถานที่อื่นได้ทุกเมื่อ",
  "Yes, with some considerations": "อาจพิจารณาเดินทางไปปฏิบัติที่สถานที่อื่น ขึ้นอยู่กับเงื่อนไข",
  "No": "ยังไม่พร้อมในขณะนี้",
};
const DIMENSION_TH = {
  "Self-Cultivation – Individual Level": "ด้านการทบทวนตนเอง (Self-cultivation – Individual Level)",
  "Role & Relationship – Team Level": "ด้านบทบาทและความสัมพันธ์กับเพื่อนร่วมงาน (Role & Relationship – Team Level)",
  "Stewardship – Organizational Level": "ด้านความรับผิดชอบและการมีส่วนร่วมต่อองค์กร (Stewardship – Organizational Level)",
  "Consciousness – Universe Level": "ด้านจิตใต้สำนึกต่อส่วนรวม (Consciousness – Universe Level)",
};
const STAGE_TH = {
  Mature: "ก้าวหน้าอย่างมั่นคง (Mature)", Established: "พัฒนามั่นคง (Established)", Developing: "กำลังพัฒนา (Developing)", Emerging: "อยู่ในขั้นเริ่มต้น (Emerging)",
  "Beyond Me": "ก้าวข้ามผ่านตัวฉัน (Beyond Me)", "As Me": "เป็นตัวฉัน (As Me)", "Through Me": "ผ่านตัวฉัน (Through Me)", "By Me": "ด้วยตัวฉัน (By Me)", "To Me": "ต่อตัวฉัน (To Me)",
};
// Visual identity for each growth stage, by position (most-advanced-first,
// matching how every stage array in this app is already ordered). Deliberately
// not a red-to-green gradient — these are different places on a path, not a
// scored scale, so the colors are just distinct brand tints, not a hierarchy.
// Icons vary by stage to keep them visually distinguishable, but color is
// deliberately the SAME for every stage — this is self-reflection, not a
// rating, and even a "brighter for advanced, duller for early" gradient
// still reads as a hierarchy (gray in particular reads as "lesser," which
// is exactly wrong for something like "To Me" or "Emerging" — those are
// legitimate starting points, not deficient ones). Uniform color removes
// any color-based ranking; only the selection state itself is highlighted.
const GROWTH_STAGE_ICONS = [Mountain, TreeDeciduous, Leaf, Sprout, Circle];
const GROWTH_STAGE_COLOR = BRAND.primary;

const SNAPSHOT_TH = {
  Flourishing: { label: "สร้างคุณค่าและเปล่งประกาย (Flourishing)", desc: "สร้างสภาพแวดล้อมที่เอื้อให้ตนเองและผู้อื่นเติบโตและพัฒนาไปด้วยกันอย่างงอกงาม" },
  Generative: { label: "สร้างสรรค์และต่อยอดสิ่งดี (Generative)", desc: "สร้างผลงานอันเป็นประโยชน์ และสามารถนำไปต่อยอดได้อย่างแท้จริง" },
  Sustaining: { label: "รักษามาตรฐานตามความคาดหวัง (Sustaining)", desc: "รักษาการทำงานให้เป็นไปอย่างราบรื่น และส่งมอบงานได้ตามที่คาดหวัง" },
  Depleting: { label: "ส่งผลให้การทำงานถดถอยลง (Depleting)", desc: "ส่งผลให้ประสิทธิภาพ ความไว้วางใจ และความสามารถในการทำงานลดลง" },
};

const COVER_TH = {
  title: "แบบประเมินการเติบโต และการสร้างผลงานที่มีความหมาย (Growth & Value-Add Journey)",
  anchorMeaning: "บทสนทนาต่อเนื่องเพื่อดูแลผลงาน ความสามารถ และการเติบโตเป็นตัวเองที่ดีขึ้น",
  intro: "ไม่มี \"เวลาที่ถูกต้อง\" ตายตัวสำหรับการทบทวน มีเพียงการทบทวนตามช่วงเวลาจริงเท่านั้น ขอแนะนำให้คุณปรับปรุงเอกสารนี้อยู่เสมอ ตามการทำงาน การเรียนรู้ และผลกระทบที่เปลี่ยนแปลงไป ในบางช่วงเวลา ข้อมูลจากการทบทวนอาจถูกนำมาสรุปรวมเพื่อสนับสนุนการตัดสินใจด้านการดูแลบุคลากร ไม่ควรมีเรื่องน่าประหลาดใจเกิดขึ้น เพราะบทสนทนานี้เกิดขึ้นอย่างต่อเนื่องตลอดเวลา",
  employeeName: "ชื่อพนักงาน",
  employeeId: "รหัสพนักงาน",
  department: "แผนก",
  jobGrade: "ระดับตำแหน่ง",
  designation: "ตำแหน่งงาน",
  directManager: "หัวหน้างาน",
  isPeopleManager: "คุณมีบทบาทเป็นหัวหน้างานหรือไม่?",
  yes: "ใช่ — ฉันมีทีมที่ดูแลอยู่",
  no: "ไม่",
  reflectionDate: "การทบทวนครั้งนี้ (ณ วันที่)",
  fromRoster: "หากพบข้อมูลพนักงานผิดพลาด กรุณาแจ้ง P&O Admin เพื่อทำการแก้ไข",
};

const UI_TH = {
  self: "ตนเอง",
  manager: "หัวหน้างาน",
  selfReflection: "โดยพนักงาน",
  managerReflection: "มุมมองของหัวหน้างาน",
  managerFeedback: "ความเห็นจากหัวหน้างาน",
  back: "ย้อนกลับ",
  next: "ถัดไป",
  finishSubmit: "เสร็จสิ้นและส่ง",
  finishReview: "เสร็จสิ้นการตรวจสอบ",
  submitted: "ส่งแล้ว",
  reviewSubmitted: "ส่งการตรวจสอบแล้ว",
  sendForConfirmation: "ส่งให้ยืนยัน",
  viewOnlyGrowthStages: "ระดับการเติบโต — สิ่งเหล่านี้ไม่ใช่คะแนนประเมิน แต่ละระดับคือจุดที่มีความหมายในเส้นทางที่ยาวไกล ลองมองหาตำแหน่งของตัวเองอย่างใจดีกับตัวเอง",
  viewGrowthStages: "ดูระดับการเติบโต",
  progression: "การก้าวหน้า",
  commonConfusion: "ความเข้าใจผิดที่พบบ่อย",
  selfStageReflection: "การทบทวนระดับของตนเอง — ลองดูว่าตัวเองอยู่ตรงไหน",
  dontWorryLength: "ไม่ต้องกังวลเรื่องความยาว เพียงไม่กี่ประโยคที่มีรายละเอียดจริงมีความหมายมากกว่าข้อความที่เรียบเรียงสวยงามแต่ไม่เจาะจง",
  waitingForManager: "รอความคิดเห็นจากหัวหน้างาน",
  peopleManagersOnly: "สำหรับหัวหน้างานที่บริหารคนเท่านั้น",
  trackA: "ส่วนที่ A — การเติบโตและการพัฒนาตนเอง",
  trackB: "ส่วนที่ B — การดูแลรับผิดชอบและการสร้างมูลค่าเพิ่ม",
  myJourney: "แบบประเมินของฉัน",
  myTeam: "ทีมของฉัน",
  delegated: "ที่ได้รับมอบหมาย",
  status: "สถานะ",
};
const STATUS_LABEL_TH = {
  not_started: "ยังไม่เริ่ม",
  in_progress: "กำลังดำเนินการ",
  pending_review: "โดยพนักงาน — รอหัวหน้างานตรวจสอบ",
  drafted_pending_confirmation: "โดยผู้ได้รับมอบหมาย — รอการตรวจสอบขั้นสุดท้าย",
  completed: "เสร็จสมบูรณ์",
};

// t(en, thKeyOrText, lang) — tiny helper: returns the Thai text when lang is
// "th" and a translation exists, otherwise falls back to the English text
// (so anything not yet translated degrades gracefully instead of breaking).
function t(en, th, lang) {
  return lang === "th" && th ? th : en;
}

// Bolds just the topic phrase inside a longer question, instead of
// underlining the whole sentence — helps people spot "which topic is this
// about" without every prompt reading as one long emphasized block.
// Falls back to the plain string untouched if the phrase isn't found.
function boldPhrase(text, phrase) {
  if (!text || !phrase) return text;
  const idx = text.indexOf(phrase);
  if (idx === -1) return text;
  return (
    <>
      {text.slice(0, idx)}
      <strong>{phrase}</strong>
      {text.slice(idx + phrase.length)}
    </>
  );
}
// Strips a trailing " (English parenthetical)" from a title/code string, to
// get just the bare topic name for bolding inside a prompt sentence.
function bareTopic(title) {
  return title ? title.replace(/\s*\([^)]*\)\s*$/, "").trim() : title;
}


/* ========================================================================
   REQUIRED FIELD MANIFEST (drives the progress bar)
========================================================================= */
const TRACK_B_KEYS = [
  "b1_learning_self", "b1_learning_manager",
  "b1_integrity_self", "b1_integrity_manager",
  "b1_coaching_self", "b1_coaching_manager",
  "b1_connection_self", "b1_connection_manager",
  "b2_collab_self", "b2_collab_manager",
  "b2_steward_self", "b2_steward_manager",
  "b2_innovation_self", "b2_innovation_manager",
  "b3_results_self", "b3_results_manager",
  "b4_self", "b4_manager", "b4_snapshot_self", "b4_snapshot_manager", "b4_commit_emp", "b4_commit_mgr",
];

function requiredFields(isPeopleManager, trackBOnly) {
  const base = [
    "reflection_date",
    "a1_self", "a1_manager",
    "a2_self", "a2_manager",
    "a3_self", "a3_manager",
    "a4_stage", "a4_self", "a4_manager",
    "a5_mobility", "a5_dimension", "a5_edge", "a5_exp1", "a5_learn1", "a5_support", "a5_manager",
    "b1_learning_self", "b1_learning_manager",
    "b2_collab_self", "b2_collab_manager",
    "b2_steward_self", "b2_steward_manager",
    "b2_innovation_self", "b2_innovation_manager",
    "b3_results_self", "b3_results_manager",
    "b4_self", "b4_manager", "b4_snapshot_self", "b4_snapshot_manager", "b4_commit_emp", "b4_commit_mgr",
  ];
  if (isPeopleManager) {
    base.push(
      "b1_integrity_self", "b1_integrity_manager",
      "b1_coaching_self", "b1_coaching_manager",
      "b1_connection_self", "b1_connection_manager"
    );
  }
  if (!trackBOnly) return base;
  return base.filter((k) => TRACK_B_KEYS.includes(k));
}

/* ========================================================================
   PER-PAGE REQUIRED KEYS — drives the "must complete before Next" gate.
   Admin/read-only browsing is never gated (nothing to fill in).
========================================================================= */
function pageRequiredFor(section, viewRole) {
  if (viewRole === "admin") return [];
  if (viewRole === "employee") {
    switch (section.id) {
      case "cover": return ["reflection_date"];
      case "a1": return ["a1_stage", "a1_self"];
      case "a2": return ["a2_stage", "a2_self"];
      case "a3": return ["a3_stage", "a3_self"];
      case "a4": return ["a4_stage", "a4_self"];
      case "a5": return ["a5_mobility", "a5_dimension", "a5_edge", "a5_exp1", "a5_learn1", "a5_support"];
      case "b3": return ["b3_results_self"];
      case "b4": return ["b4_self", "b4_snapshot_self", "b4_commit_emp"];
      default: return section.fields?.self ? [section.fields.self] : [];
    }
  }
  if (viewRole === "manager") {
    switch (section.id) {
      case "cover": return [];
      case "a5": return ["a5_manager"];
      case "b3": return ["b3_results_manager"];
      case "b4": return ["b4_manager", "b4_snapshot_manager", "b4_commit_mgr"];
      default: return section.fields?.manager ? [section.fields.manager] : [];
    }
  }
  return [];
}
function isPageComplete(section, data, viewRole) {
  const req = pageRequiredFor(section, viewRole);
  return req.every((k) => {
    const v = data[k];
    return v === true || (typeof v === "string" && v.trim().length > 0);
  });
}

function roleProgress(visibleSections, data, viewRole) {
  const keys = visibleSections.flatMap((s) => pageRequiredFor(s, viewRole));
  if (!keys.length) return 100;
  const filled = keys.filter((k) => {
    const v = data[k];
    return v === true || (typeof v === "string" && v.trim().length > 0);
  }).length;
  return Math.round((filled / keys.length) * 100);
}

function computeProgress(data, isPeopleManager, trackBOnly) {
  if (!data) return 0;
  const required = requiredFields(isPeopleManager, trackBOnly);
  const filled = required.filter((k) => {
    const v = data[k];
    return v === true || (typeof v === "string" && v.trim().length > 0);
  }).length;
  return Math.round((filled / required.length) * 100);
}

function defaultData() {
  return {
    reflection_date: "",
    employee_submitted: false,
    employee_submitted_at: "",
    manager_submitted: false,
    manager_submitted_at: "",
    manager_submitted_by: "",
    manager_draft_by: "",
    manager_draft_at: "",
    a5_mobility: "", a5_mobility_detail: "", a5_dimension: "", a5_edge: "",
    a5_exp1: "", a5_exp2: "", a5_exp3: "", a5_learn1: "", a5_learn2: "", a5_learn3: "",
    a5_support: "", a5_manager: "",
    b3_okr1: "", b3_okr2: "", b3_okr3: "", b3_okr4: "", b3_okr5: "",
    b3_progress1: "", b3_progress2: "", b3_progress3: "", b3_progress4: "", b3_progress5: "",
    b3_results_self: "", b3_results_manager: "",
    b4_snapshot_self: "", b4_snapshot_manager: "", b4_commit_emp: false, b4_commit_emp_date: "", b4_commit_mgr: false, b4_commit_mgr_date: "",
  };
}

function keyFor(employeeId, roster, cycle) {
  // Namespaced by BU and cycle — in production this would be a fully separate
  // database per BU deployment (and a real archive table per cycle); the
  // prefix here mirrors that boundary even though this prototype shares one
  // storage backend.
  const u = (roster || SEED_ROSTER).find((x) => x.id === employeeId);
  const buSlug = u?.bu ? SLUG_BY_BU[u.bu] : "unknown";
  return `${STORAGE_KEY}:${buSlug}:${employeeId}:${cycle || DEFAULT_CYCLE}`;
}
async function loadData(employeeId, roster, cycle) {
  try {
    const res = await window.storage.get(keyFor(employeeId, roster, cycle), true);
    return res && res.value ? JSON.parse(res.value) : null;
  } catch {
    return null;
  }
}
async function saveData(employeeId, data, roster, cycle) {
  try {
    await window.storage.set(keyFor(employeeId, roster, cycle), JSON.stringify(data), true);
  } catch (e) {
    console.error("save failed", e);
  }
}

// UAT clean-up: wipes every submitted journey answer (all people, all
// companies, all cycles) so the app can be handed to a fresh cohort with a
// clean slate. Deliberately does NOT touch the roster — real employee
// accounts, companies, and reporting lines are left exactly as they are.
async function clearAllJourneyData() {
  let deleted = 0;
  try {
    const res = await window.storage.list(STORAGE_KEY, true);
    const keys = res?.keys || [];
    for (const k of keys) {
      try {
        await window.storage.delete(k, true);
        deleted++;
      } catch (e) {
        console.error("delete failed for", k, e);
      }
    }
  } catch (e) {
    console.error("clearAllJourneyData failed", e);
  }
  return deleted;
}
async function clearActivityLog() {
  try {
    await window.storage.set(LOG_KEY, JSON.stringify([]), true);
  } catch (e) {
    console.error("clearActivityLog failed", e);
  }
}

/* ========================================================================
   SMALL PRIMITIVES
========================================================================= */

function InfoTip({ text }) {
  const [open, setOpen] = useState(false);
  return (
    <span className="relative inline-block align-middle">
      <button
        type="button"
        onClick={(e) => { e.preventDefault(); e.stopPropagation(); setOpen((o) => !o); }}
        className="inline-flex items-center justify-center w-4 h-4 rounded-full"
        style={open ? { backgroundColor: BRAND.primary, color: "white" } : { color: "#94A3B8" }}
      >
        <Info className="w-3.5 h-3.5" />
      </button>
      {open && (
        <>
          <div className="fixed inset-0 z-10" onClick={() => setOpen(false)} />
          <div
            className="absolute z-20 left-0 top-6 w-64 rounded-xl p-3 text-xs leading-relaxed shadow-lg"
            style={{ backgroundColor: BRAND.deep, color: "white" }}
          >
            {text}
          </div>
        </>
      )}
    </span>
  );
}

function Field({ label, children, required }) {
  return (
    <div className="mb-4">
      {label && (
        <label className="text-sm font-medium mb-1.5 block" style={{ color: BRAND.deep }}>
          {label}
          {required && (
            <span className="ml-1" style={{ color: BRAND.red }}>*</span>
          )}
        </label>
      )}
      {children}
    </div>
  );
}

function TextInput({ value, onChange, readOnly, placeholder, type = "text" }) {
  return (
    <input
      type={type}
      value={value || ""}
      onChange={(e) => onChange && onChange(e.target.value)}
      readOnly={readOnly}
      placeholder={placeholder}
      className={`w-full rounded-xl border px-3.5 py-2.5 text-sm focus:outline-none focus:ring-2 transition ${
        readOnly ? "bg-slate-50 text-slate-500" : "bg-white"
      }`}
      style={{ borderColor: BRAND.line, "--tw-ring-color": BRAND.teal }}
    />
  );
}

function TextArea({ value, onChange, readOnly, placeholder, rows = 4 }) {
  return (
    <textarea
      value={value || ""}
      onChange={(e) => onChange && onChange(e.target.value)}
      readOnly={readOnly}
      placeholder={placeholder}
      rows={rows}
      className={`w-full rounded-xl border px-3.5 py-2.5 text-sm leading-relaxed focus:outline-none focus:ring-2 transition ${
        readOnly ? "bg-slate-50 text-slate-500" : "bg-white"
      }`}
      style={{ borderColor: BRAND.line, "--tw-ring-color": BRAND.teal }}
    />
  );
}

function PromptList({ prompts }) {
  if (!prompts?.length) return null;
  return (
    <div className="text-sm text-slate-700 mb-2.5 space-y-1.5 leading-relaxed">
      {prompts.map((p, i) => (
        <p key={i}>{p}</p>
      ))}
    </div>
  );
}

function RadioGroup({ options, value, onChange, readOnly, renderOption }) {
  return (
    <div className="space-y-2">
      {options.map((opt) => {
        const key = typeof opt === "string" ? opt : opt.key;
        const selected = value === key;
        return (
          <button
            type="button"
            key={key}
            disabled={readOnly}
            onClick={() => onChange && onChange(key)}
            className={`w-full text-left rounded-xl border px-3.5 py-2.5 flex items-start gap-3 transition ${
              selected ? "ring-2" : ""
            } ${readOnly ? "cursor-default" : "cursor-pointer hover:border-slate-300"}`}
            style={{
              borderColor: selected ? BRAND.teal : BRAND.line,
              backgroundColor: selected ? "#EAF7F6" : "white",
              "--tw-ring-color": BRAND.teal,
            }}
          >
            <span
              className="w-4 h-4 rounded-full border-2 shrink-0 mt-0.5 flex items-center justify-center"
              style={{ borderColor: selected ? BRAND.teal : "#CBD5E1" }}
            >
              {selected && <span className="w-2 h-2 rounded-full" style={{ backgroundColor: BRAND.teal }} />}
            </span>
            <span className="text-sm">
              {renderOption ? renderOption(opt) : <span className="font-medium text-slate-800">{key}</span>}
            </span>
          </button>
        );
      })}
    </div>
  );
}

// Splits "ก้าวข้ามผ่านตัวฉัน (Beyond Me)" into two lines — the Thai term,
// then the English parenthetical below it — so this column can stay
// narrow and the other three (which hold the actual descriptive content)
// get the room instead.
function StageLabelTwoLine({ label }) {
  const m = label.match(/^(.*?)\s*(\([^)]*\))$/);
  if (!m) return <span>{label}</span>;
  return (
    <span>
      {m[1]}
      <br />
      {m[2]}
    </span>
  );
}

function StageTable({ stages, title }) {
  const { lang } = useRosterCtx();
  const [open, setOpen] = useState(false);
  const cols = lang === "th" ? STAGE_COLUMNS_TH : STAGE_COLUMNS;
  const label = title || t("View growth stages", "ดูระดับการเติบโต", lang);
  return (
    <div className="mb-4 rounded-xl border overflow-hidden" style={{ borderColor: BRAND.line }}>
      <button
        onClick={() => setOpen(!open)}
        className="w-full flex items-center justify-between px-4 py-2.5 text-sm font-medium"
        style={{ color: BRAND.primary, backgroundColor: "#F0F8F8" }}
      >
        {label}
        {open ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
      </button>
      {open && (
        <div className="overflow-x-auto">
          <table className="w-full text-xs">
            <thead>
              <tr style={{ backgroundColor: "#F6FBFA" }}>
                {cols.map((c, i) => (
                  <th key={c} className={`text-left px-3 py-2 font-semibold ${i === 0 ? "w-[120px]" : "whitespace-nowrap"}`} style={{ color: BRAND.deep }}>
                    {c}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {stages.map((s, i) => {
                const Icon = GROWTH_STAGE_ICONS[i] || Circle;
                const color = GROWTH_STAGE_COLOR;
                return (
                <tr key={s.stage} className="border-t" style={{ borderColor: BRAND.line }}>
                  <td className="px-3 py-2.5 align-top" style={{ borderLeft: `3px solid ${color}` }}>
                    <span className="inline-flex items-start gap-1.5 font-semibold" style={{ color }}>
                      <Icon className="w-3.5 h-3.5 mt-0.5 shrink-0" /> <StageLabelTwoLine label={s.stage} />
                    </span>
                  </td>
                  <td className="px-3 py-2.5 align-top text-slate-600 min-w-[220px]">{s.looksLike}</td>
                  <td className="px-3 py-2.5 align-top text-slate-600 min-w-[220px] italic">"{s.youdSay}"</td>
                  <td className="px-3 py-2.5 align-top text-slate-600 min-w-[220px]">{s.managerSees}</td>
                </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function NoteBox({ label, children, tone = "info" }) {
  const colors = tone === "warn" ? { bg: "#FFF4F3", text: "#8A1300", border: "#F6C8C3" } : { bg: "#F0F8F8", text: BRAND.deep, border: BRAND.line };
  const Icon = tone === "warn" ? AlertCircle : Clock;
  return (
    <div className="rounded-xl px-3.5 py-2.5 text-xs leading-relaxed mb-3 flex items-start gap-2" style={{ backgroundColor: colors.bg, color: colors.text, border: `1px solid ${colors.border}` }}>
      <Icon className="w-3.5 h-3.5 mt-0.5 shrink-0" />
      <span><span className="font-semibold">{label}: </span>{children}</span>
    </div>
  );
}

// The three-part breakdown the document asks for under every B1/B2
// self-reflection prompt: situation, action/decision, result/impact.
const STAR_PARTS_EN = ["Situation / challenge that occurred", "Action, decision, or your perspective on the situation", "Result / impact"];
const STAR_PARTS_TH = ["สถานการณ์ / ความท้าทายที่เกิดขึ้น", "การกระทำ การตัดสินใจ หรือมุมมองของคุณต่อเหตุการณ์ดังกล่าว", "ผลลัพธ์ / ผลกระทบ"];

const PILLAR_TH = {
  Person: "ตัวบุคคล",
  "Person (managers)": "ตัวบุคคล (สำหรับหัวหน้างาน)",
  "Practice & Presence": "แนวทางการทำงานและการปรากฏตัว",
  Performance: "ผลการปฏิบัติงาน",
  Overall: "ภาพรวม",
};

function SectionHeader({ code, title, chinese, anchorMeaning, pillar, track }) {
  const { lang } = useRosterCtx();
  return (
    <div className="mb-5">
      {pillar && (
        <p className="text-xs font-semibold uppercase tracking-wide mb-1 flex items-center gap-1.5" style={{ color: BRAND.mint }}>
          <Building2 className="w-3.5 h-3.5" /> {t("Track B – Stewardship & Value-Add", "ส่วนที่ B – ภาวะความเป็นผู้นำและการสร้างประโยชน์ (Stewardship & Value-Add)", lang)} · {t(pillar, PILLAR_TH[pillar], lang)}
        </p>
      )}
      {track === "A" && (
        <p className="text-xs font-semibold uppercase tracking-wide mb-1 flex items-center gap-1.5" style={{ color: BRAND.teal }}>
          <Sprout className="w-3.5 h-3.5" /> {t("Track A – Growth & Becoming", "ส่วนที่ A – การเติบโตและการพัฒนาตนเอง (Growth & Becoming)", lang)}
        </p>
      )}
      {code && !title && (
        <h2 className="text-lg font-semibold mb-1" style={{ color: BRAND.deep, fontFamily: FONT_STACK }}>
          {code}
        </h2>
      )}
      {title && (
        <h2 className="text-xl font-semibold mb-1" style={{ color: BRAND.deep, fontFamily: FONT_STACK }}>
          {title}
        </h2>
      )}
      {chinese && anchorMeaning ? (
        <p className="text-sm mb-1">
          <span style={{ fontFamily: "'Noto Sans SC', sans-serif", color: BRAND.teal, fontWeight: 600 }}>{chinese}</span>
          <span className="text-slate-500"> : {anchorMeaning}</span>
        </p>
      ) : (
        <>
          {chinese && (
            <p className="text-sm mb-1" style={{ fontFamily: "'Noto Sans SC', sans-serif", color: BRAND.teal }}>
              {chinese}
            </p>
          )}
          {anchorMeaning && <p className="text-sm text-slate-500 leading-relaxed">{anchorMeaning}</p>}
        </>
      )}
    </div>
  );
}

/* ========================================================================
   PAGE RENDERERS
========================================================================= */

function InfoRow({ label, value, icon: Icon }) {
  return (
    <div>
      <p className="text-xs text-slate-400 mb-0.5 flex items-center gap-1">
        {Icon && <Icon className="w-3 h-3" style={{ color: BRAND.teal }} />} {label}
      </p>
      <p className="text-sm font-medium text-slate-800">{value || "—"}</p>
    </div>
  );
}

function CoverPage({ data, setField, role, person, manager, isPeopleManager, managerReadyCount, managerTotalCount }) {
  const { cycle, cycleStart, cycleEnd, lang } = useRosterCtx();
  const empRO = role !== "employee" || !!data.employee_submitted;
  const status = journeyStatus(data);
  const sMeta = JOURNEY_STATUS_META[status];
  const SIcon = sMeta.icon;
  return (
    <div>
      <SectionHeader
        title={
          lang === "th" ? (
            <>แบบประเมินการเติบโต และการสร้างผลงานที่มีความหมาย<br />(Growth &amp; Value-Add Journey)</>
          ) : (
            "Growth & Value-Add Journey"
          )
        }
      />
      <div className="rounded-xl px-4 py-3 mb-4 flex items-start gap-2.5" style={{ backgroundColor: "#F0F8F8", border: `1px solid ${BRAND.line}` }}>
        <Clock className="w-4 h-4 mt-0.5 shrink-0" style={{ color: BRAND.primary }} />
        <p className="text-sm" style={{ color: BRAND.deep }}>
          {t(
            <>This journey is for the <span className="font-semibold">{cycle} review cycle</span>.</>,
            <>แบบประเมินนี้สำหรับ<span className="font-semibold">รอบการประเมินปี {cycle}</span></>,
            lang
          )}
          {cycleEnd && (
            <span className="block mt-1 font-medium">
              {t(
                `Please complete this by ${new Date(cycleEnd).toLocaleDateString()}${cycleStart ? ` (window opened ${new Date(cycleStart).toLocaleDateString()})` : ""}.`,
                `กรุณาทำให้เสร็จภายในวันที่ ${new Date(cycleEnd).toLocaleDateString("th-TH")}${cycleStart ? ` (เริ่มเมื่อ ${new Date(cycleStart).toLocaleDateString("th-TH")})` : ""}`,
                lang
              )}
            </span>
          )}
        </p>
      </div>
      <p className="text-sm text-slate-500 leading-relaxed mb-6">
        {t(
          "This evaluation is designed for employees to review and reflect on their own perspective of their recent work — including their own growth and development, working with others, day-to-day practice, results, and their contribution to the organisation as a whole. This reflection is meant to help employees see both what's going well and what they've learned, as well as opportunities to build further, supporting continuous understanding and development at the individual, department, and organisational level. Information and results from this evaluation will be analysed in aggregate to inform people-development decisions and related initiatives going forward.",
          "แบบประเมินนี้มีเป้าหมายเพื่อให้พนักงานได้ทบทวนและสะท้อนมุมมองของตนเองต่อการทำงานที่ผ่านมา ทั้งในด้านการเติบโตและการพัฒนาของตนเอง การทำงานร่วมกับผู้อื่น การปฏิบัติงาน ผลลัพธ์ รวมถึงการมีส่วนร่วมกับองค์กรในภาพรวม การสะท้อนความคิดเห็นในครั้งนี้มุ่งเน้นให้พนักงานได้มองเห็นทั้งสิ่งที่ทำได้ดี สิ่งที่ได้เรียนรู้ และโอกาสในการพัฒนาต่อยอด เพื่อให้เกิดความเข้าใจและการพัฒนาอย่างต่อเนื่องทั้งในระดับบุคคล แผนก และองค์กร ข้อมูลและผลการประเมินจะถูกนำไปวิเคราะห์ในภาพรวม เพื่อใช้เป็นข้อมูลประกอบการพัฒนาบุคลากรและแนวทางการพัฒนาในด้านต่าง ๆ ต่อไป",
          lang
        )}
      </p>

      <div className="rounded-xl px-4 py-3 mb-4 flex items-start gap-2.5" style={{ backgroundColor: `${sMeta.color}1A`, border: `1px solid ${sMeta.color}` }}>
        <SIcon className="w-4 h-4 mt-0.5 shrink-0" style={{ color: sMeta.color }} />
        <div>
          <p className="text-sm font-medium" style={{ color: sMeta.color }}>
            {t("Status", UI_TH.status, lang)}: {t(sMeta.label, STATUS_LABEL_TH[status], lang)}
          </p>
          <p className="text-xs text-slate-500 mt-0.5">
            {status === "not_started" && t("You haven't started your self-assessment yet.", "คุณยังไม่ได้เริ่มทำแบบประเมินตนเอง", lang)}
            {status === "in_progress" && t("Keep going — fill in each section, then submit for your manager's review.", "ทำต่อไป — กรอกข้อมูลในแต่ละส่วน แล้วส่งให้หัวหน้างานตรวจสอบ", lang)}
            {status === "pending_review" && (
              lang === "th" ? (
                <>ส่งแล้ว{data.employee_submitted_at ? ` เมื่อ ${new Date(data.employee_submitted_at).toLocaleDateString("th-TH")}` : ""} คำตอบของคุณถูกล็อกไว้ ความเห็นจากหัวหน้างานพร้อมแล้ว {managerReadyCount} จาก {managerTotalCount} ส่วน</>
              ) : (
                <>Submitted{data.employee_submitted_at ? ` on ${new Date(data.employee_submitted_at).toLocaleDateString()}` : ""}. Your answers are
                locked. Manager feedback ready: {managerReadyCount} of {managerTotalCount} sections.</>
              )
            )}
            {status === "completed" && (
              lang === "th" ? (
                <>หัวหน้างานตรวจสอบเสร็จสิ้นแล้ว{data.manager_submitted_at ? ` เมื่อ ${new Date(data.manager_submitted_at).toLocaleDateString("th-TH")}` : ""} ดูคำตอบของคุณและความเห็นจากหัวหน้างานได้ด้านล่าง</>
              ) : (
                <>Your manager completed their review{data.manager_submitted_at ? ` on ${new Date(data.manager_submitted_at).toLocaleDateString()}` : ""}. Browse your answers and their feedback below.</>
              )
            )}
          </p>
        </div>
      </div>

      <div
        className="rounded-xl border p-4 grid grid-cols-1 min-[480px]:grid-cols-2 gap-4 mb-2"
        style={{ borderColor: BRAND.line, backgroundColor: "#F6FBFA" }}
      >
        <InfoRow label={t("Employee name", COVER_TH.employeeName, lang)} value={person ? `${person.firstName} ${person.lastName}` : "—"} icon={User} />
        <InfoRow label={t("Employee ID", COVER_TH.employeeId, lang)} value={person?.employeeId} />
        <InfoRow label={t("Username (for login)", "ชื่อผู้ใช้ (สำหรับเข้าสู่ระบบ)", lang)} value={person?.username} />
        <InfoRow label={t("Company", "บริษัท", lang)} value={person?.bu ? `${person.bu} — ${BU_FULL_NAME[person.bu]}` : "—"} icon={Building2} />
        <InfoRow label={t("Department", COVER_TH.department, lang)} value={person?.department} />
        <InfoRow label={t("Job grade", COVER_TH.jobGrade, lang)} value={person?.jobGrade} />
        <InfoRow label={t("Designation", COVER_TH.designation, lang)} value={person?.designation} />
        <InfoRow label={t("Direct manager", COVER_TH.directManager, lang)} value={manager ? `${manager.firstName} ${manager.lastName}` : "—"} icon={UsersIcon} />
        <InfoRow label={t("Do you have a people manager role?", COVER_TH.isPeopleManager, lang)} value={isPeopleManager ? t("Yes — I have a team I manage", COVER_TH.yes, lang) : t("No", COVER_TH.no, lang)} />
      </div>
      <p className="text-xs text-slate-400 mb-6">
        {t(
          "If any employee details look wrong, please let your P&O Admin know so they can correct it.",
          COVER_TH.fromRoster,
          lang
        )}
      </p>

      <Field label={t("Current reflection (as-at date)", COVER_TH.reflectionDate, lang)} required>
        <TextInput type="date" value={data.reflection_date} onChange={(v) => setField("reflection_date", v)} readOnly={empRO} />
      </Field>
    </div>
  );
}

function GrowthStagePage({ section, data, setField, role }) {
  const { lang } = useRosterCtx();
  const th = SECTIONS_TH[section.id] || {};
  const { fields } = section;
  const empRO = role !== "employee" || !!data.employee_submitted;
  const mgrLocked = role === "manager" && !!data.manager_submitted;
  const waitingForEmployee = role === "manager" && !data.manager_submitted && !data.employee_submitted;
  const stages = lang === "th" && th.stages ? th.stages : section.stages;
  return (
    <div>
      <SectionHeader code={section.code} title={t(section.title, th.title, lang)} chinese={section.chinese} anchorMeaning={t(section.anchorMeaning, th.anchorMeaning, lang)} track="A" />
      {section.id === "a1" && (
        <p className="text-xs text-slate-400 mb-4">
          {t("Remark: Everything in Track A stays between you and your direct manager.", "หมายเหตุ: ข้อมูลทั้งหมดในส่วนนี้เป็นความลับเฉพาะพนักงานและหัวหน้างานเท่านั้น", lang)}
        </p>
      )}
      <p className="text-sm text-slate-600 leading-relaxed mb-2">{t(section.intro, th.intro, lang)}</p>

      <StageTable stages={stages} />

      {section.confusion && <NoteBox label={t("Note", "ข้อคิด", lang)}>{t(section.confusion, th.confusion, lang)}</NoteBox>}

      {(section.stageRadio || fields?.radio) && (
        <Field label={t("Which stage best describes you right now?", "คุณอยู่ในระดับใด", lang)} required>
          <p className="text-xs font-medium mb-2.5" style={{ color: BRAND.red }}>
            {t("Reminder: this is a growth stage, not a performance rating.", "ข้อควรทราบ: นี่คือระดับการเติบโต ไม่ใช่คะแนนการประเมิน", lang)}
          </p>
          <GrowthStageSelector stages={section.stages} value={data[fields.radio]} onChange={(v) => setField(fields.radio, v)} readOnly={empRO} lang={lang} />
        </Field>
      )}

      <Field label={t("By Employee", UI_TH.selfReflection, lang)} required>
        <p className="text-sm text-slate-700 mb-2">{boldPhrase((lang === "th" && th.selfPrompts ? th.selfPrompts : section.selfPrompts)[0], bareTopic(t(section.title, th.title, lang)))}</p>
        <TextArea value={data[fields.self]} onChange={(v) => setField(fields.self, v)} readOnly={empRO} placeholder={t("Write about a real moment — specifics carry further than a general theme.", "เขียนถึงเหตุการณ์จริง — รายละเอียดเฉพาะเจาะจงมีความหมายมากกว่าภาพรวมกว้าง ๆ", lang)} rows={5} />
      </Field>

      <ManagerFieldBlock
        label={t("Manager stewardship reflection", UI_TH.managerReflection, lang)}
        prompts={lang === "th" && th.managerPrompts ? th.managerPrompts : section.managerPrompts}
        value={data[fields.manager]}
        onChange={(v) => setField(fields.manager, v)}
        role={role}
        locked={mgrLocked}
        waitingForEmployee={waitingForEmployee}
        placeholder={t("What shift have you seen, and how can you support them?", "คุณเห็นการเปลี่ยนแปลงอะไร และจะสนับสนุนเขา/เธอได้อย่างไร?", lang)}
      />
    </div>
  );
}

// B1/B2/B3 self-prompts all follow "<lead-in> <the actual thing to reflect
// on> <โดยระบุ / specifying: ...>". Bold just that middle part — it's a
// different shape from Track A's prompts, so a separate helper from
// boldPhrase rather than forcing the same topic-matching logic to fit.
function boldStarLead(text, lang) {
  if (!text) return text;
  if (lang === "th") {
    const prefixes = ["จงยกตัวอย่างเหตุการณ์ที่สะท้อนให้เห็นถึง", "จงยกตัวอย่างเหตุการณ์ที่สะท้อนให้เห็นว่า"];
    const prefix = prefixes.find((p) => text.startsWith(p));
    if (!prefix) return text;
    const rest = text.slice(prefix.length);
    const idx = rest.indexOf("โดยระบุ");
    const middle = (idx === -1 ? rest : rest.slice(0, idx)).replace(/\s+$/, "");
    const suffix = idx === -1 ? "" : rest.slice(idx);
    return (
      <>
        {prefix}
        <strong>{middle}</strong>
        {suffix ? <> {suffix}</> : null}
      </>
    );
  }
  const prefix = "Give an example that reflects ";
  if (!text.startsWith(prefix)) return text;
  const rest = text.slice(prefix.length);
  const idx = rest.indexOf(", specifying:");
  const middle = idx === -1 ? rest : rest.slice(0, idx);
  const suffix = idx === -1 ? "" : rest.slice(idx);
  return (
    <>
      {prefix}
      <strong>{middle}</strong>
      {suffix}
    </>
  );
}

// Manager-side lead questions follow a handful of fixed sentence shapes
// across A1-A4/B1/B2 ("Have you seen development reflecting X?", "What
// change have you seen in X?", "Has the employee done X?"). Try each known
// prefix/suffix pair and bold whatever sits between them; fall back to
// plain text for the few prompts (B3, the generic follow-up questions)
// that don't have a clearly separable topic clause.
function boldManagerLead(text) {
  if (!text) return text;
  const patterns = [
    { prefix: "คุณมองเห็นความเปลี่ยนแปลงอะไรด้าน", suffix: "ของพนักงาน?" },
    { prefix: "คุณเห็นการพัฒนาการที่สะท้อนถึง", suffix: "หรือไม่ อย่างไร?" },
    { prefix: "พนักงานมี", suffix: "หรือไม่ อย่างไร?" },
    { prefix: "What shift have you seen emerging in the dimension of ", suffix: " in the Employee?" },
    { prefix: "What development have you seen reflecting ", suffix: "?" },
    { prefix: "Has the Employee ", suffix: "?" },
  ];
  for (const { prefix, suffix } of patterns) {
    if (!text.startsWith(prefix)) continue;
    const rest = text.slice(prefix.length);
    const idx = rest.indexOf(suffix);
    if (idx === -1) continue;
    const middle = rest.slice(0, idx).replace(/\s+$/, "");
    if (!middle) continue;
    const tail = rest.slice(idx);
    return (
      <>
        {prefix}
        <strong>{middle}</strong>
        {tail ? ` ${tail}` : ""}
      </>
    );
  }
  return text;
}

function ReflectionPage({ section, data, setField, role }) {
  const { lang } = useRosterCtx();
  const th = SECTIONS_TH[section.id] || {};
  const { fields } = section;
  const empRO = role !== "employee" || !!data.employee_submitted;
  const mgrLocked = role === "manager" && !!data.manager_submitted;
  const waitingForEmployee = role === "manager" && !data.manager_submitted && !data.employee_submitted;
  return (
    <div>
      {section.managerOnly && (
        <div className="mb-4 text-xs font-semibold px-3 py-1.5 rounded-full inline-flex items-center gap-1" style={{ backgroundColor: "#FFF4F3", color: BRAND.red, border: `1px solid #F6C8C3` }}>
          {t("For managers only — if this doesn't apply to you, please skip to the next section", "สำหรับหัวหน้างานเท่านั้น หากไม่ใช่ กรุณาข้ามไปหัวข้อถัดไป", lang)}
        </div>
      )}
      <SectionHeader code={t(section.code, th.code, lang)} pillar={section.pillar} anchorMeaning={section.anchorEn ? t(section.anchorEn, th.anchorEn, lang) : null} />
      {section.confusion && <NoteBox label={t("Note", "ข้อคิด", lang)}>{t(section.confusion, th.confusion, lang)}</NoteBox>}

      <Field label={t(section.selfLabel || "Self-Reflection", th.selfLabel, lang)} required={!section.managerOnly}>
        <p className="text-sm text-slate-700 mb-2">{boldStarLead((lang === "th" && th.selfPrompts ? th.selfPrompts : section.selfPrompts)[0], lang)}</p>
        <p className="text-xs text-slate-400 mb-1">{t("Please specify:", "โดยระบุ:", lang)}</p>
        <ul className="text-sm text-slate-500 mb-3 space-y-1 leading-relaxed list-disc pl-5">
          {(lang === "th" ? STAR_PARTS_TH : STAR_PARTS_EN).map((p, i) => <li key={i}>{p}</li>)}
        </ul>
        <TextArea value={data[fields.self]} onChange={(v) => setField(fields.self, v)} readOnly={empRO} rows={5} />
      </Field>

      <ManagerFieldBlock label={t("Manager Feedback", UI_TH.managerFeedback, lang)} prompts={lang === "th" && th.managerPrompts ? th.managerPrompts : section.managerPrompts} value={data[fields.manager]} onChange={(v) => setField(fields.manager, v)} role={role} locked={mgrLocked} waitingForEmployee={waitingForEmployee} />
    </div>
  );
}

function GrowthExperimentTable({ label, values, onChange, readOnly }) {
  return (
    <div className="rounded-xl border overflow-hidden mb-1" style={{ borderColor: BRAND.line }}>
      {values.map((v, i) => (
        <div key={i} className={`flex items-stretch ${i !== 0 ? "border-t" : ""}`} style={{ borderColor: BRAND.line }}>
          <div className="w-10 shrink-0 flex items-center justify-center text-sm font-semibold" style={{ backgroundColor: "#F6FBFA", color: BRAND.primary }}>
            {i + 1}
          </div>
          <input
            value={v || ""}
            onChange={(e) => onChange(i, e.target.value)}
            readOnly={readOnly}
            placeholder={i === 0 ? label : "Optional"}
            className="flex-1 px-3 py-2.5 text-sm focus:outline-none"
          />
        </div>
      ))}
    </div>
  );
}

function OKRTable({ objectives, progresses, onChangeObjective, onChangeProgress, readOnly, lang }) {
  return (
    <div className="rounded-xl border overflow-hidden mb-1" style={{ borderColor: BRAND.line }}>
      {objectives.map((v, i) => (
        <div key={i} className={`flex items-stretch ${i !== 0 ? "border-t" : ""}`} style={{ borderColor: BRAND.line }}>
          <div className="w-8 shrink-0 flex items-center justify-center text-sm font-semibold" style={{ backgroundColor: "#F6FBFA", color: BRAND.primary }}>
            {i + 1}
          </div>
          <div className="flex flex-col min-[560px]:flex-row flex-1">
            <input
              value={v || ""}
              onChange={(e) => onChangeObjective(i, e.target.value)}
              readOnly={readOnly}
              placeholder={t("Objective (OKR)", "เป้าหมาย (OKRs)", lang)}
              className="flex-1 px-3 py-2.5 text-sm focus:outline-none border-b min-[560px]:border-b-0 min-[560px]:border-r"
              style={{ borderColor: BRAND.line }}
            />
            <input
              value={progresses[i] || ""}
              onChange={(e) => onChangeProgress(i, e.target.value)}
              readOnly={readOnly}
              placeholder={t("Progress / Result", "ความคืบหน้า / ผลลัพธ์", lang)}
              className="flex-1 px-3 py-2.5 text-sm focus:outline-none"
            />
          </div>
        </div>
      ))}
    </div>
  );
}

function B3Page({ data, setField, role }) {
  const { lang } = useRosterCtx();
  const th = SECTIONS_TH.b3;
  const empRO = role !== "employee" || !!data.employee_submitted;
  const mgrLocked = role === "manager" && !!data.manager_submitted;
  const waitingForEmployee = role === "manager" && !data.manager_submitted && !data.employee_submitted;
  const setOkr = (i, v) => setField(`b3_okr${i + 1}`, v);
  const setProgress = (i, v) => setField(`b3_progress${i + 1}`, v);
  return (
    <div>
      <SectionHeader code={t("Performance", th.pillar, lang)} pillar="Performance" />
      <Field label={t("Objectives (OKRs)", "เป้าหมาย (OKRs)", lang)}>
        <p className="text-sm text-slate-700 mb-2">
          {t("If your role tracks formal OKRs, list up to 5 with their current progress or result. (Optional)", "หากบทบาทของคุณมีการติดตาม OKR อย่างเป็นทางการ ระบุได้สูงสุด 5 ข้อพร้อมความคืบหน้าหรือผลลัพธ์ปัจจุบัน (ไม่บังคับ)", lang)}
        </p>
        <OKRTable
          objectives={[data.b3_okr1, data.b3_okr2, data.b3_okr3, data.b3_okr4, data.b3_okr5]}
          progresses={[data.b3_progress1, data.b3_progress2, data.b3_progress3, data.b3_progress4, data.b3_progress5]}
          onChangeObjective={setOkr}
          onChangeProgress={setProgress}
          readOnly={empRO}
          lang={lang}
        />
      </Field>

      <SectionHeader title={t("Delivering Results & Impact", "การส่งมอบผลงานให้บรรลุเป้าหมายและมีความหมายต่อส่วนรวม (Delivering Results & Impact)", lang)} anchorMeaning={t("The value and meaningful impact of the Employee's work under their assigned role.", th.anchor, lang)} />

      <Field label={t(th.selfLabel, th.selfLabel, lang)} required>
        <p className="text-sm text-slate-700 mb-2">{boldStarLead(t("Give an example that reflects the value and meaningful impact from work you delivered or completed, specifying: the situation, your perspective, the action you took, and what happened.", th.selfPrompts[0], lang), lang)}</p>
        <TextArea value={data.b3_results_self} onChange={(v) => setField("b3_results_self", v)} readOnly={empRO} rows={5} />
      </Field>

      <ManagerFieldBlock
        label={t("Manager Feedback", UI_TH.managerFeedback, lang)}
        prompts={th.managerPrompts}
        value={data.b3_results_manager}
        onChange={(v) => setField("b3_results_manager", v)}
        role={role}
        locked={mgrLocked}
        waitingForEmployee={waitingForEmployee}
      />
    </div>
  );
}

function GrowthPlanPage({ data, setField, role }) {
  const { lang } = useRosterCtx();
  const setExp = (i, v) => setField(`a5_exp${i + 1}`, v);
  const setLearn = (i, v) => setField(`a5_learn${i + 1}`, v);
  const selfLocked = role === "employee" && !!data.employee_submitted;
  const empRO = role !== "employee" || selfLocked;
  const mgrLocked = role === "manager" && !!data.manager_submitted;
  const waitingForEmployee = role === "manager" && !data.manager_submitted && !data.employee_submitted;
  return (
    <div>
      <div className="mb-4 text-xs font-medium px-3 py-1.5 rounded-full inline-flex items-center gap-1.5" style={{ backgroundColor: "#FFF9E8", color: "#8A6D00", border: "1px solid #F3E3A8" }}>
        <User className="w-3.5 h-3.5" /> {t("Owned by you. Your manager can comment after you commit but cannot edit.", A5_TH.ownedByYou, lang)}
      </div>
      <SectionHeader title={t("A5 · Growth Plan", A5_TH.title, lang)} track="A" />

      <Field label={t("Mobility preference", A5_TH.mobilityLabel, lang)} required>
        <RadioGroup
          options={["Yes, able to relocate anytime", "Yes, with some considerations", "No"]}
          value={data.a5_mobility}
          onChange={(v) => setField("a5_mobility", v)}
          readOnly={empRO}
          renderOption={lang === "th" ? (opt) => <span>{MOBILITY_TH[opt] || opt}</span> : undefined}
        />
      </Field>
      {data.a5_mobility === "Yes, with some considerations" && (
        <Field label={t("If yes, country / considerations", A5_TH.mobilityDetailLabel, lang)}>
          <TextInput value={data.a5_mobility_detail} onChange={(v) => setField("a5_mobility_detail", v)} readOnly={empRO} />
        </Field>
      )}

      <Field label={t("Dimension focus — choose ONE", A5_TH.dimensionLabel, lang)} required>
        <p className="text-xs text-slate-400 mb-2">{t("Please choose only 1 dimension you would like to develop from the following list.", A5_TH.dimensionHint, lang)}</p>
        <RadioGroup
          options={A5_DIMENSIONS}
          value={data.a5_dimension}
          onChange={(v) => setField("a5_dimension", v)}
          readOnly={empRO}
          renderOption={lang === "th" ? (opt) => <span>{DIMENSION_TH[opt] || opt}</span> : undefined}
        />
      </Field>

      <Field label={t("Reason for choosing this dimension", A5_TH.edgeLabel, lang)} required>
        <p className="text-xs text-slate-400 mb-2">{t("Reason for choosing the above dimension to develop for your growth journey?", A5_TH.edgeHint, lang)}</p>
        <TextArea value={data.a5_edge} onChange={(v) => setField("a5_edge", v)} readOnly={empRO} rows={2} />
      </Field>

      <Field label={t("How You Will Grow: Through Doing & Experiencing", A5_TH.expLabel, lang)} required>
        <p className="text-xs text-slate-400 mb-2">{t("For each entry, name the experience or assignment (up to 3), what you'll deliberately try to grow through it, and roughly when. The specificity matters.", A5_TH.expHint, lang)}</p>
        <p className="text-xs text-slate-400 mb-2 italic">{t("Example: \"Pair with a colleague on Q1 cross-team coordination so they can lead by Q2\" lands better than \"Take on more cross-team work.\"", "ตัวอย่าง: \"เรียนรู้การทำงานเป็นทีมร่วมกับแผนกขาย เพื่อบรรลุเป้าหมายร่วมกันภายในไตรมาส 1\"", lang)}</p>
        <GrowthExperimentTable label={t("e.g. Pair with a colleague on Q1 cross-team coordination so they can lead by Q2", "รายละเอียด", lang)} values={[data.a5_exp1, data.a5_exp2, data.a5_exp3]} onChange={setExp} readOnly={empRO} />
      </Field>

      <Field label={t("How You Will Grow: Through Learning", A5_TH.learnLabel, lang)} required>
        <p className="text-xs text-slate-400 mb-2">{t("For each entry, name the programme/course/coaching, what skill or perspective you're going there to build (up to 3), and roughly when.", A5_TH.learnHint, lang)}</p>
        <p className="text-xs text-slate-400 mb-2 italic">{t("Example: \"Coaching for Managers MEP, Q2 cohort, to build inquiry-based coaching\" lands better than \"Take a leadership course.\"", "ตัวอย่าง: \"เข้าร่วมหลักสูตรพัฒนาการสื่อสาร ภายในไตรมาส 1 เพื่อฝึกฝนการนำเสนอผลงานกับหัวหน้างานให้ดียิ่งขึ้น\"", lang)}</p>
        <GrowthExperimentTable label={t("e.g. Coaching for Managers MEP, Q2 cohort, to build inquiry-based coaching", "รายละเอียด", lang)} values={[data.a5_learn1, data.a5_learn2, data.a5_learn3]} onChange={setLearn} readOnly={empRO} />
      </Field>

      <Field label={t("Support you need from your Direct Manager", A5_TH.supportLabel, lang)} required>
        <p className="text-xs text-slate-400 mb-2">{t("How could your Direct Manager do, say, or perhaps stop doing that would most help you in your growth experiment(s)? Specifics are gold here.", A5_TH.supportHint, lang)}</p>
        <TextArea value={data.a5_support} onChange={(v) => setField("a5_support", v)} readOnly={empRO} rows={4} />
      </Field>

      <ManagerFieldBlock
        label={t("By Direct Manager", A5_TH.managerLabel, lang)}
        prompts={[
          t("How can this plan build on the Employee's existing strengths?", "แผนนี้จะสามารถต่อยอดจากจุดแข็งที่พนักงานมีอยู่แล้วได้อย่างไร?", lang),
          t("How could you support the Employee's on his/her growth experiment(s)?", "ในฐานะหัวหน้างานจะสามารถสนับสนุนพนักงานอย่างไรบ้างเพื่อให้บรรลุเป้าหมาย?", lang),
        ]}
        value={data.a5_manager}
        onChange={(v) => setField("a5_manager", v)}
        role={role}
        locked={mgrLocked}
        waitingForEmployee={waitingForEmployee}
      />
    </div>
  );
}

function CommitBlock({ empChecked, empDate, mgrChecked, mgrDate, onEmp, onMgr, role, empDisabled, mgrDisabled }) {
  const { lang } = useRosterCtx();
  return (
    <div className="mt-2 rounded-xl border p-4 grid grid-cols-1 min-[480px]:grid-cols-2 gap-4" style={{ borderColor: BRAND.line, backgroundColor: "#F6FBFA" }}>
      <div>
        <p className="text-xs font-semibold uppercase tracking-wide mb-2" style={{ color: BRAND.deep }}>
          {t("Employee", "พนักงาน", lang)}
        </p>
        <label className="flex items-start gap-2 text-sm mb-2">
          <input type="checkbox" checked={!!empChecked} disabled={role !== "employee" || empDisabled} onChange={(e) => onEmp(e.target.checked, empDate)} className="w-4 h-4 mt-0.5" />
          {t("I have reviewed and confirmed my answers in this evaluation", "ฉันได้ตรวจสอบและยืนยันคำตอบในแบบประเมิน", lang)}
        </label>
        <label className="text-xs text-slate-400 mb-1 block">{t("As of date", "ณ วันที่", lang)}</label>
        <TextInput type="date" value={empDate} onChange={(v) => onEmp(empChecked, v)} readOnly={role !== "employee" || empDisabled} />
      </div>
      <div>
        <p className="text-xs font-semibold uppercase tracking-wide mb-2" style={{ color: BRAND.deep }}>
          {t("Direct manager", "หัวหน้างาน", lang)}
        </p>
        <label className="flex items-start gap-2 text-sm mb-2">
          <input type="checkbox" checked={!!mgrChecked} disabled={role !== "manager" || mgrDisabled} onChange={(e) => onMgr(e.target.checked, mgrDate)} className="w-4 h-4 mt-0.5" />
          {t("I have reviewed and confirmed my answers in this evaluation", "ฉันได้ตรวจสอบและยืนยันคำตอบในแบบประเมิน", lang)}
        </label>
        <label className="text-xs text-slate-400 mb-1 block">{t("As of date", "ณ วันที่", lang)}</label>
        <TextInput type="date" value={mgrDate} onChange={(v) => onMgr(mgrChecked, v)} readOnly={role !== "manager" || mgrDisabled} />
      </div>
    </div>
  );
}

function ManagerFieldBlock({ label, prompts, value, onChange, role, placeholder, rows = 4, locked, waitingForEmployee }) {
  const { lang } = useRosterCtx();
  if (role === "employee") {
    if (!value) return null; // hidden entirely until the manager has actually responded
    return (
      <div className="mb-4">
        <p className="text-sm font-medium mb-1.5 flex items-center gap-1.5" style={{ color: BRAND.deep }}>
          <UsersIcon className="w-3.5 h-3.5" style={{ color: BRAND.teal }} /> {label}
        </p>
        <p className="text-sm text-slate-700 mb-2">{boldManagerLead(prompts?.[0])}</p>
        {prompts?.length > 1 && <PromptList prompts={prompts.slice(1)} />}
        <div className="rounded-xl px-3.5 py-3 text-sm leading-relaxed" style={{ backgroundColor: "#EAF7F6", border: `1px solid ${BRAND.teal}`, color: "#0B3B3F" }}>
          {value}
        </div>
      </div>
    );
  }
  if (waitingForEmployee) {
    return (
      <div className="mb-4">
        <p className="text-sm font-medium mb-1.5 flex items-center gap-1.5" style={{ color: BRAND.deep }}>
          <UsersIcon className="w-3.5 h-3.5" style={{ color: BRAND.teal }} /> {label}
        </p>
        <div className="rounded-xl px-3.5 py-3 text-sm flex items-start gap-2" style={{ backgroundColor: "#F6FBFA", border: `1px dashed ${BRAND.line}`, color: BRAND.gray }}>
          <Clock className="w-4 h-4 mt-0.5 shrink-0" />
          {t("Waiting for the employee to submit their self-assessment before you can add feedback here.", "รอพนักงานส่งแบบประเมินตนเองก่อน จึงจะสามารถให้ความเห็นในส่วนนี้ได้", lang)}
        </div>
      </div>
    );
  }
  return (
    <Field label={label} required>
      <p className="text-sm text-slate-700 mb-2">{boldManagerLead(prompts?.[0])}</p>
      {prompts?.length > 1 && <PromptList prompts={prompts.slice(1)} />}
      <TextArea value={value} onChange={onChange} readOnly={role !== "manager" || !!locked} placeholder={placeholder} rows={rows} />
    </Field>
  );
}

const SNAPSHOT_META = {
  Flourishing: { icon: TreeDeciduous, color: BRAND.mint, order: 0 },
  Generative: { icon: Sprout, color: BRAND.teal, order: 1 },
  Sustaining: { icon: Leaf, color: BRAND.gray, order: 2 },
  Depleting: { icon: Link2Off, color: BRAND.red, order: 3 },
};
const SNAPSHOT_ORDER = ["Flourishing", "Generative", "Sustaining", "Depleting"];

function SnapshotCard({ optKey, selected, onClick, disabled }) {
  const { lang } = useRosterCtx();
  const meta = SNAPSHOT_META[optKey];
  const Icon = meta.icon;
  const opt = SNAPSHOT_OPTIONS.find((o) => o.key === optKey);
  const label = lang === "th" ? SNAPSHOT_TH[optKey]?.label || optKey : optKey;
  const desc = lang === "th" ? SNAPSHOT_TH[optKey]?.desc || opt.desc : opt.desc;
  return (
    <button
      type="button"
      disabled={disabled}
      onClick={onClick}
      className={`flex-1 min-w-[150px] text-left rounded-2xl border-2 p-3.5 transition ${disabled && !selected ? "opacity-60" : ""}`}
      style={selected ? { borderColor: meta.color, backgroundColor: `${meta.color}1A` } : { borderColor: BRAND.line, backgroundColor: "white" }}
    >
      <div
        className="w-8 h-8 rounded-full flex items-center justify-center mb-2"
        style={selected ? { backgroundColor: meta.color, color: "white" } : { backgroundColor: "#F1F5F4", color: meta.color }}
      >
        <Icon className="w-4 h-4" />
      </div>
      <p className="text-sm font-semibold" style={{ color: selected ? meta.color : "#1e293b" }}>{label}</p>
      <p className="text-xs text-slate-500 mt-0.5 leading-snug">{desc}</p>
    </button>
  );
}

function SnapshotSelector({ value, onChange, readOnly }) {
  return (
    <div className="flex flex-wrap gap-2">
      {SNAPSHOT_ORDER.map((k) => (
        <SnapshotCard key={k} optKey={k} selected={value === k} disabled={readOnly} onClick={() => onChange && onChange(k)} />
      ))}
    </div>
  );
}

function GrowthStageCard({ stageObj, index, selected, onClick, disabled, lang }) {
  const Icon = GROWTH_STAGE_ICONS[index] || Circle;
  const color = GROWTH_STAGE_COLOR;
  const label = lang === "th" ? STAGE_TH[stageObj.stage] || stageObj.stage : stageObj.stage;
  return (
    <button
      type="button"
      disabled={disabled}
      onClick={onClick}
      className={`flex-1 min-w-[130px] text-left rounded-2xl border-2 p-3.5 transition ${disabled && !selected ? "opacity-60" : ""}`}
      style={selected ? { borderColor: color, backgroundColor: `${color}1A` } : { borderColor: BRAND.line, backgroundColor: "white" }}
    >
      <div
        className="w-8 h-8 rounded-full flex items-center justify-center mb-2"
        style={selected ? { backgroundColor: color, color: "white" } : { backgroundColor: "#F1F5F4", color }}
      >
        <Icon className="w-4 h-4" />
      </div>
      <p className="text-sm font-semibold leading-snug" style={{ color: selected ? color : "#1e293b" }}>{label}</p>
    </button>
  );
}

function GrowthStageSelector({ stages, value, onChange, readOnly, lang }) {
  return (
    <div className="flex flex-wrap gap-2">
      {stages.map((s, i) => (
        <GrowthStageCard key={s.stage} stageObj={s} index={i} selected={value === s.stage} disabled={readOnly} onClick={() => onChange && onChange(s.stage)} lang={lang} />
      ))}
    </div>
  );
}

function SnapshotResultCard({ value }) {
  const { lang } = useRosterCtx();
  const meta = SNAPSHOT_META[value];
  const Icon = meta.icon;
  const opt = SNAPSHOT_OPTIONS.find((o) => o.key === value);
  const label = lang === "th" ? SNAPSHOT_TH[value]?.label || value : value;
  const desc = lang === "th" ? SNAPSHOT_TH[value]?.desc || opt.desc : opt.desc;
  return (
    <div className="flex items-center gap-3 rounded-2xl p-3.5" style={{ backgroundColor: `${meta.color}1A`, border: `1px solid ${meta.color}` }}>
      <div className="w-9 h-9 rounded-full flex items-center justify-center shrink-0" style={{ backgroundColor: meta.color, color: "white" }}>
        <Icon className="w-4.5 h-4.5" />
      </div>
      <div>
        <p className="text-sm font-semibold" style={{ color: meta.color }}>{label}</p>
        <p className="text-xs text-slate-500">{desc}</p>
      </div>
    </div>
  );
}

function ValueAddScale({ selfKey, managerKey }) {
  const { lang } = useRosterCtx();
  if (!selfKey && !managerKey) return null;
  const pos = (k) => (SNAPSHOT_META[k].order / 3) * 100;
  return (
    <div className="mt-4 mb-1">
      <p className="text-xs font-semibold uppercase tracking-wide mb-2" style={{ color: BRAND.gray }}>{t("Scale — Flourishing to Depleting", "สเกล — สร้างคุณค่าและเปล่งประกาย ไปจนถึง ส่งผลให้การทำงานถดถอยลง", lang)}</p>
      <div className="relative h-2.5 rounded-full mt-4 mb-4" style={{ background: `linear-gradient(to right, ${BRAND.mint}, ${BRAND.teal}, ${BRAND.gray}, ${BRAND.red})` }}>
        {selfKey && (
          <div
            className="absolute -top-3.5 w-4 h-4 rounded-full bg-white border-2 flex items-center justify-center"
            style={{ left: `calc(${pos(selfKey)}% - 8px)`, borderColor: SNAPSHOT_META[selfKey].color }}
            title={`Self: ${selfKey}`}
          />
        )}
        {managerKey && (
          <div
            className="absolute top-3 w-4 h-4 rounded-full border-2 border-white"
            style={{ left: `calc(${pos(managerKey)}% - 8px)`, backgroundColor: SNAPSHOT_META[managerKey].color }}
            title={`Manager (final): ${managerKey}`}
          />
        )}
      </div>
      <div className="flex items-center gap-4 text-xs text-slate-500">
        {selfKey && <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-full bg-white border-2 inline-block" style={{ borderColor: SNAPSHOT_META[selfKey].color }} /> {t("Self", "ตนเอง", lang)}</span>}
        {managerKey && <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-full inline-block" style={{ backgroundColor: SNAPSHOT_META[managerKey].color }} /> {t("Manager (final)", "หัวหน้างาน (สุดท้าย)", lang)}</span>}
      </div>
    </div>
  );
}

function B4Page({ data, setField, role }) {
  const { lang } = useRosterCtx();
  const th = SECTIONS_TH.b4;
  const selfLocked = role === "employee" && !!data.employee_submitted;
  const mgrLocked = role === "manager" && !!data.manager_submitted;
  const waitingForEmployee = role === "manager" && !data.manager_submitted && !data.employee_submitted;
  return (
    <div>
      <SectionHeader pillar="Overall" code={t("B4 · Overall", "B4 · ภาพรวม", lang)} />
      <Field label={t("By Employee", "โดยพนักงาน", lang)} required>
        <p className="text-sm text-slate-700 mb-2">{t("Any other reflections to yourself", th.selfOther, lang)}</p>
        <TextArea value={data.b4_self} onChange={(v) => setField("b4_self", v)} readOnly={role !== "employee" || selfLocked} rows={3} />
      </Field>

      <ManagerFieldBlock
        label={t("Manager Feedback", UI_TH.managerFeedback, lang)}
        prompts={[t("Any other feedback for the Employee", th.managerOther, lang)]}
        value={data.b4_manager}
        onChange={(v) => setField("b4_manager", v)}
        role={role}
        locked={mgrLocked}
        waitingForEmployee={waitingForEmployee}
        rows={3}
      />

      <div className="mb-5">
        <p className="text-sm font-medium mb-2 flex items-center gap-1.5" style={{ color: BRAND.deep }}>
          <User className="w-3.5 h-3.5" style={{ color: BRAND.mint }} /> {t("From your self-reflection, this is your snapshot of Value-Add", th.selfSnapshotLabel, lang)}
          <span className="ml-1" style={{ color: BRAND.red }}>*</span>
        </p>
        <SnapshotSelector value={data.b4_snapshot_self} onChange={(v) => setField("b4_snapshot_self", v)} readOnly={role !== "employee" || selfLocked} />
      </div>

      <div className="mb-2">
        <p className="text-sm font-medium mb-2 flex items-center gap-1.5" style={{ color: BRAND.deep }}>
          <UsersIcon className="w-3.5 h-3.5" style={{ color: BRAND.teal }} /> {t("As the Direct Manager, this is the Employee's snapshot of Value-Add (final)", th.managerSnapshotLabel, lang)}
          {role === "manager" && <span className="ml-1" style={{ color: BRAND.red }}>*</span>}
        </p>
        {role === "employee" ? (
          data.b4_snapshot_manager ? (
            <SnapshotResultCard value={data.b4_snapshot_manager} />
          ) : (
            <div className="rounded-xl px-3.5 py-3 text-sm text-slate-500" style={{ backgroundColor: "#F6FBFA", border: `1px solid ${BRAND.line}` }}>
              {t("Waiting for your manager's final assessment.", th.waitingSnapshot, lang)}
            </div>
          )
        ) : role === "manager" ? (
          <>
            <p className="text-xs text-slate-400 mb-2">{t("This is the official, final scale used for organisational reporting.", "นี่คือมาตราวัดที่เป็นทางการและสุดท้าย ใช้สำหรับการรายงานระดับองค์กร", lang)}</p>
            <SnapshotSelector value={data.b4_snapshot_manager} onChange={(v) => setField("b4_snapshot_manager", v)} readOnly={mgrLocked || waitingForEmployee} />
          </>
        ) : data.b4_snapshot_manager ? (
          <SnapshotResultCard value={data.b4_snapshot_manager} />
        ) : (
          <p className="text-xs text-slate-400">{t("Not yet assessed.", "ยังไม่ได้ประเมิน", lang)}</p>
        )}
      </div>

      <ValueAddScale selfKey={data.b4_snapshot_self} managerKey={data.b4_snapshot_manager} />

      <div className="mt-5">
        <p className="text-xs font-semibold uppercase tracking-wide mb-2 flex items-center gap-1.5" style={{ color: BRAND.gray }}>
          <CheckCircle2 className="w-3.5 h-3.5" /> {t("Final commitment", "การยืนยันขั้นสุดท้าย", lang)}
        </p>
        <CommitBlock
          empChecked={data.b4_commit_emp} empDate={data.b4_commit_emp_date}
          mgrChecked={data.b4_commit_mgr} mgrDate={data.b4_commit_mgr_date}
          onEmp={(c, d) => { setField("b4_commit_emp", c); setField("b4_commit_emp_date", d); }}
          onMgr={(c, d) => { setField("b4_commit_mgr", c); setField("b4_commit_mgr_date", d); }}
          role={role}
          empDisabled={selfLocked}
          mgrDisabled={mgrLocked || waitingForEmployee}
        />
      </div>
    </div>
  );
}

function PrintView({ person, manager, data, visibleSections, onBack }) {
  const printableSections = visibleSections.filter((s) => s.id !== "cover");
  const fieldLabel = (s, which) => {
    if (s.id === "a4" && which === "self") return "Self-stage + Self-reflection";
    return which === "self" ? "Self" : "Manager";
  };
  return (
    <div className="min-h-screen" style={{ backgroundColor: "white", fontFamily: FONT_STACK }}>
      <style>{`
        ${FONT_IMPORT}
        @media print {
          .no-print { display: none !important; }
          body { background: white !important; }
        }
      `}</style>
      <div className="no-print sticky top-0 z-10 border-b flex items-center justify-between px-4 sm:px-6 py-3" style={{ background: BRAND.bgGradient, borderColor: BRAND.line }}>
        <button onClick={onBack} className="inline-flex items-center gap-1.5 text-sm font-medium" style={{ color: BRAND.deep }}>
          <ArrowLeft className="w-4 h-4" /> Back
        </button>
        <button onClick={() => window.print()} className="inline-flex items-center gap-1.5 px-4 py-2 text-sm font-medium rounded-xl text-white" style={{ backgroundColor: BRAND.primary }}>
          <Printer className="w-4 h-4" /> Print / Save as PDF
        </button>
      </div>

      <div className="max-w-2xl mx-auto px-6 py-8">
        <div className="mb-6 pb-4 border-b" style={{ borderColor: BRAND.line }}>
          <h1 className="text-xl font-semibold" style={{ color: BRAND.deep }}>Growth &amp; Value-Add Journey</h1>
          <p className="text-sm text-slate-500 mt-1">
            {person?.firstName} {person?.lastName} · {person?.bu} · {person?.department}
          </p>
          <p className="text-xs text-slate-400 mt-0.5">Direct manager: {manager ? `${manager.firstName} ${manager.lastName}` : "—"} · Printed {new Date().toLocaleDateString()}</p>
        </div>

        {printableSections.map((s) => {
          const selfVal = s.fields?.self ? data[s.fields.self] : null;
          const mgrVal = s.fields?.manager ? data[s.fields.manager] : null;
          const stageVal = s.fields?.radio ? data[s.fields.radio] : null;
          if (s.id === "a5") {
            return (
              <div key={s.id} className="mb-6 break-inside-avoid">
                <h2 className="text-sm font-semibold mb-2" style={{ color: BRAND.deep }}>{s.nav}</h2>
                <p className="text-xs text-slate-500 mb-1">Mobility: {data.a5_mobility || "—"} · Dimension: {data.a5_dimension || "—"}</p>
                <p className="text-sm text-slate-700 mb-1"><strong>Growth edge:</strong> {data.a5_edge || "—"}</p>
                <p className="text-sm text-slate-700 mb-1"><strong>Experience:</strong> {[data.a5_exp1, data.a5_exp2, data.a5_exp3].filter(Boolean).join("; ") || "—"}</p>
                <p className="text-sm text-slate-700 mb-1"><strong>Learning:</strong> {[data.a5_learn1, data.a5_learn2, data.a5_learn3].filter(Boolean).join("; ") || "—"}</p>
                <p className="text-sm text-slate-700 mb-1"><strong>Support needed:</strong> {data.a5_support || "—"}</p>
                <p className="text-sm text-slate-700"><strong>Manager reflection:</strong> {data.a5_manager || "—"}</p>
              </div>
            );
          }
          if (s.id === "b3") {
            const okrs = [1, 2, 3, 4, 5].map((i) => [data[`b3_okr${i}`], data[`b3_progress${i}`]]).filter(([o]) => o);
            return (
              <div key={s.id} className="mb-6 break-inside-avoid">
                <h2 className="text-sm font-semibold mb-2" style={{ color: BRAND.deep }}>{s.nav}</h2>
                {okrs.length > 0 && (
                  <p className="text-sm text-slate-700 mb-1">
                    <strong>OKRs:</strong> {okrs.map(([o, p], i) => `${o}${p ? ` (${p})` : ""}`).join("; ")}
                  </p>
                )}
                <p className="text-sm text-slate-700 mb-1"><strong>{fieldLabel(s, "self")}:</strong> {selfVal || "—"}</p>
                <p className="text-sm text-slate-700"><strong>{fieldLabel(s, "manager")}:</strong> {mgrVal || "—"}</p>
              </div>
            );
          }
          if (s.id === "b4") {
            return (
              <div key={s.id} className="mb-6 break-inside-avoid">
                <h2 className="text-sm font-semibold mb-2" style={{ color: BRAND.deep }}>{s.nav}</h2>
                <p className="text-sm text-slate-700 mb-1"><strong>Self — other reflections:</strong> {data.b4_self || "—"}</p>
                <p className="text-sm text-slate-700 mb-1"><strong>Manager — other feedback:</strong> {data.b4_manager || "—"}</p>
                <p className="text-sm text-slate-700 mb-1"><strong>Value-Add (self):</strong> {data.b4_snapshot_self || "—"}</p>
                <p className="text-sm text-slate-700"><strong>Value-Add (manager, final):</strong> {data.b4_snapshot_manager || "—"}</p>
              </div>
            );
          }
          return (
            <div key={s.id} className="mb-6 break-inside-avoid">
              <h2 className="text-sm font-semibold mb-2" style={{ color: BRAND.deep }}>{s.nav}</h2>
              {stageVal && <p className="text-xs text-slate-500 mb-1">Self-stage: {stageVal}</p>}
              <p className="text-sm text-slate-700 mb-1"><strong>{fieldLabel(s, "self")}:</strong> {selfVal || "—"}</p>
              <p className="text-sm text-slate-700"><strong>{fieldLabel(s, "manager")}:</strong> {mgrVal || "—"}</p>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function ConfirmDialog({ title, body, confirmLabel, onConfirm, onCancel }) {
  const { lang } = useRosterCtx();
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ backgroundColor: "rgba(15,30,29,0.45)" }}>
      <div className="w-full max-w-sm rounded-2xl p-6" style={{ backgroundColor: BRAND.card, border: `1px solid ${BRAND.line}` }}>
        <div className="w-11 h-11 rounded-full flex items-center justify-center mb-4" style={{ backgroundColor: "#FFF4F3" }}>
          <AlertCircle className="w-5 h-5" style={{ color: BRAND.red }} />
        </div>
        <h2 className="text-base font-semibold mb-2" style={{ color: BRAND.deep }}>{title}</h2>
        <p className="text-sm text-slate-500 leading-relaxed mb-6">{body}</p>
        <div className="flex gap-2">
          <button onClick={onCancel} className="flex-1 text-sm font-medium rounded-xl py-2.5 border" style={{ borderColor: BRAND.line, color: BRAND.deep }}>
            {t("Cancel", "ยกเลิก", lang)}
          </button>
          <button onClick={onConfirm} className="flex-1 text-sm font-medium rounded-xl py-2.5 text-white" style={{ backgroundColor: BRAND.red }}>
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}

function SubmittedScreen({ variant, onContinue }) {
  const { lang } = useRosterCtx();
  const copy = {
    employee: {
      title: t("Your response has been submitted", "คำตอบของคุณถูกส่งแล้ว", lang),
      body: t("Your self-assessment is now locked and can no longer be edited. Your manager will review it next — check back here for their feedback.", "แบบประเมินตนเองของคุณถูกล็อกและไม่สามารถแก้ไขได้อีก หัวหน้างานของคุณจะตรวจสอบต่อไป — กลับมาดูความเห็นได้ที่นี่", lang),
      button: t("View my status", "ดูสถานะของฉัน", lang),
    },
    manager: {
      title: t("Your response has been submitted", "คำตอบของคุณถูกส่งแล้ว", lang),
      body: t("Your review is now locked and can no longer be edited. The employee can now see your feedback.", "การตรวจสอบของคุณถูกล็อกและไม่สามารถแก้ไขได้อีก พนักงานสามารถเห็นความเห็นของคุณได้แล้ว", lang),
      button: t("Back to my team", "กลับไปยังทีมของฉัน", lang),
    },
    delegate: {
      title: t("Your draft has been sent", "ร่างของคุณถูกส่งแล้ว", lang),
      body: t("Your draft is with the direct manager for final confirmation. It's clearly attributed to you, and you can still edit it until they confirm.", "ร่างของคุณอยู่ที่หัวหน้างานเพื่อรอการยืนยันขั้นสุดท้าย มีการระบุชื่อคุณอย่างชัดเจน และคุณยังแก้ไขได้จนกว่าจะมีการยืนยัน", lang),
      button: t("Back to delegated reviews", "กลับไปยังรายการที่ได้รับมอบหมาย", lang),
    },
  }[variant || "manager"];
  return (
    <div className="min-h-screen flex items-center justify-center p-4" style={{ background: BRAND.bgGradient, fontFamily: FONT_STACK }}>
      <style>{`${FONT_IMPORT}`}</style>
      <div className="max-w-sm w-full text-center rounded-2xl overflow-hidden" style={{ border: `1px solid ${BRAND.line}`, backgroundColor: BRAND.card }}>
        <AgateWaves height={100} />
        <div className="px-6 pb-8 pt-2 -mt-2">
          <div className="w-14 h-14 rounded-full flex items-center justify-center mx-auto mb-4 bg-white" style={{ boxShadow: "0 2px 10px rgba(0,79,89,0.12)" }}>
            <CheckCircle2 className="w-7 h-7" style={{ color: BRAND.mint }} />
          </div>
          <h1 className="text-xl font-semibold mb-2" style={{ color: BRAND.deep }}>{copy.title}</h1>
          <p className="text-sm text-slate-500 mb-6 leading-relaxed">{copy.body}</p>
          <button onClick={onContinue} className="w-full text-white text-sm font-medium rounded-xl py-2.5" style={{ backgroundColor: BRAND.primary }}>
            {copy.button}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ========================================================================
   LOGIN SCREEN — scoped to one BU's site, or the separate Admin site.
   A user whose record doesn't belong to this site's BU (or isn't a
   Master Admin, on the Admin site) is rejected even with a correct
   password — this site's code simply never looks outside its own scope.
========================================================================= */

function LoginScreen({ scope, onLogin }) {
  const { lang } = useRosterCtx();
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const isAdminSite = scope.mode === "admin";
  const siteLabel = isAdminSite ? "Master Admin" : scope.bu;

  const submit = async () => {
    if (!username.trim() || !password.trim() || busy) return;
    setBusy(true);
    setError("");
    let match = null;
    try {
      const res = await fetch("/api/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ username: username.trim(), password: password.trim() }),
      });
      if (res.ok) {
        const data = await res.json();
        match = data.user;
      } else {
        setBusy(false);
        const body = await res.json().catch(() => ({}));
        if (res.status === 503) {
          setError(t(
            "The server hasn't finished setting up yet (no roster found). Reload this page once and try again — if it keeps happening, the Redis connection may not be configured in Vercel yet.",
            "เซิร์ฟเวอร์ยังไม่พร้อม (ไม่พบข้อมูลพนักงาน) กรุณาโหลดหน้านี้ใหม่แล้วลองอีกครั้ง — หากยังไม่สำเร็จ อาจเป็นเพราะยังไม่ได้เชื่อมต่อ Redis บน Vercel",
            lang
          ));
        } else if (res.status === 401) {
          setError(t(
            "The server rejected this request (unauthorized) — this is almost always a setup issue (API secret mismatch), not a wrong password. Please check the deployment configuration.",
            "เซิร์ฟเวอร์ปฏิเสธคำขอนี้ (ไม่ได้รับอนุญาต) — มักเกิดจากปัญหาการตั้งค่า (API secret ไม่ตรงกัน) ไม่ใช่รหัสผ่านผิด กรุณาตรวจสอบการตั้งค่าระบบ",
            lang
          ));
        } else if (body.envVarsFound) {
          setError(
            `Server error (${res.status}): ${body.error || "unknown"}. Redis env vars actually present on the server: ${body.envVarsFound.length ? body.envVarsFound.join(", ") : "none found"}.`
          );
        } else {
          setError(t(
            `Server error (${res.status})${body.detail ? `: ${body.detail}` : ""}. Please try again, or check the Vercel function logs for /api/login.`,
            `เซิร์ฟเวอร์ขัดข้อง (สถานะ ${res.status})${body.detail ? `: ${body.detail}` : ""} กรุณาลองใหม่ หรือตรวจสอบ Vercel function logs ของ /api/login`,
            lang
          ));
        }
        return;
      }
    } catch (e) {
      setBusy(false);
      setError(t("Couldn't reach the server. Please try again.", "ไม่สามารถเชื่อมต่อกับเซิร์ฟเวอร์ได้ กรุณาลองใหม่", lang));
      return;
    }
    setBusy(false);
    if (!match) {
      setError(t("Username or password is incorrect. Please check with your P&O team.", "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง กรุณาตรวจสอบกับฝ่าย P&O", lang));
      return;
    }
    // Master Admin can sign in anywhere — every BU link and the admin site.
    const inScope =
      match.role === "master_admin"
        ? true
        : isAdminSite
        ? false
        : match.role === "po_admin"
        ? adminBUs(match).includes(scope.bu)
        : match.bu === scope.bu;
    if (!inScope) {
      setError(
        isAdminSite
          ? t("This account isn't a Master Admin account. Master Admin sign-in only, on this site.", "บัญชีนี้ไม่ใช่บัญชี Master Admin ไซต์นี้สำหรับ Master Admin เท่านั้น", lang)
          : t(`This account isn't registered under ${scope.bu}. Please use your own Company's link.`, `บัญชีนี้ไม่ได้ลงทะเบียนภายใต้ ${scope.bu} กรุณาใช้ลิงก์ของบริษัทคุณเอง`, lang)
      );
      return;
    }
    setError("");
    onLogin(match);
  };

  return (
    <div className="min-h-screen flex items-center justify-center p-4" style={{ background: BRAND.bgGradient, fontFamily: FONT_STACK }}>
      <style>{`${FONT_IMPORT}`}</style>
      <div className="w-full max-w-sm">
        <div className="flex items-center justify-between mb-3">
          <div
            className="inline-flex items-center gap-1.5 text-xs font-medium px-3 py-1 rounded-full"
            style={
              isAdminSite
                ? { backgroundColor: "#FFF4F3", color: BRAND.red, border: "1px solid #F6C8C3" }
                : { backgroundColor: "#EAF7F6", color: BRAND.deep, border: `1px solid ${BRAND.line}` }
            }
          >
            {isAdminSite ? <ShieldCheck className="w-3 h-3" /> : <Lock className="w-3 h-3" />}
            tpcgvajourney/{isAdminSite ? "admin" : SLUG_BY_BU[scope.bu]}
          </div>
          <LangToggle />
        </div>
        <div className="rounded-2xl overflow-hidden relative" style={{ backgroundColor: BRAND.card, border: `1px solid ${BRAND.line}` }}>
          <AgateWaves height={90} />
          <div className="absolute top-3 right-4"><TPCMark /></div>
          <div className="p-8 pt-5">
          <h1 className="text-xl font-semibold mb-1" style={{ color: BRAND.deep }}>
            {isAdminSite ? t("Master Admin Sign-In", "เข้าสู่ระบบ Master Admin", lang) : `${siteLabel} — ${t("Growth & Value-Add Journey", "แบบประเมินการเติบโต และการสร้างผลงานที่มีความหมาย (Growth & Value-Add Journey)", lang)}`}
          </h1>
          {!isAdminSite && <p className="text-xs text-slate-400 mb-1">{BU_FULL_NAME[scope.bu]}</p>}
          {!isAdminSite && <div className="mb-4" />}
          {isAdminSite && <div className="mb-6" />}

          <div className="space-y-4">
            <div>
              <label className="text-xs font-medium mb-1 block" style={{ color: BRAND.deep }}>
                {t("Username", "ชื่อผู้ใช้", lang)}
              </label>
              <TextInput value={username} onChange={setUsername} placeholder={isAdminSite ? "" : "Lalisaman"} />
              {!isAdminSite && (
                <p className="text-xs text-slate-400 mt-1" style={{ fontFamily: FONT_STACK }}>
                  {t(
                    'Your first name "and" the first 3 letters of your last name — e.g. Lalisa Manobal becomes Lalisaman',
                    'ชื่อจริงของคุณ "และ" อักษร 3 ตัวแรกของนามสกุล เช่น Lalisa Manobal เป็น Lalisaman',
                    lang
                  )}
                </p>
              )}
            </div>
            <div>
              <label className="text-xs font-medium mb-1 block" style={{ color: BRAND.deep }}>
                {t("Password", "รหัสผ่าน", lang)}
              </label>
              <TextInput type="password" value={password} onChange={setPassword} placeholder={isAdminSite ? "••••••••" : "01011995"} />
              {!isAdminSite && (
                <p className="text-xs text-slate-400 mt-1" style={{ fontFamily: FONT_STACK }}>
                  {t(
                    "Your birth date as DDMMYYYY — e.g. 11/10/1995 → type 11101995",
                    "วันเกิดของคุณ รูปแบบ วว/ดด/ปปปป เช่น 01/01/1995 พิมพ์ว่า 01011995",
                    lang
                  )}
                </p>
              )}
            </div>

            {error && (
              <div className="text-xs rounded-lg px-3 py-2" style={{ color: "#8A1300", backgroundColor: "#FFF4F3", border: "1px solid #F6C8C3" }}>
                {error}
              </div>
            )}

            <button
              onClick={submit}
              onKeyDown={(e) => e.key === "Enter" && submit()}
              disabled={busy}
              className="w-full text-white text-sm font-medium rounded-xl py-2.5 transition disabled:opacity-60"
              style={{ backgroundColor: isAdminSite ? BRAND.deep : BRAND.primary }}
            >
              {busy ? t("Signing in…", "กำลังเข้าสู่ระบบ…", lang) : t("Sign in", "เข้าสู่ระบบ", lang)}
            </button>
          </div>
          </div>
        </div>
        <p className="text-center text-xs text-slate-400 mt-4">
          {t("Wrong site?", "ผิดไซต์?", lang)} <a href="#/" className="underline" style={{ color: BRAND.primary }}>{t("Go to the site directory", "ไปที่รายชื่อไซต์", lang)}</a>
        </p>
      </div>
    </div>
  );
}

/* ========================================================================
   GATEWAY — prototype-only convenience. In production, each BU's staff
   would simply be given their own bookmarked link directly and would
   never see this list; there would be no shared directory page at all.
========================================================================= */

// Placeholder wordmark only — see note to user. Not the official TPC logo file.
function TPCMark({ size = 36 }) {
  return <img src={TPC_LOGO} alt="TPC" title="TPC" style={{ height: size, width: "auto", display: "block" }} />;
}

function GatewayPage() {
  const { lang } = useRosterCtx();
  const go = (hash) => { window.location.hash = hash; };
  return (
    <div className="min-h-screen flex items-center justify-center p-4" style={{ background: BRAND.bgGradient, fontFamily: FONT_STACK }}>
      <style>{`${FONT_IMPORT}`}</style>
      <div className="w-full max-w-md rounded-2xl overflow-hidden relative" style={{ border: `1px solid ${BRAND.line}`, backgroundColor: BRAND.card }}>
        <AgateWaves height={120} />
        <div className="absolute top-3 right-4 flex items-center gap-2">
          <LangToggle />
          <TPCMark />
        </div>
        <div className="p-6 -mt-8 relative">
          <p className="text-xs text-slate-400 mb-2 text-center">tpcgvajourney — {t("site directory", "รายชื่อไซต์", lang)}</p>
          <h1 className="text-xl font-semibold mb-1 text-center" style={{ color: BRAND.deep }}>
            {t("Growth & Value-Add Journey", "แบบประเมินการเติบโต และการสร้างผลงานที่มีความหมาย (Growth & Value-Add Journey)", lang)}
          </h1>
        <div className="rounded-2xl overflow-hidden divide-y mb-4" style={{ border: `1px solid ${BRAND.line}`, backgroundColor: BRAND.card, borderColor: BRAND.line }}>
          {BUS.map((bu) => (
            <button
              key={bu}
              type="button"
              onClick={() => go(`#/${SLUG_BY_BU[bu]}`)}
              className="w-full flex items-center gap-3 px-5 py-4 hover:bg-slate-50 transition text-left"
            >
              <div className="w-9 h-9 rounded-xl flex items-center justify-center shrink-0" style={{ backgroundColor: "#F0F8F8", color: BRAND.primary }}>
                <Building2 className="w-4 h-4" />
              </div>
              <div className="flex-1">
                <p className="text-sm font-medium text-slate-800">{bu} · <span className="text-slate-500 font-normal">{BU_FULL_NAME[bu]}</span></p>
                <p className="text-xs text-slate-400">tpcgvajourney/{SLUG_BY_BU[bu]}</p>
              </div>
              <ChevronRight className="w-4 h-4 text-slate-300 shrink-0" />
            </button>
          ))}
        </div>
        <button
          type="button"
          onClick={() => go("#/admin")}
          className="w-full flex items-center gap-3 px-5 py-4 rounded-2xl text-left"
          style={{ border: `1px dashed #F6C8C3`, backgroundColor: "#FFF4F3" }}
        >
          <div className="w-9 h-9 rounded-xl flex items-center justify-center shrink-0" style={{ backgroundColor: "white", color: BRAND.red }}>
            <ShieldCheck className="w-4 h-4" />
          </div>
          <div className="flex-1">
            <p className="text-sm font-medium" style={{ color: BRAND.red }}>tpcgvajourney/admin</p>
          </div>
          <ChevronRight className="w-4 h-4 shrink-0" style={{ color: BRAND.red }} />
        </button>
        </div>
      </div>
    </div>
  );
}

/* ========================================================================
   MANAGER — TEAM LIST (choose whose journey to open)
========================================================================= */

function DelegateRowControl({ employee, manager, locked }) {
  const { roster, setRoster, lang } = useRosterCtx();
  const [open, setOpen] = useState(false);
  const [pick, setPick] = useState(employee.reviewDelegateId || "");
  // Same-BU colleagues only, excluding the employee and the direct manager themselves.
  const candidates = roster.filter((p) => p.role === "staff" && p.bu === employee.bu && p.id !== employee.id && p.id !== manager.id);
  const current = employee.reviewDelegateId ? roster.find((p) => p.id === employee.reviewDelegateId) : null;

  const persist = async (delegateId) => {
    const next = roster.map((p) => (p.id === employee.id ? { ...p, reviewDelegateId: delegateId } : p));
    setRoster(next);
    await saveRoster(next);
    if (delegateId) {
      const d = roster.find((p) => p.id === delegateId);
      appendLog(`${manager.firstName} ${manager.lastName}`, "Delegated Review", `Delegated ${employee.firstName} ${employee.lastName}'s manager review to ${d?.firstName} ${d?.lastName}.`);
    } else {
      appendLog(`${manager.firstName} ${manager.lastName}`, "Delegation Cleared", `Cleared the review delegate for ${employee.firstName} ${employee.lastName}.`);
    }
  };

  if (locked) {
    // Review is finalized — the delegate assignment becomes a permanent
    // record of who actually did the work and can no longer be reassigned.
    if (!current) return <span className="text-sm text-slate-400">—</span>;
    return (
      <div onClick={(e) => e.stopPropagation()}>
        <span
          className="text-xs font-medium inline-flex items-center gap-1"
          style={{ color: BRAND.primary }}
          title={t("Locked — the review is complete, so this record can no longer be changed.", "ล็อกไว้ — การตรวจสอบเสร็จสมบูรณ์แล้ว จึงไม่สามารถเปลี่ยนแปลงบันทึกนี้ได้", lang)}
        >
          <Lock className="w-3 h-3" />
          {current.firstName} {current.lastName}
        </span>
      </div>
    );
  }

  return (
    <div onClick={(e) => e.stopPropagation()}>
      <button
        onClick={() => setOpen(!open)}
        className="text-xs font-medium inline-flex items-center gap-1 underline decoration-dotted underline-offset-2"
        style={{ color: current ? BRAND.primary : BRAND.gray }}
      >
        {current ? `${current.firstName} ${current.lastName}` : t("— Assign", "— มอบหมาย", lang)}
      </button>
      {open && (
        <div className="mt-1.5 flex flex-wrap items-center gap-1.5 rounded-lg p-2" style={{ backgroundColor: "#F6FBFA", border: `1px solid ${BRAND.line}` }}>
          <select value={pick} onChange={(e) => setPick(e.target.value)} className="text-xs rounded-lg border px-2 py-1 focus:outline-none" style={{ borderColor: BRAND.line }}>
            <option value="">{t("Choose a colleague (same BU)…", "เลือกเพื่อนร่วมงาน (บริษัทเดียวกัน)…", lang)}</option>
            {candidates.map((p) => <option key={p.id} value={p.id}>{p.firstName} {p.lastName}</option>)}
          </select>
          <button onClick={() => { persist(pick || null); setOpen(false); }} disabled={!pick} className="text-xs font-medium rounded-lg px-2.5 py-1 text-white disabled:opacity-40" style={{ backgroundColor: BRAND.primary }}>
            {t("Set", "ตั้งค่า", lang)}
          </button>
          {current && (
            <button onClick={() => { setPick(""); persist(null); setOpen(false); }} className="text-xs font-medium rounded-lg px-2.5 py-1 border" style={{ borderColor: BRAND.line, color: BRAND.deep }}>
              {t("Clear", "ล้าง", lang)}
            </button>
          )}
        </div>
      )}
    </div>
  );
}

function ManagerTeamList({ manager, onOpenEmployee, onGoSelf, onLogout, onSwitchUser, actingAs, delegatedCount, onGoDelegated }) {
  const { roster, cycle, lang, cycleEnd } = useRosterCtx();
  const scope = { mode: "bu", bu: manager.bu };
  const [reports, setReports] = useState(null);

  useEffect(() => {
    (async () => {
      const list = directReports(manager.id, roster);
      const withProgress = await Promise.all(
        list.map(async (u) => {
          const d = await loadData(u.id, roster, cycle);
          const merged = { ...defaultData(), ...(d || {}) };
          const visible = SECTIONS.filter((s) => !s.managerOnly || hasReports(u.id, roster));
          return {
            user: u,
            progress: roleProgress(visible, merged, "manager"),
            status: journeyStatus(merged),
            overdue: isOverdue(merged, cycleEnd),
            daysWaiting: merged.employee_submitted && !merged.manager_submitted ? daysSince(merged.employee_submitted_at) : 0,
            snapshotSelf: merged.b4_snapshot_self,
            snapshotManager: merged.b4_snapshot_manager,
            completedAt: merged.manager_submitted_at,
          };
        })
      );
      setReports(withProgress);
    })();
  }, [manager.id, roster, cycle]);

  const pending = (reports || []).filter((r) => r.status === "pending_review" || r.status === "drafted_pending_confirmation");
  const oldestDays = pending.length ? Math.max(...pending.map((r) => r.daysWaiting)) : 0;
  const overdueCount = pending.filter((r) => r.overdue).length;

  return (
    <div className="min-h-screen" style={{ background: BRAND.bgGradient, fontFamily: FONT_STACK }}>
      <style>{`${FONT_IMPORT}`}</style>
      <TopBar user={manager} onLogout={onLogout} onSwitchUser={onSwitchUser} subtitle={actingAs || t("My Team", UI_TH.myTeam, lang)} scope={scope} roleSwitch={actingAs ? undefined : { mode: "team", onGoSelf, onGoTeam: () => {}, delegatedCount, onGoDelegated, showSelf: manager.needsEvaluation !== false }} />
      <div className="max-w-2xl mx-auto px-4 sm:px-6 py-8">
        {actingAs && (
          <button onClick={onGoSelf} className="inline-flex items-center gap-1.5 text-sm font-medium mb-3" style={{ color: BRAND.deep }}>
            <ArrowLeft className="w-4 h-4" /> {t("Back to delegated list", "กลับไปยังรายการที่ได้รับมอบหมาย", lang)}
          </button>
        )}
        <h1 className="text-xl font-semibold mb-1 flex items-center gap-2" style={{ color: BRAND.deep }}>
          <UsersIcon className="w-5 h-5" style={{ color: BRAND.mint }} /> {actingAs ? actingAs : t("My Team", UI_TH.myTeam, lang)}
        </h1>
        <p className="text-sm text-slate-500 mb-1">{t("Review progress for your direct reports.", "ดูความคืบหน้าของผู้ใต้บังคับบัญชาของคุณ", lang)}</p>
        <p className="text-xs text-slate-400 mb-4">{t("An individual review may be delegated to a colleague within the same Company; final confirmation remains yours.", "คุณสามารถมอบหมายให้ผู้ใต้บังคับบัญชาของคุณช่วยให้ความคิดเห็นรายบุคคลได้ แต่ทั้งนี้คุณยังคงเป็นผู้ยืนยันผลการประเมินขั้นสุดท้าย", lang)}</p>

        {pending.length > 0 && (
          <div
            className="rounded-xl px-4 py-3 mb-4 flex items-start gap-2.5"
            style={{ backgroundColor: overdueCount ? "#FFF4F3" : "#FFF9E8", border: `1px solid ${overdueCount ? "#F6C8C3" : "#F3E3A8"}` }}
          >
            <AlertCircle className="w-4 h-4 mt-0.5 shrink-0" style={{ color: overdueCount ? BRAND.red : "#8A6D00" }} />
            <p className="text-sm" style={{ color: overdueCount ? "#8A1300" : "#8A6D00" }}>
              {lang === "th" ? (
                <>
                  <span className="font-medium">รอตรวจสอบ {pending.length} รายการ</span>
                  {oldestDays > 0 && <> รอนานที่สุด {oldestDays} วัน</>}
                  {overdueCount > 0 && <> {overdueCount} รายการเกินกำหนดเส้นตายของรอบนี้แล้ว{cycleEnd ? ` (${new Date(cycleEnd).toLocaleDateString("th-TH")})` : ""}</>}
                </>
              ) : (
                <>
                  <span className="font-medium">{pending.length} pending review{pending.length > 1 ? "s" : ""}</span>
                  {oldestDays > 0 && <> , oldest waiting {oldestDays} day{oldestDays !== 1 ? "s" : ""}</>}.
                  {overdueCount > 0 && <> {overdueCount} past this cycle's deadline{cycleEnd ? ` (${new Date(cycleEnd).toLocaleDateString()})` : ""}.</>}
                </>
              )}
            </p>
          </div>
        )}

        <div className="rounded-2xl overflow-hidden divide-y" style={{ border: `1px solid ${BRAND.line}`, backgroundColor: BRAND.card }}>
          {reports === null && <div className="px-5 py-6 text-sm text-slate-400">Loading…</div>}
          {reports?.map(({ user, progress, status, overdue, daysWaiting, snapshotSelf, snapshotManager, completedAt }) => {
            const sMeta = JOURNEY_STATUS_META[status];
            const SIcon = sMeta.icon;
            const isCompleted = status === "completed";
            const mgrMeta = isCompleted && snapshotManager ? SNAPSHOT_META[snapshotManager] : null;
            const MgrIcon = mgrMeta?.icon;
            const mgrLabel = mgrMeta ? (lang === "th" ? SNAPSHOT_TH[snapshotManager]?.label || snapshotManager : snapshotManager) : null;
            const selfMeta = isCompleted && snapshotSelf ? SNAPSHOT_META[snapshotSelf] : null;
            const SelfIcon = selfMeta?.icon;
            const selfLabel = selfMeta ? (lang === "th" ? SNAPSHOT_TH[snapshotSelf]?.label || snapshotSelf : snapshotSelf) : null;
            return (
            <div
              key={user.id}
              onClick={() => onOpenEmployee(user.id)}
              className="w-full flex items-center gap-3 px-5 py-4 text-left hover:bg-slate-50 transition cursor-pointer"
            >
              <Avatar name={`${user.firstName} ${user.lastName}`} size={34} />
              <div className="min-w-0 flex-1">
                <p className="text-sm font-medium text-slate-800 truncate">
                  {user.firstName} {user.lastName}
                </p>
                <p className="text-xs text-slate-400 mb-2.5">{user.department}</p>

                <div className="flex flex-wrap items-center gap-x-5 gap-y-2">
                  <div className="flex items-center gap-1.5 whitespace-nowrap">
                    <span className="text-[11px] text-slate-400">{t("Status", UI_TH.status, lang)}:</span>
                    <span className="inline-flex items-center gap-1 text-xs font-medium px-2.5 py-1 rounded-full" style={{ backgroundColor: `${sMeta.color}17`, color: sMeta.color }}>
                      <SIcon className="w-3 h-3 shrink-0" /> {t(sMeta.label, STATUS_LABEL_TH[status], lang)}
                    </span>
                  </div>
                  {!actingAs && (
                    <div className="flex items-center gap-1.5 whitespace-nowrap" onClick={(e) => e.stopPropagation()}>
                      <span className="text-[11px] text-slate-400">{t("Delegated to", "มอบหมายให้", lang)}:</span>
                      <DelegateRowControl employee={user} manager={manager} locked={isCompleted} />
                    </div>
                  )}
                  {overdue && (
                    <div className="flex items-center gap-1.5 whitespace-nowrap">
                      <span className="inline-flex items-center gap-1 text-xs font-medium px-2.5 py-1 rounded-full" style={{ backgroundColor: "#FFF4F3", color: BRAND.red }}>
                        <AlertCircle className="w-3 h-3" /> {t(`${daysWaiting}d overdue`, `เกินกำหนด ${daysWaiting} วัน`, lang)}
                      </span>
                    </div>
                  )}
                  {isCompleted && mgrMeta && (
                    <div className="flex items-center gap-1.5 whitespace-nowrap">
                      <span className="text-[11px] text-slate-400">{t("Manager's Rating", "คะแนนของหัวหน้างาน", lang)}:</span>
                      <span className="inline-flex items-center gap-1 text-xs font-medium px-2.5 py-1 rounded-full" style={{ backgroundColor: `${mgrMeta.color}17`, color: mgrMeta.color }} title={completedAt ? `${t("Finalized", "สรุปผลเมื่อ", lang)} ${new Date(completedAt).toLocaleDateString(lang === "th" ? "th-TH" : undefined)}` : undefined}>
                        <MgrIcon className="w-3 h-3 shrink-0" /> {mgrLabel}
                      </span>
                    </div>
                  )}
                  {isCompleted && selfMeta && (
                    <div className="flex items-center gap-1.5 whitespace-nowrap">
                      <span className="text-[11px] text-slate-400">{t("Employee's Rating", "คะแนนของพนักงาน", lang)}:</span>
                      <span className="inline-flex items-center gap-1 text-xs font-medium px-2.5 py-1 rounded-full" style={{ backgroundColor: `${selfMeta.color}17`, color: selfMeta.color }}>
                        <SelfIcon className="w-3 h-3 shrink-0" /> {selfLabel}
                      </span>
                    </div>
                  )}
                </div>
              </div>
              <div className="flex items-center gap-3 shrink-0">
                <div className="w-24 h-1.5 rounded-full overflow-hidden hidden sm:block" style={{ backgroundColor: "#DCEEEC" }}>
                  <div className="h-full rounded-full" style={{ width: `${progress}%`, backgroundColor: BRAND.mint }} />
                </div>
                <span className="text-xs font-semibold tabular-nums w-9 text-right" style={{ color: BRAND.primary }}>
                  {progress}%
                </span>
                <ChevronRight className="w-4 h-4 text-slate-300" />
              </div>
            </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

function MyDelegatedReviews({ delegate, reports, onOpenEmployee, onExit, canExitToTeam, onLogout, onSwitchUser }) {
  const { roster, lang } = useRosterCtx();
  const scope = { mode: "bu", bu: delegate.bu };
  return (
    <div className="min-h-screen" style={{ background: BRAND.bgGradient, fontFamily: FONT_STACK }}>
      <style>{`${FONT_IMPORT}`}</style>
      <TopBar user={delegate} onLogout={onLogout} onSwitchUser={onSwitchUser} subtitle={t("Delegated reviews", "รายการที่ได้รับมอบหมาย", lang)} scope={scope} />
      <div className="max-w-2xl mx-auto px-4 sm:px-6 py-8">
        {onExit && (
          <button onClick={onExit} className="inline-flex items-center gap-1.5 text-sm font-medium mb-4" style={{ color: BRAND.deep }}>
            <ArrowLeft className="w-4 h-4" /> {canExitToTeam ? t("Back to my team", "กลับไปยังทีมของฉัน", lang) : t("Back to my journey", "กลับไปยังแบบประเมินของฉัน", lang)}
          </button>
        )}
        <h1 className="text-xl font-semibold mb-1 flex items-center gap-2" style={{ color: BRAND.deep }}>
          <RefreshCcw className="w-5 h-5" style={{ color: BRAND.mint }} /> {t("Reviews delegated to me", "การตรวจสอบที่ได้รับมอบหมายให้ฉัน", lang)}
        </h1>
        <p className="text-sm text-slate-500 mb-6">
          {t(
            "You may draft manager feedback for the reviews listed below. Final confirmation remains with each person's direct manager, and all entries are attributed to you.",
            "คุณสามารถร่างความเห็นของหัวหน้างานสำหรับรายการด้านล่างนี้ได้ การยืนยันขั้นสุดท้ายยังคงเป็นของหัวหน้างานของแต่ละคน และทุกรายการจะถูกระบุชื่อคุณไว้อย่างชัดเจน",
            lang
          )}
        </p>
        <div className="rounded-2xl overflow-hidden divide-y" style={{ border: `1px solid ${BRAND.line}`, backgroundColor: BRAND.card }}>
          {reports.map((r) => {
            const mgr = roster.find((p) => p.id === r.managerId);
            const asFunctionalManager = r.functionalManagerId === delegate.id;
            return (
              <button key={r.id} onClick={() => onOpenEmployee(r.id)} className="w-full flex items-center gap-3 px-5 py-4 text-left hover:bg-slate-50 transition">
                <Avatar name={`${r.firstName} ${r.lastName}`} size={34} />
                <div className="min-w-0 flex-1">
                  <p className="text-sm font-medium text-slate-800 truncate">{r.firstName} {r.lastName}</p>
                  <p className="text-xs text-slate-400">
                    {asFunctionalManager ? (
                      <span className="inline-flex items-center gap-1"><Building2 className="w-3 h-3" /> {t("Second reviewer", "ผู้ประเมินเพิ่มเติม", lang)}</span>
                    ) : (
                      <>{t("On behalf of", "ในนามของ", lang)} {mgr ? `${mgr.firstName} ${mgr.lastName}` : "—"}</>
                    )} · {r.bu}
                  </p>
                </div>
                <ChevronRight className="w-4 h-4 text-slate-300" />
              </button>
            );
          })}
          {!reports.length && <div className="px-5 py-6 text-sm text-slate-400">{t("No delegated reviews right now.", "ยังไม่มีรายการที่ได้รับมอบหมายในขณะนี้", lang)}</div>}
        </div>
      </div>
    </div>
  );
}

function scopedUsers(scope, roster) {
  if (!scope) return roster;
  if (scope.mode === "admin") return roster.filter((u) => u.role === "master_admin");
  return roster.filter((u) => u.bu === scope.bu || (u.buList && u.buList.includes(scope.bu)));
}

function LangToggle() {
  const { lang, setLang } = useRosterCtx();
  return (
    <div className="flex rounded-full p-0.5" style={{ backgroundColor: "#E9F3F2" }}>
      <button
        onClick={() => setLang("en")}
        className="text-[10px] font-semibold px-2 py-1 rounded-full transition"
        style={lang === "en" ? { backgroundColor: BRAND.primary, color: "white" } : { color: BRAND.deep }}
      >
        EN
      </button>
      <button
        onClick={() => setLang("th")}
        className="text-[10px] font-semibold px-2 py-1 rounded-full transition"
        style={lang === "th" ? { backgroundColor: BRAND.primary, color: "white", fontFamily: FONT_STACK } : { color: BRAND.deep, fontFamily: FONT_STACK }}
      >
        ไทย
      </button>
    </div>
  );
}

function TopBar({ user, onLogout, onSwitchUser, subtitle, scope, roleSwitch }) {
  const { roster, lang } = useRosterCtx();
  const pool = scopedUsers(scope, roster);
  return (
    <div className="sticky top-0 z-10 border-b" style={{ backgroundColor: "rgba(246,251,250,0.92)", backdropFilter: "blur(6px)", borderColor: BRAND.line }}>
      <div style={{ height: 3, background: `linear-gradient(to right, ${BRAND.deep}, ${BRAND.teal}, ${BRAND.mint})` }} />
      <div className="max-w-2xl mx-auto px-4 sm:px-6 py-3.5 flex items-center justify-between gap-3">
        <div className="flex items-center gap-2.5 min-w-0">
          <Avatar name={`${user.firstName} ${user.lastName}`} size={30} />
          <div className="min-w-0">
            <p className="text-sm font-semibold" style={{ color: BRAND.deep }}>
              {user.firstName} {user.lastName}
            </p>
            <p className="text-xs text-slate-400 truncate">
              tpcgvajourney/{scope?.mode === "admin" ? "admin" : SLUG_BY_BU[scope?.bu]} · {subtitle}
            </p>
          </div>
        </div>
        <div className="flex items-center gap-2 shrink-0">
          {roleSwitch && (
            <div className="flex rounded-full p-0.5" style={{ backgroundColor: "#E9F3F2" }}>
              {roleSwitch.showSelf !== false && (
                <button
                  onClick={roleSwitch.onGoSelf}
                  className="flex items-center gap-1 text-xs font-medium px-2.5 py-1 rounded-full transition"
                  style={roleSwitch.mode === "self" ? { backgroundColor: BRAND.primary, color: "white" } : { color: BRAND.deep }}
                >
                  <User className="w-3 h-3" /> {t("My Journey", UI_TH.myJourney, lang)}
                </button>
              )}
              <button
                onClick={roleSwitch.onGoTeam}
                className="flex items-center gap-1 text-xs font-medium px-2.5 py-1 rounded-full transition"
                style={roleSwitch.mode === "team" ? { backgroundColor: BRAND.primary, color: "white" } : { color: BRAND.deep }}
              >
                <UsersIcon className="w-3 h-3" /> {t("My Team", UI_TH.myTeam, lang)}
              </button>
              {roleSwitch.delegatedCount > 0 && (
                <button
                  onClick={roleSwitch.onGoDelegated}
                  className="flex items-center gap-1 text-xs font-medium px-2.5 py-1 rounded-full transition"
                  style={roleSwitch.mode === "delegated" ? { backgroundColor: BRAND.primary, color: "white" } : { color: BRAND.deep }}
                >
                  <RefreshCcw className="w-3 h-3" /> {t("Delegated", UI_TH.delegated, lang)} ({roleSwitch.delegatedCount})
                </button>
              )}
            </div>
          )}
          <LangToggle />
          <button onClick={onLogout} className="p-1.5 text-slate-400 hover:text-red-600" title="Log out">
            <LogOut className="w-4 h-4" />
          </button>
        </div>
      </div>
    </div>
  );
}

/* ========================================================================
   ADMIN — P&O / Master Admin dashboard (Track B only — the firewall holds)
========================================================================= */

const SNAPSHOT_HEX = {
  Flourishing: BRAND.mint,
  Generative: BRAND.teal,
  Sustaining: BRAND.gray,
  Depleting: BRAND.red,
};

function snapshotColor(key) {
  switch (key) {
    case "Flourishing": return { bg: "#EAF7F6", text: BRAND.deep, border: BRAND.mint };
    case "Generative": return { bg: "#EAF7F6", text: BRAND.primary, border: BRAND.teal };
    case "Sustaining": return { bg: "#F3F6F5", text: BRAND.gray, border: BRAND.line };
    case "Depleting": return { bg: "#FFF4F3", text: "#8A1300", border: "#F6C8C3" };
    default: return { bg: "#F3F6F5", text: BRAND.gray, border: BRAND.line };
  }
}

function StatCard({ label, value, sub, icon: Icon, color, info }) {
  return (
    <div className="rounded-2xl p-4" style={{ border: `1px solid ${BRAND.line}`, backgroundColor: BRAND.card }}>
      {Icon && (
        <div className="w-7 h-7 rounded-lg flex items-center justify-center mb-2" style={{ backgroundColor: `${color || BRAND.teal}1A`, color: color || BRAND.teal }}>
          <Icon className="w-3.5 h-3.5" />
        </div>
      )}
      <p className="text-2xl font-semibold" style={{ color: BRAND.deep }}>{value}</p>
      <p className="text-xs text-slate-500 mt-0.5 flex items-center gap-1">
        {label}
        {info && <InfoTip text={info} />}
      </p>
      {sub && <p className="text-xs text-slate-400 mt-1">{sub}</p>}
    </div>
  );
}

// TPC "Agate" brand motif — layered, flowing wave forms in the Turquoise of
// Truth range. Used as soft decorative backgrounds on standalone moments
// (login, directory, confirmation) rather than flat blocks or boxed icons.
function AgateWaves({ height = 220, flip = false }) {
  const uid = useRef(`agate-${Math.random().toString(36).slice(2, 9)}`).current;
  return (
    <svg
      viewBox="0 0 800 260"
      preserveAspectRatio="none"
      style={{ width: "100%", height, display: "block", transform: flip ? "scaleY(-1)" : undefined }}
      aria-hidden="true"
    >
      <defs>
        <linearGradient id={`${uid}-a`} x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stopColor={BRAND.deep} />
          <stop offset="100%" stopColor={BRAND.primary} />
        </linearGradient>
        <linearGradient id={`${uid}-b`} x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stopColor={BRAND.primary} />
          <stop offset="100%" stopColor={BRAND.teal} />
        </linearGradient>
        <linearGradient id={`${uid}-c`} x1="0" y1="0" x2="1" y2="0">
          <stop offset="0%" stopColor={BRAND.teal} />
          <stop offset="100%" stopColor={BRAND.mint} />
        </linearGradient>
      </defs>
      <path d="M0,60 C160,120 260,10 420,55 C580,100 660,30 800,70 L800,0 L0,0 Z" fill={`url(#${uid}-a)`} opacity="0.95" />
      <path d="M0,110 C180,170 300,70 460,110 C610,148 690,90 800,120 L800,0 L0,0 Z" fill={`url(#${uid}-b)`} opacity="0.75" />
      <path d="M0,170 C200,220 320,140 480,172 C630,202 700,150 800,178 L800,0 L0,0 Z" fill={`url(#${uid}-c)`} opacity="0.55" />
    </svg>
  );
}

function Avatar({ name, size = 32, color }) {
  const initials = (name || "").trim().split(/\s+/).map((w) => w[0]).slice(0, 2).join("").toUpperCase() || "?";
  return (
    <div
      className="rounded-full flex items-center justify-center font-semibold shrink-0"
      style={{ width: size, height: size, fontSize: size * 0.38, backgroundColor: color || "#E9F3F2", color: BRAND.deep }}
    >
      {initials}
    </div>
  );
}

function DonutWithLegend({ title, icon: Icon, data, colorMap, total }) {
  return (
    <div className="rounded-2xl p-4" style={{ border: `1px solid ${BRAND.line}`, backgroundColor: BRAND.card }}>
      <p className="text-xs font-semibold uppercase tracking-wide mb-3 flex items-center gap-1.5" style={{ color: BRAND.gray }}>
        {Icon && <Icon className="w-3.5 h-3.5" />} {title}
      </p>
      {total ? (
        <div className="flex items-center gap-4">
          <div className="relative shrink-0" style={{ width: 130, height: 130 }}>
            <ResponsiveContainer width={130} height={130}>
              <PieChart>
                <Pie data={data} dataKey="value" nameKey="name" innerRadius={40} outerRadius={62} paddingAngle={2} stroke="none">
                  {data.map((d) => (
                    <Cell key={d.key || d.name} fill={colorMap[d.key || d.name]} />
                  ))}
                </Pie>
                <Tooltip contentStyle={{ fontSize: 12, borderRadius: 8, borderColor: BRAND.line }} />
              </PieChart>
            </ResponsiveContainer>
            <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
              <span className="text-lg font-semibold" style={{ color: BRAND.deep }}>{total}</span>
              <span className="text-[10px] text-slate-400">total</span>
            </div>
          </div>
          <div className="flex-1 space-y-1.5 min-w-0">
            {data.map((d) => (
              <div key={d.key || d.name} className="flex items-center justify-between gap-2 text-xs">
                <span className="flex items-center gap-1.5 text-slate-600 min-w-0">
                  <span className="w-2 h-2 rounded-full shrink-0" style={{ backgroundColor: colorMap[d.key || d.name] }} />
                  <span className="truncate">{d.name}</span>
                </span>
                <span className="font-semibold whitespace-nowrap" style={{ color: BRAND.deep }}>
                  {d.value} <span className="text-slate-400 font-normal">({total ? Math.round((d.value / total) * 100) : 0}%)</span>
                </span>
              </div>
            ))}
          </div>
        </div>
      ) : (
        <div className="h-[130px] flex items-center justify-center text-xs text-slate-400">No data in scope yet</div>
      )}
    </div>
  );
}

const JOURNEY_STATUS_META = {
  not_started: { label: "Not Started", icon: Circle, color: BRAND.gray },
  in_progress: { label: "In Progress", icon: Clock, color: BRAND.teal },
  pending_review: { label: "By Employee — Pending Manager's Review", icon: AlertCircle, color: "#C9A227" },
  drafted_pending_confirmation: { label: "By Delegate — Pending Final Review", icon: RefreshCcw, color: BRAND.primary },
  completed: { label: "Completed", icon: CheckCircle2, color: BRAND.mint },
};
function journeyStatus(d) {
  if (d.employee_submitted && d.manager_submitted) return "completed";
  if (d.employee_submitted && d.manager_draft_by) return "drafted_pending_confirmation";
  if (d.employee_submitted) return "pending_review";
  const started = ["a1_self", "reflection_date", "a4_self"].some((k) => (d[k] || "").toString().trim());
  return started ? "in_progress" : "not_started";
}

function daysSince(iso) {
  if (!iso) return 0;
  return Math.floor((Date.now() - new Date(iso).getTime()) / (1000 * 60 * 60 * 24));
}
// Overdue is anchored to the cycle's real deadline (an actual end date the
// company sets, e.g. "complete by 15 Oct 2026") — not a per-person relative
// countdown from when they happened to submit. If no end date is set yet,
// nothing is ever flagged overdue.
function isOverdue(d, cycleEnd) {
  if (!cycleEnd) return false;
  const isComplete = !!d.employee_submitted && !!d.manager_submitted;
  if (isComplete) return false;
  return new Date() > new Date(`${cycleEnd}T23:59:59`);
}

function AdminView({ admin, onOpenEmployee, onLogout, onSwitchUser }) {
  const { roster, cycle, cyclesList, startNewCycle, lang, cycleStart, cycleEnd, setCycleWindow } = useRosterCtx();
  const isMaster = admin.role === "master_admin";
  const myBUs = isMaster ? BUS : adminBUs(admin);
  const scope = isMaster ? { mode: "admin" } : { mode: "bu", bu: myBUs[0] };
  const viewOnly = !!admin.testingOnly;
  const [buFilter, setBuFilter] = useState("ALL");
  const [rows, setRows] = useState(null);
  const [tab, setTab] = useState("dashboard");
  const [showExportMenu, setShowExportMenu] = useState(false);
  const [peopleSearch, setPeopleSearch] = useState("");
  const [viewCycle, setViewCycle] = useState(cycle);
  const [newCycleName, setNewCycleName] = useState("");
  const [showNewCycle, setShowNewCycle] = useState(false);
  const [editWindow, setEditWindow] = useState(false);
  const [startInput, setStartInput] = useState(cycleStart);
  const [endInput, setEndInput] = useState(cycleEnd);
  const isHistorical = viewCycle !== cycle;

  const scopeBUs = buFilter === "ALL" ? myBUs : [buFilter];

  useEffect(() => {
    (async () => {
      const list = roster.filter((u) => u.role === "staff" && u.needsEvaluation !== false && scopeBUs.includes(u.bu));
      const withData = await Promise.all(
        list.map(async (u) => {
          const d = await loadData(u.id, roster, viewCycle);
          const merged = { ...defaultData(), ...(d || {}) };
          const mgr = roster.find((m) => m.id === u.managerId);
          const peopleManager = hasReports(u.id, roster);
          return { user: u, manager: mgr, data: merged, progress: computeProgress(merged, peopleManager, false), status: journeyStatus(merged), overdue: isOverdue(merged, cycleEnd) };
        })
      );
      setRows(withData);
    })();
  }, [buFilter, roster, viewCycle]); // eslint-disable-line

  const exportSummary = () => {
    const data = (rows || []).map(({ user, manager, data: d, progress, status }) => ({
      Employee: `${user.firstName} ${user.lastName}`,
      "Employee ID": user.employeeId || "-",
      Username: user.username,
      BU: user.bu,
      Department: user.department || "-",
      "Job Grade": user.jobGrade || "-",
      Designation: user.designation || "-",
      Manager: manager ? `${manager.firstName} ${manager.lastName}` : "-",
      "Manager Employee ID": manager?.employeeId || "-",
      Status: JOURNEY_STATUS_META[status].label,
      "Progress %": progress,
      "Value-Add (Employee's self-view)": d.b4_snapshot_self || "-",
      "Value-Add (Manager's final view)": d.b4_snapshot_manager || "-",
      "Date Completed": status === "completed" ? new Date(d.manager_submitted_at).toLocaleDateString() : "-",
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [{ wch: 20 }, { wch: 14 }, { wch: 14 }, { wch: 10 }, { wch: 16 }, { wch: 10 }, { wch: 18 }, { wch: 20 }, { wch: 16 }, { wch: 22 }, { wch: 10 }, { wch: 24 }, { wch: 24 }, { wch: 14 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Summary");
    XLSX.writeFile(wb, "gva_summary.xlsx");
    appendLog(`${admin.firstName} ${admin.lastName} (${ROLE_LABEL[admin.role]})`, "Export", `High-level summary — ${buFilter === "ALL" ? myBUs.join(", ") : buFilter} — ${data.length} records`);
  };

  const exportFullAnswers = () => {
    const data = (rows || []).map(({ user, manager, data: d }) => ({
      Employee: `${user.firstName} ${user.lastName}`,
      "Employee ID": user.employeeId || "-",
      BU: user.bu,
      Manager: manager ? `${manager.firstName} ${manager.lastName}` : "-",
      "A1 Self-Cultivation — Self": d.a1_self || "",
      "A1 Self-Cultivation — Manager": d.a1_manager || "",
      "A2 Roles & Relationships — Self": d.a2_self || "",
      "A2 Roles & Relationships — Manager": d.a2_manager || "",
      "A3 Stewardship — Self": d.a3_self || "",
      "A3 Stewardship — Manager": d.a3_manager || "",
      "A4 Consciousness — Stage": d.a4_stage || "",
      "A4 Consciousness — Self": d.a4_self || "",
      "A4 Consciousness — Manager": d.a4_manager || "",
      "A5 Mobility": d.a5_mobility || "",
      "A5 Dimension Focus": d.a5_dimension || "",
      "A5 Growth Edge": d.a5_edge || "",
      "A5 Experience 1-3": [d.a5_exp1, d.a5_exp2, d.a5_exp3].filter(Boolean).join(" | "),
      "A5 Learning 1-3": [d.a5_learn1, d.a5_learn2, d.a5_learn3].filter(Boolean).join(" | "),
      "A5 Support Needed": d.a5_support || "",
      "A5 Manager Reflection": d.a5_manager || "",
      "B1 Learning — Self": d.b1_learning_self || "",
      "B1 Learning — Manager": d.b1_learning_manager || "",
      "B1 Integrity — Self": d.b1_integrity_self || "",
      "B1 Integrity — Manager": d.b1_integrity_manager || "",
      "B1 Coaching — Self": d.b1_coaching_self || "",
      "B1 Coaching — Manager": d.b1_coaching_manager || "",
      "B1 Connection — Self": d.b1_connection_self || "",
      "B1 Connection — Manager": d.b1_connection_manager || "",
      "B2 Collaboration — Self": d.b2_collab_self || "",
      "B2 Collaboration — Manager": d.b2_collab_manager || "",
      "B2 Accountability — Self": d.b2_steward_self || "",
      "B2 Accountability — Manager": d.b2_steward_manager || "",
      "B2 Innovation — Self": d.b2_innovation_self || "",
      "B2 Innovation — Manager": d.b2_innovation_manager || "",
      "B3 OKRs": [1, 2, 3, 4, 5].map((i) => d[`b3_okr${i}`]).filter(Boolean).join(" | "),
      "B3 OKR Progress": [1, 2, 3, 4, 5].map((i) => d[`b3_progress${i}`]).filter(Boolean).join(" | "),
      "B3 Delivering Results — Self": d.b3_results_self || "",
      "B3 Delivering Results — Manager": d.b3_results_manager || "",
      "B4 Self — Other reflections": d.b4_self || "",
      "B4 Manager — Other feedback": d.b4_manager || "",
      "Snapshot (Self)": d.b4_snapshot_self || "",
      "Snapshot (Manager, Final)": d.b4_snapshot_manager || "",
      "Committed by Employee": d.b4_commit_emp ? `Yes (${d.b4_commit_emp_date || "-"})` : "No",
      "Committed by Manager": d.b4_commit_mgr ? `Yes (${d.b4_commit_mgr_date || "-"})` : "No",
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Full Answers");
    XLSX.writeFile(wb, "gva_full_answers.xlsx");
    appendLog(`${admin.firstName} ${admin.lastName} (${ROLE_LABEL[admin.role]})`, "Export", `Full answers (Track A + B) — ${buFilter === "ALL" ? myBUs.join(", ") : buFilter} — ${data.length} records`);
  };

  // ---- aggregates for the dashboard ----
  const total = rows?.length || 0;
  const avgProgress = total ? Math.round(rows.reduce((s, r) => s + r.progress, 0) / total) : 0;
  const byBU = myBUs
    .filter((b) => scopeBUs.includes(b))
    .map((b) => ({ bu: b, count: (rows || []).filter((r) => r.user.bu === b).length }));
  const statusCounts = Object.keys(JOURNEY_STATUS_META).map((k) => ({
    name: JOURNEY_STATUS_META[k].label,
    key: k,
    value: (rows || []).filter((r) => r.status === k).length,
  }));
  const completedCount = statusCounts.find((s) => s.key === "completed")?.value || 0;
  const overdueCount = (rows || []).filter((r) => r.overdue).length;
  const needsConfirmationCount = statusCounts.find((s) => s.key === "drafted_pending_confirmation")?.value || 0;
  const snapshotCounts = ["Flourishing", "Generative", "Sustaining", "Depleting"].map((k) => ({
    name: k,
    value: (rows || []).filter((r) => r.data.b4_snapshot_manager === k).length,
  }));
  const rated = snapshotCounts.reduce((s, x) => s + x.value, 0);
  const snapshotByBU = isMaster
    ? byBU.map(({ bu }) => {
        const entry = { bu };
        ["Flourishing", "Generative", "Sustaining", "Depleting"].forEach((k) => {
          entry[k] = (rows || []).filter((r) => r.user.bu === bu && r.data.b4_snapshot_manager === k).length;
        });
        return entry;
      })
    : [];

  const managerGroups = useMemo(() => {
    if (!rows) return [];
    const managers = roster.filter((m) => m.role === "staff" && myBUs.includes(m.bu) && hasReports(m.id, roster));
    return managers
      .map((m) => {
        const team = rows.filter((r) => r.user.managerId === m.id);
        if (!team.length) return null;
        const avg = Math.round(team.reduce((s, r) => s + r.progress, 0) / team.length);
        const snaps = ["Flourishing", "Generative", "Sustaining", "Depleting"].map((k) => team.filter((r) => r.data.b4_snapshot_manager === k).length);
        return { manager: m, teamSize: team.length, avg, snaps };
      })
      .filter(Boolean);
  }, [rows, roster]); // eslint-disable-line

  const tabs = [{ id: "dashboard", label: t("Dashboard", "แดชบอร์ด", lang), icon: LayoutDashboard }];
  tabs.push({ id: "people", label: t("People", "บุคลากร", lang), icon: ClipboardList });
  if (isMaster) {
    tabs.push({ id: "logs", label: t("Activity Log", "ประวัติการใช้งาน", lang), icon: FileClock });
  }

  return (
    <div className="min-h-screen" style={{ background: BRAND.bgGradient, fontFamily: FONT_STACK }}>
      <style>{`${FONT_IMPORT}`}</style>
      <TopBar user={admin} onLogout={onLogout} onSwitchUser={onSwitchUser} subtitle="Admin Dashboard" scope={scope} />

      <div className="max-w-4xl mx-auto px-4 sm:px-6 pt-4">
        <div className="flex gap-1.5 mb-5 overflow-x-auto">
          {tabs.map((t) => {
            const Icon = t.icon;
            return (
              <button
                key={t.id}
                onClick={() => setTab(t.id)}
                className="inline-flex items-center gap-1.5 text-xs font-medium px-3 py-1.5 rounded-full shrink-0"
                style={tab === t.id ? { backgroundColor: BRAND.deep, color: "white" } : { backgroundColor: "white", color: BRAND.gray, border: `1px solid ${BRAND.line}` }}
              >
                <Icon className="w-3.5 h-3.5" /> {t.label}
              </button>
            );
          })}
        </div>
      </div>

      {tab === "dashboard" && (
        <div className="max-w-4xl mx-auto px-4 sm:px-6 pb-8">
          {isMaster && (
            <div className="flex flex-wrap items-center gap-2 mb-4">
              <span className="text-xs font-medium px-2.5 py-1 rounded-full inline-flex items-center gap-1.5" style={{ backgroundColor: "#F0F8F8", color: BRAND.primary }}>
                <Clock className="w-3 h-3" /> {t("Cycle", "รอบการประเมิน", lang)}: {viewCycle}{isHistorical ? ` (${t("archived", "เก็บถาวร", lang)})` : ""}
              </span>
              {cyclesList.length > 1 && (
                <select
                  value={viewCycle}
                  onChange={(e) => setViewCycle(e.target.value)}
                  className="text-xs rounded-lg border px-2 py-1 focus:outline-none"
                  style={{ borderColor: BRAND.line }}
                >
                  {cyclesList.map((c) => (
                    <option key={c} value={c}>{c}{c === cycle ? ` (${t("current", "ปัจจุบัน", lang)})` : ""}</option>
                  ))}
                </select>
              )}
              {!viewOnly && !showNewCycle && (
                <button onClick={() => setShowNewCycle(true)} className="text-xs font-medium px-2.5 py-1 rounded-full border" style={{ borderColor: BRAND.line, color: BRAND.deep }}>
                  {t("Start New Cycle", "เริ่มรอบใหม่", lang)}
                </button>
              )}
              {showNewCycle && (
                <div className="flex items-center gap-1.5">
                  <TextInput value={newCycleName} onChange={setNewCycleName} placeholder="e.g. 2027" />
                  <button
                    onClick={async () => {
                      if (!newCycleName.trim()) return;
                      await startNewCycle(newCycleName.trim());
                      setViewCycle(newCycleName.trim());
                      setNewCycleName("");
                      setShowNewCycle(false);
                    }}
                    className="text-xs font-medium px-3 py-1.5 rounded-lg text-white"
                    style={{ backgroundColor: BRAND.primary }}
                  >
                    {t("Create & switch", "สร้างและสลับ", lang)}
                  </button>
                  <button onClick={() => { setShowNewCycle(false); setNewCycleName(""); }} className="text-xs text-slate-400">{t("Cancel", "ยกเลิก", lang)}</button>
                </div>
              )}
              <span className="text-xs font-medium px-2.5 py-1 rounded-full inline-flex items-center gap-1.5" style={{ backgroundColor: "#F6FBFA", color: BRAND.gray, border: `1px solid ${BRAND.line}` }}>
                <Clock className="w-3 h-3" />
                {cycleEnd
                  ? t(
                      `Deadline: ${new Date(cycleEnd).toLocaleDateString()}${cycleStart ? ` (from ${new Date(cycleStart).toLocaleDateString()})` : ""}`,
                      `กำหนดเสร็จ: ${new Date(cycleEnd).toLocaleDateString("th-TH")}${cycleStart ? ` (เริ่ม ${new Date(cycleStart).toLocaleDateString("th-TH")})` : ""}`,
                      lang
                    )
                  : t("No deadline set yet", "ยังไม่ได้กำหนดวันสิ้นสุด", lang)}
              </span>
              {!viewOnly && !editWindow && (
                <button onClick={() => { setStartInput(cycleStart); setEndInput(cycleEnd); setEditWindow(true); }} className="text-xs font-medium px-2.5 py-1 rounded-full border" style={{ borderColor: BRAND.line, color: BRAND.deep }}>
                  {cycleEnd ? t("Change", "เปลี่ยน", lang) : t("Set deadline", "กำหนดวันสิ้นสุด", lang)}
                </button>
              )}
              {editWindow && (
                <div className="flex items-center gap-1.5">
                  <div>
                    <label className="text-[10px] text-slate-400 block">{t("Start", "วันเริ่ม", lang)}</label>
                    <input
                      type="date"
                      value={startInput}
                      onChange={(e) => setStartInput(e.target.value)}
                      className="text-xs rounded-lg border px-2 py-1.5 focus:outline-none"
                      style={{ borderColor: BRAND.line }}
                    />
                  </div>
                  <div>
                    <label className="text-[10px] text-slate-400 block">{t("End (deadline)", "วันสิ้นสุด", lang)}</label>
                    <input
                      type="date"
                      value={endInput}
                      onChange={(e) => setEndInput(e.target.value)}
                      className="text-xs rounded-lg border px-2 py-1.5 focus:outline-none"
                      style={{ borderColor: BRAND.line }}
                    />
                  </div>
                  <button
                    onClick={async () => { await setCycleWindow(startInput, endInput); setEditWindow(false); }}
                    className="text-xs font-medium px-3 py-1.5 rounded-lg text-white self-end"
                    style={{ backgroundColor: BRAND.primary }}
                  >
                    {t("Save", "บันทึก", lang)}
                  </button>
                  <button onClick={() => setEditWindow(false)} className="text-xs text-slate-400 self-end pb-1.5">{t("Cancel", "ยกเลิก", lang)}</button>
                </div>
              )}
            </div>
          )}
          {isHistorical && (
            <div className="rounded-xl px-4 py-3 mb-4 flex items-start gap-2.5" style={{ backgroundColor: "#F0F8F8", border: `1px solid ${BRAND.line}` }}>
              <Clock className="w-4 h-4 mt-0.5 shrink-0" style={{ color: BRAND.primary }} />
              <p className="text-sm" style={{ color: BRAND.deep }}>
                {t(
                  <>Viewing the archived <span className="font-medium">{viewCycle}</span> cycle — read-only summary and export. Switch back to <span className="font-medium">{cycle}</span> above to resume live work.</>,
                  <>กำลังดูรอบที่เก็บถาวร <span className="font-medium">{viewCycle}</span> — สรุปและส่งออกได้อย่างเดียว สลับกลับไปที่ <span className="font-medium">{cycle}</span> ด้านบนเพื่อทำงานต่อ</>,
                  lang
                )}
              </p>
            </div>
          )}
          <div className="flex flex-col sm:flex-row sm:items-end sm:justify-between gap-3 mb-5">
            <div>
              <h1 className="text-xl font-semibold mb-1 flex items-center gap-2" style={{ color: BRAND.deep }}>
                <LayoutDashboard className="w-5 h-5" style={{ color: BRAND.mint }} /> {t("People & Growth Overview", "ภาพรวมบุคลากรและการเติบโต", lang)}
              </h1>
              <p className="text-xs text-slate-400">{myBUs.length > 1 ? `${myBUs.join(", ")}` : myBUs[0]}</p>
            </div>
            <div className="flex items-center gap-2 flex-wrap">
              {myBUs.length > 1 && (
                <select
                  value={buFilter}
                  onChange={(e) => setBuFilter(e.target.value)}
                  className="rounded-lg border text-sm px-3 py-2 focus:outline-none"
                  style={{ borderColor: BRAND.line }}
                >
                  <option value="ALL">All ({myBUs.length} BUs)</option>
                  {myBUs.map((b) => (
                    <option key={b} value={b}>{b}</option>
                  ))}
                </select>
              )}
              <button
                onClick={() => setShowExportMenu((v) => !v)}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 text-sm font-medium rounded-xl text-white"
                style={{ backgroundColor: BRAND.primary }}
              >
                <FileSpreadsheet className="w-4 h-4" /> {t("Export Excel", "ส่งออก Excel", lang)} <ChevronDown className="w-3.5 h-3.5" />
              </button>
            </div>
          </div>

          {showExportMenu && (
            <div className="rounded-2xl p-1.5 mb-5 flex flex-col sm:flex-row gap-1.5" style={{ border: `1px solid ${BRAND.line}`, backgroundColor: BRAND.card }}>
              <button
                onClick={() => { exportSummary(); setShowExportMenu(false); }}
                className="flex-1 text-left px-3.5 py-2.5 rounded-xl hover:bg-slate-50 transition"
              >
                <p className="text-sm font-medium flex items-center gap-1.5" style={{ color: BRAND.deep }}><Download className="w-3.5 h-3.5" /> {t("High-level Summary", "สรุปภาพรวม", lang)}</p>
                <p className="text-xs text-slate-400 mt-0.5">{t("Employee details, status, Value-Add mix (self & manager), date completed.", "รายละเอียดพนักงาน สถานะ สัดส่วนการสร้างมูลค่าเพิ่ม (ตนเองและหัวหน้างาน) วันที่เสร็จสมบูรณ์", lang)}</p>
              </button>
              <button
                onClick={() => { exportFullAnswers(); setShowExportMenu(false); }}
                className="flex-1 text-left px-3.5 py-2.5 rounded-xl hover:bg-slate-50 transition"
              >
                <p className="text-sm font-medium flex items-center gap-1.5" style={{ color: BRAND.deep }}><Sprout className="w-3.5 h-3.5" /> {t("Full Answers (Track A + B)", "คำตอบทั้งหมด (ส่วน A และ B)", lang)}</p>
                <p className="text-xs text-slate-400 mt-0.5">{t("Every written reflection, both tracks — for deep-dive analysis.", "การทบทวนที่เขียนไว้ทั้งหมด ทั้งสองส่วน — สำหรับการวิเคราะห์เชิงลึก", lang)}</p>
              </button>
            </div>
          )}

          {/* stat cards */}
          <div className="grid grid-cols-3 sm:grid-cols-6 gap-3 mb-4">
            <StatCard label={t("People in scope", "จำนวนบุคลากรในขอบเขต", lang)} value={total} icon={UsersIcon} color={BRAND.primary} />
            <StatCard label={t("Avg progress", "ความคืบหน้าเฉลี่ย", lang)} value={`${avgProgress}%`} icon={Clock} color={BRAND.teal} />
            <StatCard label={t("Completed", "เสร็จสมบูรณ์", lang)} value={completedCount} sub={total ? `${Math.round((completedCount / total) * 100)}%` : undefined} icon={CheckCircle2} color={BRAND.mint} />
            <StatCard
              label={t("Needs confirmation", "รอการยืนยัน", lang)}
              value={needsConfirmationCount}
              icon={RefreshCcw}
              color={BRAND.primary}
              info={t(
                "A delegate has drafted the manager's feedback for this person, but the direct manager hasn't reviewed and confirmed it yet — it isn't official until they do.",
                "ผู้ได้รับมอบหมายได้ร่างความเห็นของหัวหน้างานให้คนนี้แล้ว แต่หัวหน้างานตัวจริงยังไม่ได้ตรวจสอบและยืนยัน จึงยังไม่ถือเป็นทางการ",
                lang
              )}
            />
            <StatCard
              label={t("Overdue", "เกินกำหนด", lang)}
              value={overdueCount}
              icon={AlertCircle}
              color={BRAND.red}
              info={t(
                cycleEnd
                  ? `This person hasn't completed their review (both self-assessment and manager review) by this cycle's deadline of ${new Date(cycleEnd).toLocaleDateString()}. Master Admin can change the deadline below.`
                  : "No cycle deadline is set yet, so nothing is flagged overdue. Master Admin can set one below.",
                cycleEnd
                  ? `บุคคลนี้ยังไม่เสร็จสิ้นการประเมิน (ทั้งการประเมินตนเองและการตรวจสอบของหัวหน้างาน) ภายในกำหนดของรอบนี้คือ ${new Date(cycleEnd).toLocaleDateString("th-TH")} Master Admin สามารถเปลี่ยนกำหนดได้ด้านล่าง`
                  : "ยังไม่ได้กำหนดวันสิ้นสุดของรอบนี้ จึงยังไม่มีการแจ้งเตือนเกินกำหนด Master Admin สามารถกำหนดได้ด้านล่าง",
                lang
              )}
            />
            <StatCard label={t("Manager groups", "กลุ่มหัวหน้างาน", lang)} value={managerGroups.length} icon={Building2} color={BRAND.gray} />
          </div>

          {/* charts */}
          <div className={`grid gap-4 mb-4 ${isMaster ? "sm:grid-cols-3" : "sm:grid-cols-2"}`}>
            {isMaster && (
              <div className="rounded-2xl p-4" style={{ border: `1px solid ${BRAND.line}`, backgroundColor: BRAND.card }}>
                <p className="text-xs font-semibold uppercase tracking-wide mb-2 flex items-center gap-1.5" style={{ color: BRAND.gray }}>
                  <Building2 className="w-3.5 h-3.5" /> {t("Headcount by Company", "จำนวนบุคลากรตามบริษัท", lang)}
                </p>
                <ResponsiveContainer width="100%" height={170}>
                  <BarChart data={byBU} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={BRAND.line} vertical={false} />
                    <XAxis dataKey="bu" tick={{ fontSize: 10, fill: BRAND.gray }} axisLine={{ stroke: BRAND.line }} tickLine={false} />
                    <YAxis allowDecimals={false} tick={{ fontSize: 11, fill: BRAND.gray }} axisLine={false} tickLine={false} />
                    <Tooltip contentStyle={{ fontSize: 12, borderRadius: 8, borderColor: BRAND.line }} />
                    <Bar dataKey="count" fill={BRAND.mint} radius={[6, 6, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            )}
            <DonutWithLegend
              title={t(`Completion status${!isMaster ? ` (${myBUs[0]})` : ""}`, `สถานะความสำเร็จ${!isMaster ? ` (${myBUs[0]})` : ""}`, lang)}
              icon={CheckCircle2}
              data={statusCounts}
              colorMap={Object.fromEntries(Object.entries(JOURNEY_STATUS_META).map(([k, v]) => [k, v.color]))}
              total={total}
            />
            <DonutWithLegend
              title={t("Value-Add Snapshot mix", "สัดส่วนภาพรวมการสร้างมูลค่าเพิ่ม", lang)}
              icon={Sprout}
              data={snapshotCounts}
              colorMap={SNAPSHOT_HEX}
              total={rated}
            />
          </div>

          {isMaster && snapshotByBU.some((r) => r.Flourishing + r.Generative + r.Sustaining + r.Depleting > 0) && (
            <div className="rounded-2xl p-4 mb-4" style={{ border: `1px solid ${BRAND.line}`, backgroundColor: BRAND.card }}>
              <p className="text-xs font-semibold uppercase tracking-wide mb-2 flex items-center gap-1.5" style={{ color: BRAND.gray }}>
                <Sprout className="w-3.5 h-3.5" /> {t("Value-Add Snapshot mix by Company", "สัดส่วนการสร้างมูลค่าเพิ่มตามบริษัท", lang)}
              </p>
              <ResponsiveContainer width="100%" height={190}>
                <BarChart data={snapshotByBU} margin={{ top: 4, right: 8, left: -18, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={BRAND.line} vertical={false} />
                  <XAxis dataKey="bu" tick={{ fontSize: 11, fill: BRAND.gray }} axisLine={{ stroke: BRAND.line }} tickLine={false} />
                  <YAxis allowDecimals={false} tick={{ fontSize: 11, fill: BRAND.gray }} axisLine={false} tickLine={false} />
                  <Tooltip contentStyle={{ fontSize: 12, borderRadius: 8, borderColor: BRAND.line }} />
                  <Legend wrapperStyle={{ fontSize: 11 }} />
                  <Bar dataKey="Flourishing" stackId="snap" fill={SNAPSHOT_HEX.Flourishing} />
                  <Bar dataKey="Generative" stackId="snap" fill={SNAPSHOT_HEX.Generative} />
                  <Bar dataKey="Sustaining" stackId="snap" fill={SNAPSHOT_HEX.Sustaining} />
                  <Bar dataKey="Depleting" stackId="snap" fill={SNAPSHOT_HEX.Depleting} radius={[6, 6, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          )}

          {/* by manager group */}
          <div className="rounded-2xl overflow-hidden mb-5" style={{ border: `1px solid ${BRAND.line}`, backgroundColor: BRAND.card }}>
            <p className="text-xs font-semibold uppercase tracking-wide px-5 pt-4 pb-2" style={{ color: BRAND.gray }}>{t("By manager group", "ตามกลุ่มหัวหน้างาน", lang)}</p>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left text-xs" style={{ color: BRAND.gray, backgroundColor: "#F6FBFA" }}>
                    <th className="px-5 py-2 font-medium">Manager</th>
                    <th className="px-3 py-2 font-medium">Company</th>
                    <th className="px-3 py-2 font-medium">Team size</th>
                    <th className="px-3 py-2 font-medium">Avg progress</th>
                    <th className="px-3 py-2 font-medium">Snapshot mix</th>
                  </tr>
                </thead>
                <tbody className="divide-y" style={{ borderColor: BRAND.line }}>
                  {managerGroups.map(({ manager, teamSize, avg, snaps }) => (
                    <tr key={manager.id}>
                      <td className="px-5 py-2.5 font-medium text-slate-800 whitespace-nowrap">{manager.firstName} {manager.lastName}</td>
                      <td className="px-3 py-2.5 text-slate-500">{manager.bu}</td>
                      <td className="px-3 py-2.5 text-slate-500">{teamSize}</td>
                      <td className="px-3 py-2.5 text-slate-500">{avg}%</td>
                      <td className="px-3 py-2.5">
                        <div className="flex h-2.5 w-32 rounded-full overflow-hidden" style={{ backgroundColor: "#F1F5F4" }}>
                          {snaps.map((n, i) =>
                            n > 0 ? (
                              <div key={i} style={{ width: `${(n / teamSize) * 100}%`, backgroundColor: Object.values(SNAPSHOT_HEX)[i] }} />
                            ) : null
                          )}
                        </div>
                      </td>
                    </tr>
                  ))}
                  {!managerGroups.length && (
                    <tr><td className="px-5 py-4 text-sm text-slate-400" colSpan={5}>No manager groups in scope yet.</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>

          {/* full detail table */}
          <div className="rounded-2xl overflow-hidden" style={{ border: `1px solid ${BRAND.line}`, backgroundColor: BRAND.card }}>
            <div className="flex items-center justify-between gap-2 px-5 pt-4 pb-2 flex-wrap">
              <p className="text-xs font-semibold uppercase tracking-wide" style={{ color: BRAND.gray }}>{t("All people in scope", "บุคลากรทั้งหมดในขอบเขต", lang)}</p>
              <div className="relative">
                <Search className="w-3.5 h-3.5 absolute left-2.5 top-2.5 text-slate-300" />
                <input
                  value={peopleSearch}
                  onChange={(e) => setPeopleSearch(e.target.value)}
                  placeholder={t("Search by name…", "ค้นหาตามชื่อ…", lang)}
                  className="text-xs rounded-lg border pl-7 pr-3 py-1.5 focus:outline-none"
                  style={{ borderColor: BRAND.line }}
                />
              </div>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left text-xs" style={{ color: BRAND.gray, backgroundColor: "#F6FBFA" }}>
                    <th className="px-5 py-3 font-medium">{t("Employee", "พนักงาน", lang)}</th>
                    <th className="px-3 py-3 font-medium">{t("Company", "บริษัท", lang)}</th>
                    <th className="px-3 py-3 font-medium">{t("Manager", "หัวหน้างาน", lang)}</th>
                    <th className="px-3 py-3 font-medium">{t("Status", UI_TH.status, lang)}</th>
                    <th className="px-3 py-3 font-medium">{t("Progress", "ความคืบหน้า", lang)}</th>
                    <th className="px-3 py-3 font-medium">{t("Employee's Snapshot", "ภาพรวมของพนักงาน", lang)}</th>
                    <th className="px-3 py-3 font-medium">{t("Manager's Snapshot", "ภาพรวมของหัวหน้างาน", lang)} <span className="font-normal text-slate-400">({t("final", "สุดท้าย", lang)})</span></th>
                    <th className="px-5 py-3 font-medium"></th>
                  </tr>
                </thead>
                <tbody className="divide-y" style={{ borderColor: BRAND.line }}>
                  {rows === null && (
                    <tr><td className="px-5 py-6 text-sm text-slate-400" colSpan={8}>Loading…</td></tr>
                  )}
                  {rows?.filter(({ user }) => `${user.firstName} ${user.lastName}`.toLowerCase().includes(peopleSearch.toLowerCase())).map(({ user, manager, data: d, progress, status, overdue }) => {
                    const scSelf = d.b4_snapshot_self ? snapshotColor(d.b4_snapshot_self) : null;
                    const scMgr = d.b4_snapshot_manager ? snapshotColor(d.b4_snapshot_manager) : null;
                    const discrepancy = d.b4_snapshot_self && d.b4_snapshot_manager && d.b4_snapshot_self !== d.b4_snapshot_manager;
                    const sMeta = JOURNEY_STATUS_META[status];
                    const SIcon = sMeta.icon;
                    const rowLocked = viewOnly || isHistorical;
                    return (
                      <tr
                        key={user.id}
                        onClick={rowLocked ? undefined : () => onOpenEmployee(user.id)}
                        className={rowLocked ? "" : "cursor-pointer hover:bg-slate-50 transition"}
                      >
                        <td className="px-5 py-3 font-medium text-slate-800 whitespace-nowrap">
                          <span className="flex items-center gap-2">
                            <Avatar name={`${user.firstName} ${user.lastName}`} size={24} />
                            {user.firstName} {user.lastName}
                          </span>
                        </td>
                        <td className="px-3 py-3 text-slate-500">{user.bu}</td>
                        <td className="px-3 py-3 text-slate-500 whitespace-nowrap">{manager ? `${manager.firstName} ${manager.lastName}` : "-"}</td>
                        <td className="px-3 py-3">
                          <span className="inline-flex items-center gap-1 text-xs font-medium px-2 py-0.5 rounded-full whitespace-nowrap" style={{ backgroundColor: `${sMeta.color}1A`, color: sMeta.color }}>
                            <SIcon className="w-3 h-3" /> {t(sMeta.label, STATUS_LABEL_TH[status], lang)}
                          </span>
                          {overdue && (
                            <span className="ml-1 inline-flex items-center gap-1 text-xs font-medium px-2 py-0.5 rounded-full whitespace-nowrap" style={{ backgroundColor: "#FFF4F3", color: BRAND.red }}>
                              <AlertCircle className="w-3 h-3" /> {t("Overdue", "เกินกำหนด", lang)}
                            </span>
                          )}
                        </td>
                        <td className="px-3 py-3">
                          <div className="flex items-center gap-2">
                            <div className="w-16 h-1.5 rounded-full overflow-hidden" style={{ backgroundColor: "#DCEEEC" }}>
                              <div className="h-full rounded-full" style={{ width: `${progress}%`, backgroundColor: BRAND.mint }} />
                            </div>
                            <span className="text-xs font-semibold" style={{ color: BRAND.primary }}>{progress}%</span>
                          </div>
                        </td>
                        <td className="px-3 py-3">
                          {scSelf ? (
                            <span className="text-xs font-medium px-2 py-0.5 rounded-full" style={{ backgroundColor: scSelf.bg, color: scSelf.text, border: `1px solid ${scSelf.border}` }}>
                              {t(d.b4_snapshot_self, SNAPSHOT_TH[d.b4_snapshot_self]?.label, lang)}
                            </span>
                          ) : (
                            <span className="text-xs text-slate-300">—</span>
                          )}
                        </td>
                        <td className="px-3 py-3">
                          {scMgr ? (
                            <span className="text-xs font-medium px-2 py-0.5 rounded-full" style={{ backgroundColor: scMgr.bg, color: scMgr.text, border: `1px solid ${scMgr.border}` }}>
                              {t(d.b4_snapshot_manager, SNAPSHOT_TH[d.b4_snapshot_manager]?.label, lang)}
                            </span>
                          ) : (
                            <span className="text-xs text-slate-300">—</span>
                          )}
                        </td>
                        <td className="px-5 py-3 text-right">
                          {rowLocked ? (
                            <span className="text-xs text-slate-300">{isHistorical ? t("Archived", "เก็บถาวร", lang) : t("View only", "ดูอย่างเดียว", lang)}</span>
                          ) : (
                            <span className="inline-flex items-center gap-1 text-xs font-medium" style={{ color: BRAND.deep }}>
                              {t("View Track A & B", "ดูส่วน A และ B", lang)} <ChevronRight className="w-3 h-3" />
                            </span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {tab === "people" && <PeopleManager admin={admin} scopeBUs={isMaster ? null : myBUs} viewOnly={viewOnly} />}
      {tab === "logs" && isMaster && <ExportLogView />}
    </div>
  );
}



/* ========================================================================
   PEOPLE MANAGEMENT — Master Admin only. Import/export/edit the roster
   that everything else in the app (login, BU scope, org chart) reads from.
========================================================================= */

function PersonEditPanel({ initial, roster, onSave, onCancel, restrictBUs, restrictStaffOnly }) {
  const isNew = !initial;
  const [form, setForm] = useState(() =>
    initial
      ? { ...initial, password: undefined } // never load the stored hash into the form
      : { id: `p-${Date.now()}`, firstName: "", lastName: "", username: "", role: "staff", bu: restrictBUs ? restrictBUs[0] : BUS[0], buList: null, department: "", jobGrade: "", designation: "", managerId: null, employeeId: "", needsEvaluation: true, functionalManagerId: null }
  );
  const [newPassword, setNewPassword] = useState("");
  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  const buChoices = restrictBUs || BUS;
  const managerOptions = roster
    .filter((p) => p.role === "staff" && p.id !== form.id && p.bu === form.bu)
    .sort((a, b) => `${a.lastName}${a.firstName}`.localeCompare(`${b.lastName}${b.firstName}`));
  // Cross-company on purpose — a functional manager is explicitly someone
  // outside the person's own reporting line, often at TPC Group itself.
  const functionalManagerOptions = roster
    .filter((p) => p.role === "staff" && p.id !== form.id)
    .sort((a, b) => (a.bu || "").localeCompare(b.bu || "") || `${a.lastName}${a.firstName}`.localeCompare(`${b.lastName}${b.firstName}`));
  const toggleBu = (b) => {
    const list = form.buList || [];
    set("buList", list.includes(b) ? list.filter((x) => x !== b) : [...list, b]);
  };
  const selectCls = "w-full rounded-xl border px-3.5 py-2.5 text-sm focus:outline-none";

  return (
    <div className="rounded-2xl p-5 mb-4" style={{ border: `1px solid ${BRAND.teal}`, backgroundColor: "#F6FBFA" }}>
      <p className="text-sm font-semibold mb-3" style={{ color: BRAND.deep }}>
        {isNew ? "Add Person" : `Edit ${initial.firstName} ${initial.lastName}`}
      </p>
      <div className="grid grid-cols-1 min-[480px]:grid-cols-2 gap-3 mb-3">
        <Field label="First name"><TextInput value={form.firstName} onChange={(v) => set("firstName", v)} /></Field>
        <Field label="Last name"><TextInput value={form.lastName} onChange={(v) => set("lastName", v)} /></Field>
        <Field label="Employee Code"><TextInput value={form.employeeId} onChange={(v) => set("employeeId", v)} placeholder="e.g. TH08-25003" /></Field>
        <Field label="Username (for login)"><TextInput value={form.username} onChange={(v) => set("username", v)} /></Field>
        <Field label={isNew ? "Password" : "Set new password (leave blank to keep unchanged)"}>
          <TextInput value={newPassword} onChange={setNewPassword} placeholder={isNew ? "" : "•••••••• (unchanged)"} />
        </Field>
        {!restrictStaffOnly && (
          <Field label="Account type">
            <select value={form.role} onChange={(e) => set("role", e.target.value)} className={selectCls} style={{ borderColor: BRAND.line }}>
              <option value="staff">Staff</option>
              <option value="po_admin">P&amp;O Admin</option>
              <option value="master_admin">Master Admin</option>
            </select>
          </Field>
        )}
        {form.role !== "po_admin" && (
          <Field label="Company Name">
            <select value={form.bu || ""} onChange={(e) => set("bu", e.target.value || null)} className={selectCls} style={{ borderColor: BRAND.line }}>
              {!restrictBUs && <option value="">— none (Master Admin) —</option>}
              {buChoices.map((b) => <option key={b} value={b}>{b} — {BU_FULL_NAME[b]}</option>)}
            </select>
          </Field>
        )}
        {form.role === "staff" && (
          <>
            <Field label="Department"><TextInput value={form.department} onChange={(v) => set("department", v)} /></Field>
            <Field label="Job grade"><TextInput value={form.jobGrade} onChange={(v) => set("jobGrade", v)} /></Field>
            <Field label="Designation"><TextInput value={form.designation} onChange={(v) => set("designation", v)} /></Field>
            <Field label={`Direct manager${form.bu ? ` (${form.bu})` : ""}`}>
              <select value={form.managerId || ""} onChange={(e) => set("managerId", e.target.value || null)} className={selectCls} style={{ borderColor: BRAND.line }}>
                <option value="">— none (top of chain) —</option>
                {managerOptions.map((m) => <option key={m.id} value={m.id}>{m.firstName} {m.lastName}</option>)}
              </select>
              {!managerOptions.length && <p className="text-xs text-slate-400 mt-1">No other staff yet in this company to choose from.</p>}
            </Field>
            <Field
              label={
                <span className="inline-flex items-center gap-1.5">
                  Second reviewer (optional)
                  <InfoTip text="Someone from another company who can also give feedback — like a manager at TPC Group. Their input is a draft; the direct manager above still makes the final call." />
                </span>
              }
            >
              <select value={form.functionalManagerId || ""} onChange={(e) => set("functionalManagerId", e.target.value || null)} className={selectCls} style={{ borderColor: BRAND.line }}>
                <option value="">— none —</option>
                {functionalManagerOptions.map((m) => <option key={m.id} value={m.id}>{m.firstName} {m.lastName} — {m.bu}</option>)}
              </select>
            </Field>
          </>
        )}
      </div>

      {form.role === "staff" && (
        <label className="flex items-center gap-2.5 rounded-xl border p-3 mb-3 text-sm" style={{ borderColor: BRAND.line, backgroundColor: "#F6FBFA" }}>
          <input
            type="checkbox"
            checked={form.needsEvaluation !== false}
            onChange={(e) => set("needsEvaluation", e.target.checked)}
            className="w-4 h-4"
          />
          <span className="font-medium flex items-center gap-1.5" style={{ color: BRAND.deep }}>
            Does this person need to complete an evaluation?
            <InfoTip text="Turn this off for someone who only reviews others — like a senior manager or a helper who never fills out their own evaluation. They can still log in and manage their team." />
          </span>
        </label>
      )}

      {form.role === "po_admin" && (
        <div className="mb-3">
          <p className="text-xs font-medium mb-1.5" style={{ color: BRAND.deep }}>Companies covered (multi-entity P&amp;O)</p>
          <div className="flex flex-wrap gap-2">
            {BUS.map((b) => (
              <button
                key={b} type="button" onClick={() => toggleBu(b)}
                className="text-xs px-2.5 py-1 rounded-full border"
                style={(form.buList || []).includes(b) ? { backgroundColor: BRAND.primary, color: "white", borderColor: BRAND.primary } : { borderColor: BRAND.line, color: BRAND.gray }}
                title={BU_FULL_NAME[b]}
              >
                {b}
              </button>
            ))}
          </div>
        </div>
      )}

      <div className="flex justify-end gap-2 mt-3">
        <button onClick={onCancel} className="px-4 py-2 text-sm rounded-xl border" style={{ borderColor: BRAND.line, color: BRAND.deep }}>
          Cancel
        </button>
        <button
          onClick={() => onSave(form, newPassword)}
          disabled={!form.firstName || !form.lastName || !form.username || (isNew && !newPassword)}
          className="px-4 py-2 text-sm rounded-xl text-white disabled:opacity-40"
          style={{ backgroundColor: BRAND.primary }}
        >
          Save
        </button>
      </div>
    </div>
  );
}

function PeopleManager({ admin, scopeBUs, viewOnly }) {
  const { roster, setRoster, cycle } = useRosterCtx();
  const [editingId, setEditingId] = useState(null); // person id, "new", or null
  const [importMsg, setImportMsg] = useState("");
  const [deletingId, setDeletingId] = useState(null);
  const [deleteMsg, setDeleteMsg] = useState("");
  const [showDangerZone, setShowDangerZone] = useState(false);
  const [confirmWipeText, setConfirmWipeText] = useState("");
  const [wipeBusy, setWipeBusy] = useState(false);
  const [wipeMsg, setWipeMsg] = useState("");
  const fileRef = useRef(null);
  const isScoped = !!scopeBUs;
  const isMasterAdmin = !isScoped && !viewOnly;
  const actorLabel = `${admin.firstName} ${admin.lastName} (${ROLE_LABEL[admin.role]})`;
  const visibleRoster = isScoped ? roster.filter((p) => p.role === "staff" && scopeBUs.includes(p.bu)) : roster;

  const persist = async (next) => {
    setRoster(next);
    await saveRoster(next);
  };

  const handleDeletePerson = async (person) => {
    if (hasReports(person.id, roster)) {
      setDeleteMsg(`Can't delete ${person.firstName} ${person.lastName} — they still manage direct reports. Reassign those first (edit each report's Direct Manager).`);
      setDeletingId(null);
      return;
    }
    const next = roster
      .filter((p) => p.id !== person.id)
      .map((p) => (p.reviewDelegateId === person.id ? { ...p, reviewDelegateId: null } : p));
    await persist(next);
    await appendLog(actorLabel, "Deleted Person", `Removed ${person.firstName} ${person.lastName} (${person.username}) from the roster.`);
    setDeleteMsg(`${person.firstName} ${person.lastName} was removed from the roster.`);
    setDeletingId(null);
  };

  const handleWipeJourneyData = async () => {
    setWipeBusy(true);
    const count = await clearAllJourneyData();
    await clearActivityLog();
    await appendLog(actorLabel, "UAT Reset", `Cleared all journey responses (${count} records) and the activity log ahead of go-live. Roster was not affected.`);
    setWipeBusy(false);
    setConfirmWipeText("");
    setWipeMsg(`Done — cleared ${count} journey record${count === 1 ? "" : "s"} across every company and cycle. The roster (people, companies, reporting lines) was left untouched.`);
  };

  const handleSaveEdit = async (person, newPassword) => {
    const existingRecord = roster.find((p) => p.id === person.id);
    // Never write a plaintext password into the roster blob — keep the
    // existing stored hash (or a placeholder for a brand-new person; the
    // /api/set-passwords call right below fills in the real hash).
    const personWithPassword = { ...person, password: existingRecord ? existingRecord.password : "unset" };
    const safePerson = isScoped ? { ...personWithPassword, role: "staff", bu: scopeBUs.includes(personWithPassword.bu) ? personWithPassword.bu : scopeBUs[0], buList: null } : personWithPassword;
    const exists = !!existingRecord;
    const next = exists ? roster.map((p) => (p.id === safePerson.id ? safePerson : p)) : [...roster, safePerson];
    await persist(next);
    if (newPassword) {
      try {
        await fetch("/api/set-passwords", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ updates: [{ id: safePerson.id, newPassword }] }),
        });
      } catch (e) {
        console.error("set-passwords failed", e);
      }
    }
    await appendLog(actorLabel, exists ? "Edit" : "Add", `${safePerson.firstName} ${safePerson.lastName} (${safePerson.username})`);
    setEditingId(null);
  };

  const handleImportFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      let added = 0, updated = 0, skipped = 0, rejected = 0, delegateLocked = 0;
      let next = [...roster];
      const pendingPasswords = []; // [{id, newPassword}] — hashed server-side, never written to the roster blob directly
      for (const row of json) {
        const username = String(row["Username"] || row["username"] || "").trim();
        const employeeCode = String(row["Employee Code"] || row["EmployeeCode"] || row["Employee ID"] || "").trim();
        if (!username) { skipped++; continue; } // a login still requires a username
        let idx = -1;
        if (employeeCode) idx = next.findIndex((p) => (p.employeeId || "").toLowerCase() === employeeCode.toLowerCase());
        if (idx < 0) idx = next.findIndex((p) => p.username.toLowerCase() === username.toLowerCase());
        const managerCode = String(row["Manager Code"] || "").trim();
        const managerUsername = String(row["Manager Username"] || row["Manager"] || "").trim();
        const managerPerson = managerCode
          ? next.find((p) => (p.employeeId || "").toLowerCase() === managerCode.toLowerCase())
          : managerUsername
          ? next.find((p) => p.username.toLowerCase() === managerUsername.toLowerCase())
          : null;
        const fnMgrCode = String(row["Functional Manager Code"] || "").trim();
        const fnMgrPerson = fnMgrCode ? next.find((p) => (p.employeeId || "").toLowerCase() === fnMgrCode.toLowerCase()) : null;
        const delegateCode = String(row["Delegate Code"] || "").trim();
        const delegatePerson = delegateCode ? next.find((p) => (p.employeeId || "").toLowerCase() === delegateCode.toLowerCase()) : null;
        const buRaw = String(row["Company Code"] || row["Company Name"] || row["BU"] || row["Business Unit"] || "").trim();
        // Accept either the short code (UTSE) or the full legal name (reverse-looked-up).
        const bu = buRaw ? (BUS.includes(buRaw) ? buRaw : Object.keys(BU_FULL_NAME).find((code) => BU_FULL_NAME[code] === buRaw) || buRaw) : null;
        const roleRaw = String(row["Role"] || row["Type"] || "staff").trim().toLowerCase();
        let role = ["po_admin", "master_admin"].includes(roleRaw) ? roleRaw : "staff";
        const buListRaw = String(row["Company List"] || row["BU List"] || "").trim();
        let buList = buListRaw ? buListRaw.split(/[,;]\s*/).filter(Boolean) : null;
        let finalBu = role === "po_admin" ? (buList ? null : bu) : bu;
        if (isScoped) {
          // A scoped P&O admin can only import staff into their own BU(s), and can never set Role.
          if (!bu || !scopeBUs.includes(bu)) { rejected++; continue; }
          role = "staff";
          finalBu = bu;
          buList = null;
        }
        // A delegate assignment locks once that person's current-cycle review
        // is completed — same rule as the one-by-one control, now enforced
        // here too so a bulk re-upload can't quietly overwrite a finalized record.
        let resolvedDelegateId = delegatePerson ? delegatePerson.id : idx >= 0 ? next[idx].reviewDelegateId || null : null;
        if (idx >= 0 && delegatePerson && delegatePerson.id !== next[idx].reviewDelegateId) {
          const existingData = await loadData(next[idx].id, next, cycle);
          if (existingData && existingData.manager_submitted) {
            resolvedDelegateId = next[idx].reviewDelegateId || null; // keep the locked, original value
            delegateLocked++;
          }
        }
        const personId = idx >= 0 ? next[idx].id : `imp-${Date.now()}-${username}`;
        const plainPasswordInFile = String(row["Password"] || "").trim();
        if (plainPasswordInFile) pendingPasswords.push({ id: personId, newPassword: plainPasswordInFile });
        const person = {
          id: personId,
          firstName: String(row["First Name"] || row["FirstName"] || "").trim() || (idx >= 0 ? next[idx].firstName : ""),
          lastName: String(row["Last Name"] || row["LastName"] || "").trim() || (idx >= 0 ? next[idx].lastName : ""),
          employeeId: employeeCode || (idx >= 0 ? next[idx].employeeId : ""),
          username,
          // Never write a plaintext password into the roster blob — keep the
          // existing hash (or "unset" for a brand-new row); pendingPasswords
          // above gets hashed server-side in one batch call after the loop.
          password: idx >= 0 ? next[idx].password : "unset",
          role,
          bu: finalBu,
          buList,
          department: String(row["Department Name"] || row["Department"] || "").trim(),
          jobGrade: String(row["Job Grade"] || row["JobGrade"] || "").trim(),
          designation: String(row["Designation"] || "").trim(),
          managerId: managerPerson ? managerPerson.id : idx >= 0 ? next[idx].managerId : null,
          functionalManagerId: fnMgrPerson ? fnMgrPerson.id : idx >= 0 ? next[idx].functionalManagerId || null : null,
          reviewDelegateId: resolvedDelegateId,
          needsEvaluation: (() => {
            const raw = String(row["Needs Evaluation"] || "").trim().toLowerCase();
            if (raw === "no" || raw === "n" || raw === "false") return false;
            if (raw === "yes" || raw === "y" || raw === "true") return true;
            return idx >= 0 ? next[idx].needsEvaluation !== false : true; // default true, keep existing otherwise
          })(),
        };
        if (idx >= 0) { next[idx] = person; updated++; } else { next.push(person); added++; }
      }
      await persist(next);
      if (pendingPasswords.length) {
        try {
          await fetch("/api/set-passwords", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ updates: pendingPasswords }),
          });
        } catch (e) {
          console.error("bulk set-passwords failed", e);
        }
      }
      setImportMsg(
        `Imported: ${added} added, ${updated} updated, ${skipped} skipped (missing username)${isScoped ? `, ${rejected} rejected (outside your BU)` : ""}${delegateLocked ? `, ${delegateLocked} delegate change${delegateLocked === 1 ? "" : "s"} skipped (review already completed)` : ""}${pendingPasswords.length ? `, ${pendingPasswords.length} password${pendingPasswords.length === 1 ? "" : "s"} set` : ""}.`
      );
      await appendLog(actorLabel, "Import", `${added} added, ${updated} updated, ${skipped} skipped${isScoped ? `, ${rejected} rejected` : ""}${delegateLocked ? `, ${delegateLocked} delegate changes blocked by completion lock` : ""}`);
    } catch (err) {
      setImportMsg("Import failed — check the file is a valid .xlsx/.csv export from this tool.");
    }
    if (fileRef.current) fileRef.current.value = "";
  };

  const exportRoster = () => {
    const data = visibleRoster.map((p) => {
      const mgr = p.managerId ? roster.find((m) => m.id === p.managerId) : null;
      const fnMgr = p.functionalManagerId ? roster.find((m) => m.id === p.functionalManagerId) : null;
      const delegate = p.reviewDelegateId ? roster.find((m) => m.id === p.reviewDelegateId) : null;
      return {
        "Company Code": p.bu || "",
        "Company Name": p.bu ? BU_FULL_NAME[p.bu] || "" : "",
        "Department Name": p.department || "",
        "Employee Code": p.employeeId || "",
        "Employee Name": `${p.firstName} ${p.lastName}`.trim(),
        Designation: p.designation || "",
        "Job Grade": p.jobGrade || "",
        "Manager Code": mgr?.employeeId || "",
        "Manager Name": mgr ? `${mgr.firstName} ${mgr.lastName}` : "",
        "Functional Manager Code": fnMgr?.employeeId || "",
        "Functional Manager Name": fnMgr ? `${fnMgr.firstName} ${fnMgr.lastName} (${fnMgr.bu})` : "",
        "Delegate Code": delegate?.employeeId || "",
        "Delegate Name": delegate ? `${delegate.firstName} ${delegate.lastName}` : "",
        "First Name": p.firstName,
        "Last Name": p.lastName,
        Username: p.username,
        Password: p.password,
        Role: p.role,
        "Company List": p.buList ? p.buList.join(", ") : "",
        "Needs Evaluation": p.role === "staff" ? (p.needsEvaluation === false ? "No" : "Yes") : "",
      };
    });
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [
      { wch: 12 }, { wch: 42 }, { wch: 16 }, { wch: 14 }, { wch: 22 }, { wch: 18 }, { wch: 10 },
      { wch: 14 }, { wch: 20 }, { wch: 16 }, { wch: 26 }, { wch: 14 }, { wch: 20 },
      { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 22 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Roster");
    XLSX.writeFile(wb, "gva_people_roster.xlsx");
    appendLog(actorLabel, "Export", `People roster — ${data.length} records`);
  };

  const editingPerson = editingId && editingId !== "new" ? roster.find((p) => p.id === editingId) : null;
  const [selectedIds, setSelectedIds] = useState(new Set());
  const [bulkManagerId, setBulkManagerId] = useState("");
  const staffOnly = visibleRoster.filter((p) => p.role === "staff");

  const toggleSelect = (id) => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  };
  const toggleSelectAll = () => {
    setSelectedIds((prev) => (prev.size === staffOnly.length ? new Set() : new Set(staffOnly.map((p) => p.id))));
  };
  const applyBulkManager = async () => {
    if (!bulkManagerId || !selectedIds.size) return;
    const newManager = roster.find((p) => p.id === bulkManagerId);
    const next = roster.map((p) => (selectedIds.has(p.id) ? { ...p, managerId: bulkManagerId } : p));
    await persist(next);
    appendLog(actorLabel, "Bulk Reassign", `${selectedIds.size} people reassigned to manager ${newManager ? `${newManager.firstName} ${newManager.lastName}` : bulkManagerId}`);
    setSelectedIds(new Set());
    setBulkManagerId("");
  };

  return (
    <div className="max-w-4xl mx-auto px-4 sm:px-6 pb-8">
      <div className="flex flex-col sm:flex-row sm:items-end sm:justify-between gap-3 mb-4">
        <div>
          <h1 className="text-xl font-semibold mb-1" style={{ color: BRAND.deep }}>People &amp; Roster</h1>
          <p className="text-xs text-slate-400">
            {isScoped ? `${visibleRoster.length} staff — ${scopeBUs.join(", ")} only` : `${roster.length} accounts — staff, P&O Admin, Master Admin`}
          </p>
        </div>
        <div className="flex items-center gap-2 flex-wrap">
          {!viewOnly && (
            <>
              <button onClick={() => fileRef.current?.click()} className="inline-flex items-center gap-1.5 px-3 py-2 text-sm font-medium rounded-xl border" style={{ borderColor: BRAND.line, color: BRAND.deep }}>
                <UploadCloud className="w-4 h-4" /> Import Excel
              </button>
              <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleImportFile} />
            </>
          )}
          <button onClick={exportRoster} className="inline-flex items-center gap-1.5 px-3 py-2 text-sm font-medium rounded-xl border" style={{ borderColor: BRAND.line, color: BRAND.deep }}>
            <Download className="w-4 h-4" /> Export Excel
          </button>
          {!viewOnly && (
            <button onClick={() => setEditingId("new")} className="inline-flex items-center gap-1.5 px-3 py-2 text-sm font-medium rounded-xl text-white" style={{ backgroundColor: BRAND.primary }}>
              <Plus className="w-4 h-4" /> Add Person
            </button>
          )}
        </div>
      </div>
      {viewOnly && (
        <div className="rounded-xl px-4 py-3 mb-4 text-xs" style={{ backgroundColor: "#FFF4F3", border: "1px solid #F6C8C3", color: "#8A1300" }}>
          Testing account — view and export only. Add, edit, and import are disabled.
        </div>
      )}

      {importMsg && (
        <div className="text-xs rounded-lg px-3 py-2 mb-4" style={{ backgroundColor: "#EAF7F6", color: BRAND.deep, border: `1px solid ${BRAND.line}` }}>
          {importMsg}
        </div>
      )}
      {deleteMsg && (
        <div className="text-xs rounded-lg px-3 py-2 mb-4" style={{ backgroundColor: "#EAF7F6", color: BRAND.deep, border: `1px solid ${BRAND.line}` }}>
          {deleteMsg}
        </div>
      )}

      <div className="rounded-xl border p-3 text-xs text-slate-500 mb-4 leading-relaxed" style={{ borderColor: BRAND.line, backgroundColor: "#FFF9E8" }}>
        <p><span className="font-semibold" style={{ color: BRAND.deep }}>Required:</span> Employee Code, First Name, Last Name, Username, Password{isScoped ? `, Company Code (${scopeBUs.join(" or ")})` : ", Company Code"}.</p>
        <p className="mt-1"><span className="font-semibold" style={{ color: BRAND.deep }}>Optional, by Employee Code:</span> Manager Code, Functional Manager Code (any company — e.g. a TPC Group dotted-line reviewer), Delegate Code, Needs Evaluation (Yes/No). Leave any of these blank to keep the existing value — blank never clears one.</p>
        <p className="mt-1"><span className="font-semibold" style={{ color: BRAND.deep }}>Made a mistake?</span> Re-upload a file with the correct rows — matching Employee Codes get updated, nothing else changes. Or edit one person directly in the table below.</p>
      </div>

      {editingId === "new" && (
        <PersonEditPanel initial={null} roster={roster} onSave={handleSaveEdit} onCancel={() => setEditingId(null)} restrictBUs={scopeBUs} restrictStaffOnly={isScoped} />
      )}
      {editingId && editingId !== "new" && editingPerson && (
        <PersonEditPanel initial={editingPerson} roster={roster} onSave={handleSaveEdit} onCancel={() => setEditingId(null)} restrictBUs={scopeBUs} restrictStaffOnly={isScoped} />
      )}

      {selectedIds.size > 0 && (
        <div className="rounded-xl px-4 py-3 mb-4 flex flex-wrap items-center gap-2.5" style={{ backgroundColor: "#F0F8F8", border: `1px solid ${BRAND.teal}` }}>
          <span className="text-sm font-medium" style={{ color: BRAND.deep }}>{selectedIds.size} selected</span>
          <span className="text-xs text-slate-400">Reassign manager to:</span>
          <select
            value={bulkManagerId}
            onChange={(e) => setBulkManagerId(e.target.value)}
            className="text-sm rounded-lg border px-2.5 py-1.5 focus:outline-none"
            style={{ borderColor: BRAND.line }}
          >
            <option value="">Choose a manager…</option>
            {staffOnly.filter((p) => !selectedIds.has(p.id)).map((p) => (
              <option key={p.id} value={p.id}>{p.firstName} {p.lastName} ({p.bu})</option>
            ))}
          </select>
          <button
            onClick={applyBulkManager}
            disabled={!bulkManagerId}
            className="text-sm font-medium rounded-lg px-3 py-1.5 text-white disabled:opacity-40"
            style={{ backgroundColor: BRAND.primary }}
          >
            Apply
          </button>
          <button onClick={() => setSelectedIds(new Set())} className="text-xs text-slate-400 hover:text-slate-600 ml-auto">
            Clear selection
          </button>
        </div>
      )}

      <div className="rounded-2xl overflow-hidden" style={{ border: `1px solid ${BRAND.line}`, backgroundColor: BRAND.card }}>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left text-xs" style={{ color: BRAND.gray, backgroundColor: "#F6FBFA" }}>
                <th className="px-4 py-3 font-medium">
                  {!viewOnly && (
                    <input type="checkbox" checked={staffOnly.length > 0 && selectedIds.size === staffOnly.length} onChange={toggleSelectAll} className="w-4 h-4" />
                  )}
                </th>
                <th className="px-3 py-3 font-medium">Name</th>
                <th className="px-3 py-3 font-medium">Employee Code</th>
                <th className="px-3 py-3 font-medium">Username</th>
                <th className="px-3 py-3 font-medium">Type</th>
                <th className="px-3 py-3 font-medium">Company</th>
                <th className="px-3 py-3 font-medium">Department</th>
                <th className="px-3 py-3 font-medium">Manager</th>
                <th className="px-5 py-3 font-medium"></th>
              </tr>
            </thead>
            <tbody className="divide-y" style={{ borderColor: BRAND.line }}>
              {visibleRoster.map((p) => (
                <React.Fragment key={p.id}>
                  <tr>
                    <td className="px-4 py-2.5">
                      {!viewOnly && p.role === "staff" && (
                        <input type="checkbox" checked={selectedIds.has(p.id)} onChange={() => toggleSelect(p.id)} className="w-4 h-4" />
                      )}
                    </td>
                    <td className="px-3 py-2.5 font-medium text-slate-800 whitespace-nowrap">{p.firstName} {p.lastName}</td>
                    <td className="px-3 py-2.5 text-slate-500 whitespace-nowrap">{p.employeeId || "—"}</td>
                    <td className="px-3 py-2.5 text-slate-500 whitespace-nowrap">{p.username}</td>
                    <td className="px-3 py-2.5 text-slate-500 whitespace-nowrap">
                      {ROLE_LABEL[p.role]}{p.testingOnly && <span className="ml-1 text-[10px] px-1.5 py-0.5 rounded-full" style={{ backgroundColor: "#FFF4F3", color: BRAND.red }}>testing only</span>}
                    </td>
                    <td className="px-3 py-2.5 text-slate-500">{p.buList ? p.buList.join(", ") : p.bu || "—"}</td>
                    <td className="px-3 py-2.5 text-slate-500">{p.department || "—"}</td>
                    <td className="px-3 py-2.5 text-slate-500 whitespace-nowrap">
                      {p.managerId ? (() => { const m = roster.find((x) => x.id === p.managerId); return m ? `${m.firstName} ${m.lastName}` : "—"; })() : "—"}
                    </td>
                    <td className="px-5 py-2.5 text-right">
                      {!viewOnly && (
                        <div className="inline-flex items-center gap-1.5">
                          <button onClick={() => { setEditingId(p.id); window.scrollTo({ top: 0, behavior: "smooth" }); }} className="inline-flex items-center gap-1 text-xs font-medium rounded-lg px-2.5 py-1 border hover:bg-slate-50" style={{ borderColor: BRAND.line, color: BRAND.deep }}>
                            <Pencil className="w-3 h-3" /> Edit
                          </button>
                          {p.role === "staff" && p.id !== admin.id && (
                            <button onClick={() => { setDeletingId(p.id); setDeleteMsg(""); }} className="inline-flex items-center gap-1 text-xs font-medium rounded-lg px-2.5 py-1 border hover:bg-red-50" style={{ borderColor: "#F6C8C3", color: BRAND.red }}>
                              <Trash2 className="w-3 h-3" />
                            </button>
                          )}
                        </div>
                      )}
                    </td>
                  </tr>
                  {deletingId === p.id && (
                    <tr>
                      <td colSpan={9} className="px-5 pb-4">
                        <div className="rounded-xl px-4 py-3 flex flex-wrap items-center gap-2.5" style={{ backgroundColor: "#FFF4F3", border: "1px solid #F6C8C3" }}>
                          <AlertCircle className="w-4 h-4 shrink-0" style={{ color: BRAND.red }} />
                          <span className="text-sm" style={{ color: "#8A1300" }}>
                            Delete {p.firstName} {p.lastName}? Their login stops working and they drop off all team lists and dashboards immediately. This can't be undone from here — you'd need to re-add them.
                          </span>
                          <div className="flex items-center gap-2 ml-auto">
                            <button onClick={() => setDeletingId(null)} className="text-xs font-medium rounded-lg px-3 py-1.5 border" style={{ borderColor: BRAND.line, color: BRAND.deep }}>
                              Cancel
                            </button>
                            <button onClick={() => handleDeletePerson(p)} className="text-xs font-medium rounded-lg px-3 py-1.5 text-white" style={{ backgroundColor: BRAND.red }}>
                              Yes, delete
                            </button>
                          </div>
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {isMasterAdmin && (
        <div className="rounded-2xl border mt-6" style={{ borderColor: "#F6C8C3" }}>
          <button onClick={() => setShowDangerZone(!showDangerZone)} className="w-full flex items-center justify-between px-4 py-3 text-sm font-medium" style={{ color: BRAND.red }}>
            <span className="flex items-center gap-2"><AlertCircle className="w-4 h-4" /> Danger Zone — UAT reset</span>
            {showDangerZone ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
          </button>
          {showDangerZone && (
            <div className="px-4 pb-4">
              <p className="text-xs text-slate-500 mb-3 leading-relaxed">
                Once a test round (UAT) is finished, use this to clear every submitted journey response and the
                activity log before handing the app to real users — so nobody sees test answers. This does <span className="font-medium">not</span> touch
                the roster: people, companies, and reporting lines are all kept exactly as they are.
              </p>
              {wipeMsg && (
                <div className="text-xs rounded-lg px-3 py-2 mb-3" style={{ backgroundColor: "#EAF7F6", color: BRAND.deep, border: `1px solid ${BRAND.line}` }}>
                  {wipeMsg}
                </div>
              )}
              <div className="flex flex-wrap items-center gap-2">
                <TextInput value={confirmWipeText} onChange={setConfirmWipeText} placeholder={`Type RESET to confirm`} />
                <button
                  onClick={handleWipeJourneyData}
                  disabled={confirmWipeText.trim().toUpperCase() !== "RESET" || wipeBusy}
                  className="text-sm font-medium rounded-xl px-4 py-2.5 text-white disabled:opacity-40"
                  style={{ backgroundColor: BRAND.red }}
                >
                  {wipeBusy ? "Clearing…" : "Clear all journey data"}
                </button>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function ExportLogView() {
  const [logs, setLogs] = useState(null);
  useEffect(() => {
    (async () => setLogs(await loadLogs()))();
  }, []);
  return (
    <div className="max-w-4xl mx-auto px-4 sm:px-6 pb-8">
      <h1 className="text-xl font-semibold mb-1 flex items-center gap-2" style={{ color: BRAND.deep }}>
        <FileClock className="w-5 h-5" style={{ color: BRAND.mint }} /> Activity Log
      </h1>
      <p className="text-xs text-slate-400 mb-4">Excel exports, roster imports, and every self-assessment submission or manager review completion are recorded here.</p>
      <div className="rounded-2xl overflow-hidden" style={{ border: `1px solid ${BRAND.line}`, backgroundColor: BRAND.card }}>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left text-xs" style={{ color: BRAND.gray, backgroundColor: "#F6FBFA" }}>
                <th className="px-5 py-3 font-medium">Timestamp</th>
                <th className="px-3 py-3 font-medium">Actor</th>
                <th className="px-3 py-3 font-medium">Action</th>
                <th className="px-5 py-3 font-medium">Detail</th>
              </tr>
            </thead>
            <tbody className="divide-y" style={{ borderColor: BRAND.line }}>
              {logs === null && <tr><td className="px-5 py-6 text-sm text-slate-400" colSpan={4}>Loading…</td></tr>}
              {logs?.map((l) => (
                <tr key={l.id}>
                  <td className="px-5 py-2.5 text-slate-500 whitespace-nowrap">{new Date(l.ts).toLocaleString()}</td>
                  <td className="px-3 py-2.5 text-slate-700 whitespace-nowrap">{l.actor}</td>
                  <td className="px-3 py-2.5">
                    <span className="text-xs font-medium px-2 py-0.5 rounded-full" style={{ backgroundColor: "#EAF7F6", color: BRAND.deep }}>{l.action}</span>
                  </td>
                  <td className="px-5 py-2.5 text-slate-500">{l.detail}</td>
                </tr>
              ))}
              {logs && !logs.length && <tr><td className="px-5 py-6 text-sm text-slate-400" colSpan={4}>No exports or imports yet.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function parseRoute(hash) {
  const clean = (hash || "").replace(/^#\/?/, "").toLowerCase().trim();
  if (!clean) return null; // gateway
  if (clean === "admin") return { mode: "admin" };
  if (BU_SLUGS[clean]) return { mode: "bu", bu: BU_SLUGS[clean] };
  return null; // unrecognized slug -> treat as gateway
}

export default function App() {
  // Roster starts populated immediately (SEED_ROSTER) so the UI is
  // interactive right away — it never blocks first render on a storage
  // round-trip. Any saved roster (from a previous import/edit) is then
  // loaded in the background and swapped in if found.
  const [roster, setRosterState] = useState(SEED_ROSTER);
  const [cycle, setCycleState] = useState(DEFAULT_CYCLE);
  const [cycleStart, setCycleStartState] = useState("");
  const [cycleEnd, setCycleEndState] = useState("");
  const [lang, setLang] = useState("th"); // "en" | "th" — Thai default: most users are Thai-speaking staff
  const [cyclesList, setCyclesListState] = useState([DEFAULT_CYCLE]);
  const [hash, setHash] = useState(() => (typeof window !== "undefined" ? window.location.hash : ""));
  const [currentUser, setCurrentUser] = useState(null);
  const [selectedEmployeeId, setSelectedEmployeeId] = useState(null);
  const [viewMode, setViewMode] = useState("self"); // "self" | "team" — for staff who also manage people
  const [data, setData] = useState(null);
  const [pageIdx, setPageIdx] = useState(0);
  const [saveState, setSaveState] = useState("idle"); // idle | saving | saved
  const [blockedMsg, setBlockedMsg] = useState("");
  const [showSubmitted, setShowSubmitted] = useState(false);
  const [confirmFinish, setConfirmFinish] = useState(false);
  const [printMode, setPrintMode] = useState(false);
  const saveTimer = useRef(null);
  const isFirstLoad = useRef(true);

  const isAdminRole = (r) => r === "po_admin" || r === "master_admin";
  const route = useMemo(() => parseRoute(hash), [hash]);

  // Background-load any previously saved roster; never blocks rendering.
  useEffect(() => {
    (async () => {
      try {
        const r = await loadRoster();
        if (r && r.length) {
          setRosterState(r);
        } else {
          await saveRoster(SEED_ROSTER);
        }
      } catch {
        // storage unavailable — keep working from SEED_ROSTER in memory
      }
    })();
  }, []);
  const setRoster = (next) => setRosterState(next);

  // Background-load the current review cycle and the list of known cycles.
  useEffect(() => {
    (async () => {
      try {
        const c = await loadCurrentCycle();
        if (c) setCycleState(c);
        else await saveCurrentCycle(DEFAULT_CYCLE);
        const list = await loadCyclesList();
        if (list && list.length) setCyclesListState(list);
        else await saveCyclesList([DEFAULT_CYCLE]);
      } catch {
        // storage unavailable — keep working from DEFAULT_CYCLE in memory
      }
    })();
  }, []);
  const startNewCycle = async (name) => {
    const trimmed = name.trim();
    if (!trimmed) return;
    const nextList = cyclesList.includes(trimmed) ? cyclesList : [...cyclesList, trimmed];
    setCyclesListState(nextList);
    setCycleState(trimmed);
    await saveCyclesList(nextList);
    await saveCurrentCycle(trimmed);
  };

  // Background-load the cycle's real start/end dates (the actual deadline
  // the company sets, not a per-person relative countdown).
  useEffect(() => {
    (async () => {
      try {
        const w = await loadCycleWindow();
        if (w) {
          setCycleStartState(w.start || "");
          setCycleEndState(w.end || "");
        }
      } catch {
        // storage unavailable — keep working with no deadline set
      }
    })();
  }, []);
  const setCycleWindow = async (start, end) => {
    setCycleStartState(start || "");
    setCycleEndState(end || "");
    await saveCycleWindow({ start: start || "", end: end || "" });
  };

  useEffect(() => {
    const onHashChange = () => setHash(window.location.hash);
    window.addEventListener("hashchange", onHashChange);
    return () => window.removeEventListener("hashchange", onHashChange);
  }, []);

  const inScopeFor = (user, r) => {
    if (!r) return false;
    if (user.role === "master_admin") return true; // Master Admin: every link, always
    if (r.mode === "admin") return false; // admin site is Master Admin only
    if (user.role === "po_admin") return adminBUs(user).includes(r.bu);
    return user.bu === r.bu;
  };

  const resetToLanding = (user) => {
    setCurrentUser(user);
    // Staff default to Thai; Master/P&O admins default to English — either
    // can still toggle manually, this just sets a sensible starting point.
    setLang(user.role === "staff" ? "th" : "en");
    if (user.role === "staff" && user.needsEvaluation === false) {
      // Pure reviewer / delegate-helper — no personal journey of their own.
      if (hasReports(user.id, roster)) {
        setViewMode("team");
        setSelectedEmployeeId(null);
      } else {
        setViewMode("delegated");
        setSelectedEmployeeId(null);
      }
    } else {
      setViewMode("self");
      setSelectedEmployeeId(user.role === "staff" ? user.id : null);
    }
    setData(null);
    setPageIdx(0);
    setShowSubmitted(false);
    setConfirmFinish(false);
    setPrintMode(false);
  };

  const handleLogin = (user) => {
    if (!inScopeFor(user, route)) return; // defense in depth — LoginScreen already filtered
    resetToLanding(user);
  };
  const handleSwitchUser = (user) => resetToLanding(user);
  const handleLogout = () => {
    setCurrentUser(null);
    setSelectedEmployeeId(null);
    setData(null);
    setPageIdx(0);
    setShowSubmitted(false);
    setConfirmFinish(false);
    setPrintMode(false);
  };
  const handleGoTeam = () => {
    setSelectedEmployeeId(null);
    setViewMode("team");
    setPageIdx(0);
    setShowSubmitted(false);
    setConfirmFinish(false);
    setPrintMode(false);
  };
  const handleGoSelf = () => {
    setSelectedEmployeeId(currentUser.id);
    setViewMode("self");
    setPageIdx(0);
    setShowSubmitted(false);
    setConfirmFinish(false);
    setPrintMode(false);
  };
  const handleGoDelegated = () => {
    setSelectedEmployeeId(null);
    setViewMode("delegated");
    setPageIdx(0);
    setShowSubmitted(false);
    setConfirmFinish(false);
    setPrintMode(false);
  };
  const handleOpenPerson = (id) => {
    setSelectedEmployeeId(id);
    setPageIdx(0);
    setShowSubmitted(false);
    setConfirmFinish(false);
    setPrintMode(false);
  };

  // A session never survives a site change — navigating to a different
  // BU's link (or the gateway) always starts fresh. Master Admin is the
  // one account type that keeps working across every link, by design.
  useEffect(() => {
    if (!currentUser) return;
    if (!route) { handleLogout(); return; }
    if (!inScopeFor(currentUser, route)) handleLogout();
  }, [route]); // eslint-disable-line

  useEffect(() => {
    if (!selectedEmployeeId || !roster) return;
    isFirstLoad.current = true;
    (async () => {
      const d = await loadData(selectedEmployeeId, roster, cycle);
      setData({ ...defaultData(), ...(d || {}) });
    })();
  }, [selectedEmployeeId, roster, cycle]);

  useEffect(() => {
    if (!data || !selectedEmployeeId || !roster) return;
    if (isFirstLoad.current) {
      isFirstLoad.current = false;
      return;
    }
    setSaveState("saving");
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      await saveData(selectedEmployeeId, data, roster, cycle);
      setSaveState("saved");
    }, 700);
    return () => clearTimeout(saveTimer.current);
  }, [data]); // eslint-disable-line

  useEffect(() => setBlockedMsg(""), [pageIdx, selectedEmployeeId]);

  const setField = (key, value) => setData((prev) => ({ ...prev, [key]: value }));

  const ctxValue = { roster, setRoster: (r) => { setRosterState(r); }, cycle, cyclesList, startNewCycle, lang, setLang, cycleStart, cycleEnd, setCycleWindow };
  const wrap = (node) => <RosterCtx.Provider value={ctxValue}>{node}</RosterCtx.Provider>;

  if (!route) return wrap(<GatewayPage />);
  if (!currentUser) return wrap(<LoginScreen scope={route} onLogin={handleLogin} />);

  const iAmManager = currentUser.role === "staff" && hasReports(currentUser.id, roster);
  const needsOwnEval = currentUser.role === "staff" && currentUser.needsEvaluation !== false;
  const delegatedReports = currentUser.role === "staff" ? roster.filter((p) => p.role === "staff" && (p.reviewDelegateId === currentUser.id || p.functionalManagerId === currentUser.id)) : [];

  if (currentUser.role === "staff" && viewMode === "delegated" && !selectedEmployeeId) {
    return wrap(
      <MyDelegatedReviews
        delegate={currentUser}
        reports={delegatedReports}
        onOpenEmployee={handleOpenPerson}
        onExit={iAmManager ? handleGoTeam : needsOwnEval ? handleGoSelf : null}
        canExitToTeam={iAmManager}
        onLogout={handleLogout}
        onSwitchUser={handleSwitchUser}
      />
    );
  }

  if (currentUser.role === "staff" && iAmManager && viewMode === "team" && !selectedEmployeeId) {
    return wrap(
      <ManagerTeamList
        manager={currentUser}
        onOpenEmployee={handleOpenPerson}
        onGoSelf={handleGoSelf}
        onLogout={handleLogout}
        onSwitchUser={handleSwitchUser}
        delegatedCount={delegatedReports.length}
        onGoDelegated={handleGoDelegated}
      />
    );
  }

  if (isAdminRole(currentUser.role) && !selectedEmployeeId) {
    return wrap(
      <AdminView admin={currentUser} onOpenEmployee={handleOpenPerson} onLogout={handleLogout} onSwitchUser={handleSwitchUser} />
    );
  }

  if (!data) {
    return wrap(<div className="min-h-screen" style={{ background: BRAND.bgGradient }} />);
  }

  const isAdminViewing = isAdminRole(currentUser.role);
  const viewingSelf = selectedEmployeeId === currentUser.id;
  const effectiveRole = isAdminViewing ? "admin" : viewingSelf ? "employee" : "manager";
  const person = roster.find((p) => p.id === selectedEmployeeId);
  const personManager = person?.managerId ? roster.find((p) => p.id === person.managerId) : null;
  const isPeopleManagerForThis = hasReports(selectedEmployeeId, roster);
  // A delegate or functional manager can draft the manager's side, but only
  // the real direct manager can give the final, audited confirmation.
  const isDelegateReviewing = effectiveRole === "manager" && person?.reviewDelegateId === currentUser.id && person?.managerId !== currentUser.id;
  const isFunctionalManagerReviewing = effectiveRole === "manager" && person?.functionalManagerId === currentUser.id && person?.managerId !== currentUser.id;
  const isHelperReviewing = isDelegateReviewing || isFunctionalManagerReviewing;
  const managerCanFinalize = effectiveRole === "manager" && !isHelperReviewing;

  const visibleSections = SECTIONS.filter((s) => !s.managerOnly || isPeopleManagerForThis);
  // Progress reflects only the current viewer's own side of the document —
  // an employee who has filled in everything they're responsible for sees
  // 100%, regardless of whether their manager has responded yet.
  const progress = isAdminViewing ? computeProgress(data, isPeopleManagerForThis, false) : roleProgress(visibleSections, data, effectiveRole);
  const section = visibleSections[pageIdx];
  const isLast = pageIdx === visibleSections.length - 1;
  const isFirst = pageIdx === 0;
  const pageComplete = isPageComplete(section, data, effectiveRole);

  const managerFieldKeys = visibleSections.filter((s) => s.id !== "cover").flatMap((s) => pageRequiredFor(s, "manager"));
  const managerReadyCount = managerFieldKeys.filter((k) => {
    const v = data[k];
    return v === true || (typeof v === "string" && v.trim().length > 0);
  }).length;
  const managerTotalCount = managerFieldKeys.length;

  const goTo = (i) => setPageIdx(Math.max(0, Math.min(visibleSections.length - 1, i)));
  const goNext = () => {
    if (!pageComplete) {
      setBlockedMsg(t("Please complete this section before moving on.", "กรุณากรอกข้อมูลในส่วนนี้ให้ครบก่อนไปต่อ", lang));
      return;
    }
    setBlockedMsg("");
    goTo(pageIdx + 1);
  };
  const goBackSite = () => {
    if (isAdminViewing) setSelectedEmployeeId(null);
    else if (viewMode === "delegated") setSelectedEmployeeId(null); // stay in the delegated reviews list
    else handleGoTeam();
  };
  const handleFinishClick = () => {
    if (effectiveRole === "employee") {
      if (data.employee_submitted) return;
      const allDone = visibleSections.every((s) => isPageComplete(s, data, "employee"));
      if (!allDone) {
        setBlockedMsg(t("Please complete every section before submitting.", "ท่านยังกรอกแบบประเมินไม่ครบ กรุณากรอกแบบประเมินให้ครบทุกส่วน", lang));
        return;
      }
      setBlockedMsg("");
      setConfirmFinish(true);
    } else if (effectiveRole === "manager") {
      if (managerCanFinalize && data.manager_submitted) return;
      const allDone = visibleSections.every((s) => isPageComplete(s, data, "manager"));
      if (!allDone) {
        setBlockedMsg(t("Please complete every section before finishing.", "ท่านยังกรอกความเห็นไม่ครบ กรุณากรอกให้ครบทุกส่วน", lang));
        return;
      }
      setBlockedMsg("");
      setConfirmFinish(true);
    }
  };
  const handleFinishConfirmed = async () => {
    setConfirmFinish(false);
    const personLabel = person ? `${person.firstName} ${person.lastName}` : "Unknown";
    const managerLabel = personManager ? `${personManager.firstName} ${personManager.lastName}` : "the manager";
    let updated = data;
    if (effectiveRole === "employee") {
      updated = { ...data, employee_submitted: true, employee_submitted_at: new Date().toISOString() };
      setData(updated);
      await saveData(selectedEmployeeId, updated, roster, cycle);
      setShowSubmitted(true);
      appendLog(personLabel, "Self-Assessment Submitted", `${personLabel} submitted their self-assessment for manager review.`);
    } else if (effectiveRole === "manager" && managerCanFinalize) {
      updated = { ...data, manager_submitted: true, manager_submitted_at: new Date().toISOString(), manager_submitted_by: currentUser.id };
      setData(updated);
      await saveData(selectedEmployeeId, updated, roster, cycle);
      setShowSubmitted(true);
      const draftedBySomeoneElse = data.manager_draft_by && data.manager_draft_by !== currentUser.id;
      const drafter = draftedBySomeoneElse ? roster.find((p) => p.id === data.manager_draft_by) : null;
      appendLog(
        `${currentUser.firstName} ${currentUser.lastName} (Manager)`,
        "Manager Review Completed",
        drafter
          ? `${currentUser.firstName} ${currentUser.lastName} confirmed and finalized the review for ${personLabel}. Draft was prepared by ${drafter.firstName} ${drafter.lastName} on their behalf.`
          : `Completed review for ${personLabel}.`
      );
    } else if (effectiveRole === "manager" && !managerCanFinalize) {
      // Delegate reviewer: this only records a draft — the real manager still has to confirm.
      updated = { ...data, manager_draft_by: currentUser.id, manager_draft_at: new Date().toISOString() };
      setData(updated);
      await saveData(selectedEmployeeId, updated, roster, cycle);
      setShowSubmitted(true);
      appendLog(
        `${currentUser.firstName} ${currentUser.lastName} (delegate, on behalf of ${managerLabel})`,
        "Manager Review Drafted",
        `${currentUser.firstName} ${currentUser.lastName} drafted the manager review for ${personLabel} on behalf of ${managerLabel}. Awaiting ${managerLabel}'s confirmation.`
      );
    }
  };

  if (printMode) {
    return wrap(
      <PrintView
        person={person}
        manager={personManager}
        data={data}
        visibleSections={visibleSections}
        onBack={() => setPrintMode(false)}
      />
    );
  }

  if (confirmFinish) {
    const managerLabel = personManager ? `${personManager.firstName} ${personManager.lastName}` : t("your direct manager", "หัวหน้างานของคุณ", lang);
    const drafter = data.manager_draft_by && data.manager_draft_by !== currentUser.id ? roster.find((p) => p.id === data.manager_draft_by) : null;
    return wrap(
      <ConfirmDialog
        title={
          effectiveRole === "employee"
            ? t("Submit your journey?", "ส่งแบบประเมินของคุณ?", lang)
            : managerCanFinalize
            ? t("Confirm your review and send your feedback?", "ยืนยันการตรวจสอบและส่งความเห็น?", lang)
            : t(`Send to ${managerLabel} for confirmation?`, `ส่งให้ ${managerLabel} เพื่อยืนยัน?`, lang)
        }
        body={
          effectiveRole === "employee"
            ? t("You will not be able to change your answers after submitting. Once submitted, your manager will review and share their feedback.", "คุณจะไม่สามารถแก้ไขคำตอบได้หลังจากส่งแบบประเมิน เมื่อส่งแบบประเมินแล้ว หัวหน้างานของคุณจะเข้ามาตรวจสอบและแสดงความคิดเห็นต่อไป", lang)
            : managerCanFinalize
            ? drafter
              ? t(
                  `This is final — you will not be able to change your feedback after finishing. ${drafter.firstName} ${drafter.lastName} drafted this on your behalf; confirming means you approve it as your own.`,
                  `นี่คือขั้นตอนสุดท้าย — คุณจะไม่สามารถแก้ไขความเห็นได้หลังจากเสร็จสิ้น ${drafter.firstName} ${drafter.lastName} ได้ร่างเนื้อหานี้แทนคุณ การยืนยันหมายความว่าคุณอนุมัติเนื้อหานี้ในฐานะของคุณเอง`,
                  lang
                )
              : t("Once you confirm, you will not be able to change your feedback again. The employee will be able to see your feedback after this.", "เมื่อกดยืนยันการประเมินแล้ว คุณจะไม่สามารถแก้ไขความเห็นได้อีก พนักงานจะสามารถเห็นความเห็นของคุณได้หลังจากนี้", lang)
            : t(
                `${managerLabel} will need to review your draft and give the final confirmation before it becomes official. You can still edit your draft until then.`,
                `${managerLabel} จะต้องตรวจสอบร่างของคุณและให้การยืนยันขั้นสุดท้ายก่อนจึงจะเป็นทางการ คุณยังสามารถแก้ไขร่างได้จนกว่าจะถึงเวลานั้น`,
                lang
              )
        }
        confirmLabel={
          effectiveRole === "employee"
            ? t("Yes, submit", "ใช่ ส่งเลย", lang)
            : managerCanFinalize
            ? t("Yes, finish", "ใช่ เสร็จสิ้น", lang)
            : t("Send draft", "ส่งร่าง", lang)
        }
        onCancel={() => setConfirmFinish(false)}
        onConfirm={handleFinishConfirmed}
      />
    );
  }

  if (showSubmitted) {
    return wrap(
      <SubmittedScreen
        variant={effectiveRole === "employee" ? "employee" : managerCanFinalize ? "manager" : "delegate"}
        onContinue={() => {
          setShowSubmitted(false);
          setConfirmFinish(false);
          if (effectiveRole === "employee") goTo(0);
          else if (managerCanFinalize) handleGoTeam();
          else handleGoDelegated();
        }}
      />
    );
  }

  return wrap(
    <div className="min-h-screen" style={{ background: BRAND.bgGradient, fontFamily: FONT_STACK }}>
      <style>{`${FONT_IMPORT}`}</style>

      {/* Header */}
      <div className="sticky top-0 z-10 border-b" style={{ backgroundColor: "rgba(246,251,250,0.92)", backdropFilter: "blur(6px)", borderColor: BRAND.line }}>
        <div style={{ height: 3, background: `linear-gradient(to right, ${BRAND.deep}, ${BRAND.teal}, ${BRAND.mint})` }} />
        <div className="max-w-3xl mx-auto px-4 sm:px-6 py-3.5">
          <div className="flex flex-wrap items-center justify-between gap-3 mb-3">
            <div className="flex items-center gap-2 min-w-0">
              {!(currentUser.role === "staff" && viewingSelf && !isAdminViewing) && (
                <button
                  onClick={goBackSite}
                  className="p-1.5 rounded-lg hover:bg-white shrink-0"
                  style={{ color: BRAND.deep }}
                  title={isAdminViewing ? "Back to dashboard" : "Back to team"}
                >
                  <ArrowLeft className="w-4 h-4" />
                </button>
              )}
              {person && <Avatar name={`${person.firstName} ${person.lastName}`} size={28} />}
              <div className="min-w-0">
                <p className="text-sm font-semibold truncate" style={{ color: BRAND.deep }}>
                  {person ? `${person.firstName} ${person.lastName}` : "Growth & Value-Add Journey"}
                </p>
                <p className="text-xs text-slate-400 truncate">
                  tpcgvajourney/{route.mode === "admin" ? "admin" : SLUG_BY_BU[route.bu]} · {section.nav} · Viewing as {effectiveRole === "employee" ? "Employee" : effectiveRole === "manager" ? "Manager" : ROLE_LABEL[currentUser.role]}
                  {isFunctionalManagerReviewing ? " (as second reviewer)" : isDelegateReviewing ? " (as delegate)" : ""}
                </p>
              </div>
            </div>
            <div className="flex items-center gap-2 flex-wrap">
              {!isAdminViewing && (iAmManager || delegatedReports.length > 0) && (
                <div className="flex rounded-full p-0.5" style={{ backgroundColor: "#E9F3F2" }}>
                  {currentUser.role === "staff" && currentUser.needsEvaluation !== false && (
                    <button
                      onClick={handleGoSelf}
                      className="flex items-center gap-1 text-xs font-medium px-2.5 py-1 rounded-full transition"
                      style={viewingSelf ? { backgroundColor: BRAND.primary, color: "white" } : { color: BRAND.deep }}
                    >
                      <User className="w-3 h-3" /> {t("My Journey", UI_TH.myJourney, lang)}
                    </button>
                  )}
                  {iAmManager && (
                    <button
                      onClick={handleGoTeam}
                      className="flex items-center gap-1 text-xs font-medium px-2.5 py-1 rounded-full transition"
                      style={!viewingSelf && !isHelperReviewing ? { backgroundColor: BRAND.primary, color: "white" } : { color: BRAND.deep }}
                    >
                      <UsersIcon className="w-3 h-3" /> {t("My Team", UI_TH.myTeam, lang)}
                    </button>
                  )}
                  {delegatedReports.length > 0 && (
                    <button
                      onClick={handleGoDelegated}
                      className="flex items-center gap-1 text-xs font-medium px-2.5 py-1 rounded-full transition"
                      style={isHelperReviewing ? { backgroundColor: BRAND.primary, color: "white" } : { color: BRAND.deep }}
                    >
                      <RefreshCcw className="w-3 h-3" /> {t("Delegated", UI_TH.delegated, lang)} ({delegatedReports.length})
                    </button>
                  )}
                </div>
              )}
              <span className="text-xs text-slate-400 hidden sm:inline">
                {saveState === "saving" ? "Saving…" : saveState === "saved" ? "Saved" : ""}
              </span>
              <Save className="w-3.5 h-3.5 text-slate-300" />
              {!viewingSelf && (
                <button onClick={() => setPrintMode(true)} className="p-1.5 rounded-lg hover:bg-white" style={{ color: BRAND.deep }} title="Print / Save as PDF">
                  <Printer className="w-4 h-4" />
                </button>
              )}
              <LangToggle />
              <button onClick={handleLogout} className="p-1.5 text-slate-400 hover:text-red-600" title="Log out">
                <LogOut className="w-4 h-4" />
              </button>
            </div>
          </div>

          {/* progress bar */}
          <div className="flex items-center gap-3">
            <div className="flex-1 h-1.5 rounded-full overflow-hidden" style={{ backgroundColor: "#DCEEEC" }}>
              <div className="h-full rounded-full transition-all duration-500" style={{ width: `${progress}%`, backgroundColor: BRAND.mint }} />
            </div>
            <span className="text-xs font-semibold tabular-nums" style={{ color: BRAND.primary }}>
              {progress}%
            </span>
          </div>

          {/* section pills */}
          <div className="flex gap-1.5 overflow-x-auto mt-3 pb-1 -mx-1 px-1">
            {visibleSections.map((s, i) => (
              <button
                key={s.id}
                onClick={() => goTo(i)}
                className="shrink-0 text-xs px-2.5 py-1 rounded-full font-medium transition flex items-center gap-1"
                style={
                  i === pageIdx
                    ? { backgroundColor: BRAND.deep, color: "white" }
                    : { backgroundColor: "white", color: BRAND.gray, border: `1px solid ${BRAND.line}` }
                }
              >
                {s.id === "cover" ? <User className="w-3 h-3" /> : s.track === "A" ? <Sprout className="w-3 h-3" /> : <Building2 className="w-3 h-3" />}
                {s.nav}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* Page content */}
      <div className="max-w-3xl mx-auto px-4 sm:px-6 py-8">
        {effectiveRole === "manager" && managerCanFinalize && data.manager_draft_by && !data.manager_submitted && (
          <div className="rounded-xl px-4 py-3 mb-4 flex items-start gap-2.5" style={{ backgroundColor: "#FFF9E8", border: "1px solid #F3E3A8" }}>
            <RefreshCcw className="w-4 h-4 mt-0.5 shrink-0" style={{ color: "#8A6D00" }} />
            <p className="text-sm" style={{ color: "#8A6D00" }}>
              <span className="font-medium">
                Drafted by {(() => { const d = roster.find((p) => p.id === data.manager_draft_by); return d ? `${d.firstName} ${d.lastName}` : "a delegate"; })()}
                {data.manager_draft_at ? ` on ${new Date(data.manager_draft_at).toLocaleDateString()}` : ""}.
              </span>{" "}
              Review, edit as needed, then confirm to finalize.
            </p>
          </div>
        )}
        {effectiveRole === "manager" && !managerCanFinalize && (
          <div className="rounded-xl px-4 py-3 mb-4 flex items-start gap-2.5" style={{ backgroundColor: "#F0F8F8", border: `1px solid ${BRAND.line}` }}>
            <RefreshCcw className="w-4 h-4 mt-0.5 shrink-0" style={{ color: BRAND.primary }} />
            <p className="text-sm" style={{ color: BRAND.deep }}>
              {isFunctionalManagerReviewing
                ? `Adding your input for ${personManager ? `${personManager.firstName} ${personManager.lastName}` : "the direct manager"} to confirm.`
                : `Drafting on behalf of ${personManager ? `${personManager.firstName} ${personManager.lastName}` : "the direct manager"}. Final confirmation rests with them.`}
            </p>
          </div>
        )}
        <div className="rounded-2xl p-6 sm:p-8" style={{ backgroundColor: BRAND.card, border: `1px solid ${BRAND.line}` }}>
          {section.id === "cover" && <CoverPage data={data} setField={setField} role={effectiveRole} person={person} manager={personManager} isPeopleManager={isPeopleManagerForThis} managerReadyCount={managerReadyCount} managerTotalCount={managerTotalCount} />}
          {section.id === "a1" && <GrowthStagePage section={section} data={data} setField={setField} role={effectiveRole} />}
          {section.id === "a2" && <GrowthStagePage section={section} data={data} setField={setField} role={effectiveRole} />}
          {section.id === "a3" && <GrowthStagePage section={section} data={data} setField={setField} role={effectiveRole} />}
          {section.id === "a4" && <GrowthStagePage section={section} data={data} setField={setField} role={effectiveRole} />}
          {section.id === "a5" && <GrowthPlanPage data={data} setField={setField} role={effectiveRole} />}
          {["b1-learning", "b1-integrity", "b1-coaching", "b1-connection", "b2-collab", "b2-steward", "b2-innovation"].includes(section.id) && (
            <ReflectionPage section={section} data={data} setField={setField} role={effectiveRole} />
          )}
          {section.id === "b3" && <B3Page data={data} setField={setField} role={effectiveRole} />}
          {section.id === "b4" && <B4Page data={data} setField={setField} role={effectiveRole} />}
        </div>

        {/* Nav footer */}
        <div className="flex items-center justify-between mt-5">
          <button
            onClick={() => { setBlockedMsg(""); goTo(pageIdx - 1); }}
            disabled={isFirst}
            className={`inline-flex items-center gap-1.5 px-4 py-2 text-sm font-medium rounded-xl border ${isFirst ? "opacity-30 cursor-not-allowed" : "hover:bg-white"}`}
            style={{ borderColor: BRAND.line, color: BRAND.deep }}
          >
            <ChevronLeft className="w-4 h-4" /> {t("Back", UI_TH.back, lang)}
          </button>
          <span className="text-xs text-center px-2" style={blockedMsg ? { color: BRAND.red } : { color: "#94A3B8" }}>
            {blockedMsg || t(`${pageIdx + 1} of ${visibleSections.length}`, `${pageIdx + 1} จาก ${visibleSections.length}`, lang)}
          </span>
          {!isLast ? (
            <button
              onClick={goNext}
              className="inline-flex items-center gap-1.5 px-4 py-2 text-sm font-medium rounded-xl text-white"
              style={{ backgroundColor: BRAND.primary }}
            >
              {t("Next", UI_TH.next, lang)} <ChevronRight className="w-4 h-4" />
            </button>
          ) : effectiveRole === "employee" && data.employee_submitted ? (
            <span className="inline-flex items-center gap-1.5 px-4 py-2 text-sm font-medium rounded-xl" style={{ backgroundColor: "#EAF7F6", color: BRAND.deep }}>
              <Check className="w-4 h-4" /> {t("Submitted", UI_TH.submitted, lang)}
            </span>
          ) : effectiveRole === "manager" && data.manager_submitted ? (
            <span className="inline-flex items-center gap-1.5 px-4 py-2 text-sm font-medium rounded-xl" style={{ backgroundColor: "#EAF7F6", color: BRAND.deep }}>
              <Check className="w-4 h-4" /> {t("Review submitted", UI_TH.reviewSubmitted, lang)}
            </span>
          ) : effectiveRole === "admin" ? null : (
            <button
              onClick={handleFinishClick}
              className="inline-flex items-center gap-1.5 px-4 py-2 text-sm font-medium rounded-xl text-white"
              style={{ backgroundColor: BRAND.red }}
            >
              <Check className="w-4 h-4" />{" "}
              {effectiveRole === "employee"
                ? t("Finish & Submit", UI_TH.finishSubmit, lang)
                : managerCanFinalize
                ? t("Finish Review", UI_TH.finishReview, lang)
                : lang === "th"
                ? `ส่งให้${personManager ? personManager.firstName : "หัวหน้างาน"}ยืนยัน`
                : `Send for ${personManager ? personManager.firstName : "manager"}'s confirmation`}
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

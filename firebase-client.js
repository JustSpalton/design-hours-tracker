import { initializeApp } from "https://www.gstatic.com/firebasejs/12.3.0/firebase-app.js";
import { getAuth, signInAnonymously, setPersistence, browserSessionPersistence } from "https://www.gstatic.com/firebasejs/12.3.0/firebase-auth.js";
import {
  getFirestore,
  doc,
  getDoc,
  runTransaction,
  serverTimestamp,
  setDoc,
  deleteDoc,
  Timestamp,
  collection,
  getDocs,
  writeBatch
} from "https://www.gstatic.com/firebasejs/12.3.0/firebase-firestore.js";

const config = window.DESIGN_HOURS_FIREBASE_CONFIG;
if (!config || String(config.projectId || "").startsWith("REPLACE_")) {
  throw new Error("Firebase project configuration has not been installed.");
}

const firebaseApp = initializeApp(config);
const auth = getAuth(firebaseApp);
const db = getFirestore(firebaseApp);
const stateRef = doc(db, "tracker", "state");
const breakdownRef = doc(db, "tracker", "breakdowns");
const ncrChunksRef = collection(db, "ncrChunks");

window.firebaseReady = (async () => {
  await setPersistence(auth, browserSessionPersistence);
  if (typeof auth.authStateReady === "function") await auth.authStateReady();
  if (auth.currentUser) return { user: auth.currentUser };
  return signInAnonymously(auth);
})();

const clone = value => JSON.parse(JSON.stringify(value));
const cleanName = value => String(value || "").replace(/\s+/g, " ").trim();
const designerAliases = new Map([
  ["jon wilson", "Jonathan Wilson"],
  ["kerry mui", "Kerry Gardiner"]
]);
const canonicalDesignerName = value => {
  const cleaned = cleanName(value);
  return designerAliases.get(cleaned.toLowerCase()) || cleaned;
};
const canonicalizeStateNames = state => {
  for (const designer of state.designers || []) designer.name = canonicalDesignerName(designer.name);
  for (const row of state.hours || []) row.designer = canonicalDesignerName(row.designer);
  for (const row of state.holidays || []) row.designer = canonicalDesignerName(row.designer);
  return state;
};
const canonicalizeBreakdownRecords = records => {
  for (const row of records || []) row.designer = canonicalDesignerName(row.designer);
  return records;
};
const validDate = value => /^\d{4}-\d{2}-\d{2}$/.test(String(value || ""));
const defaultTeams = () => [
  { id: "my-team", name: "My Team", designer_ids: [] },
  { id: "lynseys-team", name: "Lynsey's Team", designer_ids: [] },
  { id: "abbies-team", name: "Abbie's Team", designer_ids: [] }
];
const ensureTeams = state => {
  if (!Array.isArray(state.teams) || !state.teams.length) state.teams = defaultTeams();
  return state.teams;
};

function parseBody(options) {
  if (!options?.body) return {};
  return typeof options.body === "string" ? JSON.parse(options.body) : options.body;
}

function cleanCounts(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  const result = {};
  for (const [key, raw] of Object.entries(value)) {
    const label = cleanName(key);
    const count = Number(raw);
    if (label && Number.isFinite(count) && count >= 0) result[label] = Math.round(count);
  }
  return result;
}

function cleanHours(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  const result = {};
  for (const [key, raw] of Object.entries(value)) {
    const label = cleanName(key);
    const hours = Number(raw);
    if (label && Number.isFinite(hours) && hours >= 0) result[label] = Math.round(hours * 100) / 100;
  }
  return result;
}

function cleanNcrText(value, max = 2000) {
  return String(value ?? "").replace(/\s+/g, " ").trim().slice(0, max);
}
function cleanNcrNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? Math.round(n * 100) / 100 : 0;
}
function cleanNcrRecord(raw) {
  return {
    rowId: cleanNcrText(raw?.rowId, 80),
    ncrNo: cleanNcrText(raw?.ncrNo, 40),
    dateReported: validDate(raw?.dateReported) ? String(raw.dateReported) : "",
    ifo: cleanNcrText(raw?.ifo, 120),
    invoiceMonth: cleanNcrText(raw?.invoiceMonth, 40),
    customer: cleanNcrText(raw?.customer, 180),
    fnumber: cleanNcrText(raw?.fnumber, 80),
    site: cleanNcrText(raw?.site, 300),
    complaint: cleanNcrText(raw?.complaint, 2400),
    employee: cleanNcrText(raw?.employee, 60),
    approvedBy: cleanNcrText(raw?.approvedBy, 60),
    category: cleanNcrText(raw?.category, 100),
    beams: cleanNcrNumber(raw?.beams),
    posi: cleanNcrNumber(raw?.posi),
    ancillaries: cleanNcrNumber(raw?.ancillaries),
    delivery: cleanNcrNumber(raw?.delivery),
    total: cleanNcrNumber(raw?.total),
    correctiveAction: cleanNcrText(raw?.correctiveAction, 2400),
    investigated: cleanNcrText(raw?.investigated, 300),
    finding: cleanNcrText(raw?.finding, 300),
    productionComments: cleanNcrText(raw?.productionComments, 2400),
    details: cleanNcrText(raw?.details, 300),
    collectReplace: cleanNcrText(raw?.collectReplace, 300),
    notes: cleanNcrText(raw?.notes, 3000)
  };
}

async function readRequired(ref, label) {
  const snapshot = await getDoc(ref);
  if (!snapshot.exists()) throw new Error(`${label} has not been migrated to Firebase yet.`);
  return snapshot.data();
}

window.firebaseApi = async function firebaseApi(path, options = {}) {
  await window.firebaseReady;
  const method = String(options.method || "GET").toUpperCase();

  if (path === "/api/access" && method === "GET") {
    await getDoc(stateRef);
    return { ok: true };
  }

  if (path === "/api/unlock" && method === "POST") {
    if (path === "/api/ncr" && method === "GET") {
    const snapshot = await getDocs(ncrChunksRef);
    const chunks = snapshot.docs.map(item => item.data()).sort((a, b) => Number(a.index || 0) - Number(b.index || 0));
    const records = chunks.flatMap(item => Array.isArray(item.records) ? item.records : []);
    const latest = chunks.find(item => item.sourceFile || item.importedAt) || {};
    return { records, sourceFile: latest.sourceFile || "", importedAt: latest.importedAt || null };
  }

    const body = parseBody(options);

  if (path === "/api/ncr-import" && method === "POST") {
    const incoming = Array.isArray(body.records) ? body.records.slice(0, 6000) : [];
    if (!incoming.length) throw new Error("No NCR records were supplied.");
    const records = incoming.map(cleanNcrRecord).filter(row => row.ncrNo);
    if (!records.length) throw new Error("No valid NCR rows were found.");
    const sourceFile = cleanNcrText(body.sourceFile || "NCR Excel import", 255);
    const importedAt = new Date().toISOString();
    const chunkSize = 150;
    const chunks = [];
    for (let i = 0; i < records.length; i += chunkSize) chunks.push(records.slice(i, i + chunkSize));

    const existing = await getDocs(ncrChunksRef);
    const batch = writeBatch(db);
    for (const item of existing.docs) batch.delete(item.ref);
    chunks.forEach((rows, index) => {
      const ref = doc(db, "ncrChunks", `chunk-${String(index).padStart(3, "0")}`);
      batch.set(ref, { index, sourceFile, importedAt, records: rows });
    });
    await batch.commit();
    return { ok: true, saved: records.length, chunks: chunks.length, sourceFile, importedAt, records };
  }
    const pin = String(body.pin || "").trim();
    if (!/^\d{4}$/.test(pin)) throw new Error("Enter a 4-digit PIN.");
    const user = auth.currentUser;
    if (!user) throw new Error("Secure session is not ready. Refresh and try again.");
    const sessionRef = doc(db, "accessSessions", user.uid);
    const expiresAt = Timestamp.fromMillis(Date.now() + 12 * 60 * 60 * 1000);
    try {
      await setDoc(sessionRef, { pin, verified: true, expiresAt });
      return { ok: true };
    } catch (error) {
      if (String(error?.code || "").includes("permission-denied")) {
        throw new Error("Incorrect PIN.");
      }
      throw error;
    }
  }

  if (path === "/api/lock" && method === "POST") {
    const user = auth.currentUser;
    if (user) {
      try { await deleteDoc(doc(db, "accessSessions", user.uid)); } catch (_) {}
    }
    return { ok: true };
  }

  if (path === "/api/data" && method === "GET") {
    const data = await readRequired(stateRef, "Tracker data");
    const result = canonicalizeStateNames(clone(data));
    ensureTeams(result);
    delete result.updatedAt;
    return result;
  }


  if (path === "/api/breakdowns" && method === "GET") {
    const snapshot = await getDoc(breakdownRef);
    return snapshot.exists() ? { records: canonicalizeBreakdownRecords(clone(snapshot.data().records || [])) } : { records: [] };
  }

  const body = parseBody(options);

  if (path === "/api/designers" && method === "POST") {
    const name = canonicalDesignerName(body.name);
    if (name.length < 3 || name.length > 80) throw new Error("Enter a valid designer name.");
    let saved;
    await runTransaction(db, async transaction => {
      const snapshot = await transaction.get(stateRef);
      if (!snapshot.exists()) throw new Error("Tracker data has not been migrated yet.");
      const state = canonicalizeStateNames(clone(snapshot.data()));
      const existing = state.designers.find(item => item.name.toLowerCase() === name.toLowerCase());
      saved = existing || {
        id: Math.max(0, ...state.designers.map(item => Number(item.id) || 0)) + 1,
        name,
        tracker_code: null
      };
      if (!existing) state.designers.push(saved);
      state.designers.sort((a, b) => a.name.localeCompare(b.name));
      transaction.set(stateRef, { ...state, updatedAt: serverTimestamp() });
    });
    return { designer: saved };
  }

  if (path === "/api/teams" && method === "POST") {
    let teams;
    await runTransaction(db, async transaction => {
      const snapshot = await transaction.get(stateRef);
      if (!snapshot.exists()) throw new Error("Tracker data has not been migrated yet.");
      const state = canonicalizeStateNames(clone(snapshot.data()));
      ensureTeams(state);
      const action = String(body.action || "");
      const teamId = String(body.teamId || "");
      const team = state.teams.find(item => item.id === teamId);
      if (action === "add") {
        const name = cleanName(body.name);
        if (name.length < 2 || name.length > 60) throw new Error("Enter a valid team name.");
        state.teams.push({ id: "team-" + crypto.randomUUID(), name, designer_ids: [] });
      } else if (action === "rename") {
        if (!team) throw new Error("Team not found.");
        const name = cleanName(body.name);
        if (name.length < 2 || name.length > 60) throw new Error("Enter a valid team name.");
        team.name = name;
      } else if (action === "move") {
        const designerId = Number(body.designerId);
        if (!state.designers.some(item => Number(item.id) === designerId)) throw new Error("Designer not found.");
        for (const item of state.teams) item.designer_ids = (item.designer_ids || []).filter(id => Number(id) !== designerId);
        if (teamId) {
          if (!team) throw new Error("Team not found.");
          team.designer_ids.push(designerId);
        }
      } else if (action === "delete") {
        if (!team) throw new Error("Team not found.");
        state.teams = state.teams.filter(item => item.id !== teamId);
      } else {
        throw new Error("Unknown team action.");
      }
      teams = clone(state.teams);
      transaction.set(stateRef, { ...state, updatedAt: serverTimestamp() });
    });
    return { ok: true, teams };
  }

  if (path === "/api/holidays" && method === "POST") {
    const designer = cleanName(body.designer);
    const date = String(body.date || "");
    const holiday = Boolean(body.holiday);
    if (!validDate(date)) throw new Error("Invalid holiday date.");
    await runTransaction(db, async transaction => {
      const snapshot = await transaction.get(stateRef);
      if (!snapshot.exists()) throw new Error("Tracker data has not been migrated yet.");
      const state = canonicalizeStateNames(clone(snapshot.data()));
      const matched = state.designers.find(item => item.name.toLowerCase() === designer.toLowerCase());
      if (!matched) throw new Error("Designer not found.");
      state.holidays = (state.holidays || []).filter(item =>
        !(item.designer.toLowerCase() === matched.name.toLowerCase() && item.date === date)
      );
      if (holiday) state.holidays.push({ designer: matched.name, date });
      state.holidays.sort((a, b) => a.date.localeCompare(b.date) || a.designer.localeCompare(b.designer));
      transaction.set(stateRef, { ...state, updatedAt: serverTimestamp() });
    });
    return { ok: true, holiday };
  }

  if (path === "/api/import-batch" && method === "POST") {
    const sourceFile = String(body.sourceFile || "Excel import").slice(0, 255);
    const weeks = Array.isArray(body.weeks) ? body.weeks.slice(0, 60) : [];
    if (!weeks.length) throw new Error("No weekly data was supplied.");
    let response;
    await runTransaction(db, async transaction => {
      const snapshot = await transaction.get(stateRef);
      if (!snapshot.exists()) throw new Error("Tracker data has not been migrated yet.");
      const state = canonicalizeStateNames(clone(snapshot.data()));
      const byName = new Map(state.designers.map(item => [item.name.toLowerCase(), item]));
      const imported = [], skipped = [], logs = [];
      const now = new Date().toISOString();
      for (const group of weeks) {
        const week = String(group?.week || "");
        const rows = Array.isArray(group?.records) ? group.records.slice(0, 100) : [];
        if (!validDate(week) || new Date(`${week}T12:00:00Z`).getUTCDay() !== 1 || !rows.length) continue;
        let rowsImported = 0;
        for (const row of rows) {
          const name = canonicalDesignerName(row?.name);
          const hours = Number(row?.hours);
          if (!name || !Number.isFinite(hours) || hours < 0 || hours > 1000) continue;
          const designer = byName.get(name.toLowerCase());
          if (!designer) { skipped.push(name); continue; }
          state.hours = state.hours.filter(item =>
            !(item.designer.toLowerCase() === designer.name.toLowerCase() && item.week === week)
          );
          state.hours.push({ designer: designer.name, week, hours, source_file: sourceFile, imported_at: now });
          imported.push({ week, name: designer.name, hours });
          rowsImported++;
        }
        logs.push({ week, rows_imported: rowsImported });
      }
      state.hours.sort((a, b) => a.week.localeCompare(b.week) || a.designer.localeCompare(b.designer));
      let nextId = Math.max(0, ...(state.importLog || []).map(item => Number(item.id) || 0)) + 1;
      const newLogs = logs.slice().reverse().map(item => ({
        id: nextId++,
        week: item.week,
        source_file: sourceFile,
        imported_at: now,
        rows_imported: item.rows_imported
      }));
      state.importLog = [...newLogs, ...(state.importLog || [])].slice(0, 50);
      transaction.set(stateRef, { ...state, updatedAt: serverTimestamp() });
      response = { imported, skipped, weeks: logs.length };
    });
    return response;
  }

  if (path === "/api/import" && method === "POST") {
    const week = String(body.week || "");
    const sourceFile = String(body.sourceFile || "Excel import").slice(0, 255);
    const records = Array.isArray(body.records) ? body.records.slice(0, 100) : [];
    if (!validDate(week) || new Date(`${week}T12:00:00Z`).getUTCDay() !== 1) {
      throw new Error("Week commencing date must be a Monday.");
    }
    if (!records.length) throw new Error("No designer totals were supplied.");
    let response;
    await runTransaction(db, async transaction => {
      const snapshot = await transaction.get(stateRef);
      if (!snapshot.exists()) throw new Error("Tracker data has not been migrated yet.");
      const state = canonicalizeStateNames(clone(snapshot.data()));
      const byName = new Map(state.designers.map(item => [item.name.toLowerCase(), item]));
      const imported = [];
      const skipped = [];
      const now = new Date().toISOString();
      for (const row of records) {
        const name = cleanName(row?.name);
        const hours = Number(row?.hours);
        if (!name || !Number.isFinite(hours) || hours < 0 || hours > 1000) continue;
        const designer = byName.get(name.toLowerCase());
        if (!designer) {
          skipped.push(name);
          continue;
        }
        state.hours = state.hours.filter(item =>
          !(item.designer.toLowerCase() === designer.name.toLowerCase() && item.week === week)
        );
        state.hours.push({
          designer: designer.name,
          week,
          hours,
          source_file: sourceFile,
          imported_at: now
        });
        imported.push({ name: designer.name, hours });
      }
      state.hours.sort((a, b) => a.week.localeCompare(b.week) || a.designer.localeCompare(b.designer));
      const nextId = Math.max(0, ...(state.importLog || []).map(item => Number(item.id) || 0)) + 1;
      state.importLog = [
        { id: nextId, week, source_file: sourceFile, imported_at: now, rows_imported: imported.length },
        ...(state.importLog || [])
      ].slice(0, 50);
      transaction.set(stateRef, { ...state, updatedAt: serverTimestamp() });
      response = { imported, skipped };
    });
    return response;
  }

  if (path === "/api/breakdowns" && method === "POST") {
    const incoming = Array.isArray(body.records) ? body.records : null;
    if (!incoming) throw new Error("records must be an array");
    let result;
    await runTransaction(db, async transaction => {
      const snapshot = await transaction.get(breakdownRef);
      const records = snapshot.exists() ? canonicalizeBreakdownRecords(clone(snapshot.data().records || [])) : [];
      for (const raw of incoming) {
        const designer = canonicalDesignerName(raw?.designer);
        const week = cleanName(raw?.week);
        if (!designer || !validDate(week)) continue;
        const record = {
          designer,
          week,
          products: cleanCounts(raw?.products),
          statuses: cleanCounts(raw?.statuses),
          product_hours: cleanHours(raw?.product_hours),
          status_hours: cleanHours(raw?.status_hours),
          source_file: cleanName(raw?.source_file),
          imported_at: new Date().toISOString()
        };
        const index = records.findIndex(item =>
          item.week === record.week && item.designer.toLowerCase() === record.designer.toLowerCase()
        );
        if (index >= 0) records[index] = record;
        else records.push(record);
      }
      records.sort((a, b) => a.week.localeCompare(b.week) || a.designer.localeCompare(b.designer));
      transaction.set(breakdownRef, { records, updatedAt: serverTimestamp() });
      result = { ok: true, saved: incoming.length, records };
    });
    return result;
  }

  throw new Error("Unsupported tracker operation.");
};

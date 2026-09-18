import { initializeApp } from "https://www.gstatic.com/firebasejs/12.3.0/firebase-app.js";
import { getAuth, signInAnonymously } from "https://www.gstatic.com/firebasejs/12.3.0/firebase-auth.js";
import {
  getFirestore,
  doc,
  getDoc,
  runTransaction,
  serverTimestamp
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

window.firebaseReady = signInAnonymously(auth);

const clone = value => JSON.parse(JSON.stringify(value));
const cleanName = value => String(value || "").replace(/\s+/g, " ").trim();
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

async function readRequired(ref, label) {
  const snapshot = await getDoc(ref);
  if (!snapshot.exists()) throw new Error(`${label} has not been migrated to Firebase yet.`);
  return snapshot.data();
}

window.firebaseApi = async function firebaseApi(path, options = {}) {
  await window.firebaseReady;
  const method = String(options.method || "GET").toUpperCase();

  if (path === "/api/data" && method === "GET") {
    const data = await readRequired(stateRef, "Tracker data");
    const result = clone(data);
    ensureTeams(result);
    delete result.updatedAt;
    return result;
  }

  if (path === "/api/auth" && method === "POST") return { ok: true };

  if (path === "/api/breakdowns" && method === "GET") {
    const snapshot = await getDoc(breakdownRef);
    return snapshot.exists() ? { records: clone(snapshot.data().records || []) } : { records: [] };
  }

  const body = parseBody(options);

  if (path === "/api/designers" && method === "POST") {
    const name = cleanName(body.name);
    if (name.length < 3 || name.length > 80) throw new Error("Enter a valid designer name.");
    let saved;
    await runTransaction(db, async transaction => {
      const snapshot = await transaction.get(stateRef);
      if (!snapshot.exists()) throw new Error("Tracker data has not been migrated yet.");
      const state = clone(snapshot.data());
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
      const state = clone(snapshot.data());
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
      const state = clone(snapshot.data());
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
      const state = clone(snapshot.data());
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
      const records = snapshot.exists() ? clone(snapshot.data().records || []) : [];
      for (const raw of incoming) {
        const designer = cleanName(raw?.designer);
        const week = cleanName(raw?.week);
        if (!designer || !validDate(week)) continue;
        const record = {
          designer,
          week,
          products: cleanCounts(raw?.products),
          statuses: cleanCounts(raw?.statuses),
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

import type { Context, Config } from "@netlify/functions";
import { getStore, getDeployStore } from "@netlify/blobs";
import { seedState } from "./_shared/seed";

type Designer = { id: number; name: string; tracker_code?: string | null };
type Hour = { designer: string; week: string; hours: number; source_file: string; imported_at: string };
type Holiday = { designer: string; date: string };
type ImportLog = { id: number; week: string; source_file: string; imported_at: string; rows_imported: number };
type Team = { id: string; name: string; designer_ids: number[] };
type State = { designers: Designer[]; hours: Hour[]; holidays: Holiday[]; holidayReady: boolean; importLog: ImportLog[]; teams: Team[] };

function storeFor(_context: Context) {
  const deployContext = (globalThis as any).Netlify?.context?.deploy?.context;
  return deployContext === "production"
    ? getStore("design-hours-tracker", { consistency: "strong" })
    : getDeployStore("design-hours-tracker");
}

function defaultTeams(): Team[] {
  return [
    { id: "my-team", name: "My Team", designer_ids: [] },
    { id: "lynseys-team", name: "Lynsey's Team", designer_ids: [] },
    { id: "abbies-team", name: "Abbie's Team", designer_ids: [] }
  ];
}
function cloneSeed(): State {
  const state = JSON.parse(JSON.stringify(seedState)) as State;
  state.teams = defaultTeams();
  return state;
}
function ensureTeams(state: State) {
  if (!Array.isArray(state.teams) || !state.teams.length) {
    state.teams = defaultTeams();
    return true;
  }
  let changed = false;
  const validIds = new Set(state.designers.map(d => d.id));
  for (const team of state.teams) {
    const cleanIds = [...new Set((team.designer_ids || []).map(Number).filter(id => validIds.has(id)))];
    if (cleanIds.length !== (team.designer_ids || []).length) changed = true;
    team.designer_ids = cleanIds;
  }
  return changed;
}

function migrateCanonicalNames(state: State) {
  let changed = false;
  const oldName = "Jon Wilson";
  const newName = "Jonathan Wilson";
  const oldDesigner = state.designers.find(d => d.name.toLowerCase() === oldName.toLowerCase());
  const newDesigner = state.designers.find(d => d.name.toLowerCase() === newName.toLowerCase());

  if (oldDesigner) {
    if (newDesigner && newDesigner !== oldDesigner) {
      if (!newDesigner.tracker_code && oldDesigner.tracker_code) newDesigner.tracker_code = oldDesigner.tracker_code;
      state.designers = state.designers.filter(d => d !== oldDesigner);
    } else {
      oldDesigner.name = newName;
    }
    changed = true;
  }

  for (const row of state.hours) {
    if (row.designer.toLowerCase() === oldName.toLowerCase()) { row.designer = newName; changed = true; }
  }
  for (const row of state.holidays || []) {
    if (row.designer.toLowerCase() === oldName.toLowerCase()) { row.designer = newName; changed = true; }
  }
  state.designers.sort((a,b) => a.name.localeCompare(b.name));
  return changed;
}

async function getState(context: Context): Promise<State> {
  const store = storeFor(context);
  const existing = await store.get("state", { type: "json" }) as State | null;
  if (existing) {
    const state = { ...existing, holidays: existing.holidays || [], holidayReady: true, importLog: existing.importLog || [], teams: existing.teams || [] } as State;
    const canonicalChanged = migrateCanonicalNames(state);
    const teamsChanged = ensureTeams(state);
    if (canonicalChanged || teamsChanged) await store.setJSON("state", state);
    return state;
  }
  const initial = cloneSeed();
  await store.setJSON("state", initial);
  return initial;
}

async function saveState(context: Context, state: State) {
  await storeFor(context).setJSON("state", state);
}

function json(data: unknown, status = 200) {
  return new Response(JSON.stringify(data), { status, headers: { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" } });
}

function cleanName(value: unknown) { return String(value || "").replace(/\s+/g, " ").trim(); }
function validDate(value: string) { return /^\d{4}-\d{2}-\d{2}$/.test(value); }

export default async (req: Request, context: Context) => {
  try {
    const path = new URL(req.url).pathname;

    if (path === "/api/data" && req.method === "GET") {
      return json(await getState(context));
    }

    if (path === "/api/health" && req.method === "GET") {
      const state = await getState(context);
      return json({ ok: true, storage: "netlify-blobs", designers: state.designers.length, hours: state.hours.length });
    }

    if (path === "/api/auth" && req.method === "POST") {
      return json({ ok: true });
    }

    if (!["/api/designers", "/api/import", "/api/holidays", "/api/teams"].includes(path)) return json({ error: "Not found" }, 404);
    if (req.method !== "POST") return json({ error: "Method not allowed" }, 405);

    const body = await req.json().catch(() => ({})) as any;
    const state = await getState(context);

    if (path === "/api/designers") {
      const name = cleanName(body.name);
      if (name.length < 3 || name.length > 80) return json({ error: "Enter a valid designer name." }, 400);
      const existing = state.designers.find(d => d.name.toLowerCase() === name.toLowerCase());
      if (!existing) state.designers.push({ id: Math.max(0, ...state.designers.map(d => d.id || 0)) + 1, name, tracker_code: null });
      state.designers.sort((a,b) => a.name.localeCompare(b.name));
      await saveState(context, state);
      return json({ designer: existing || state.designers.find(d => d.name.toLowerCase() === name.toLowerCase()) });
    }

    if (path === "/api/teams") {
      const action = String(body.action || "");
      if (action === "add") {
        const name = cleanName(body.name);
        if (name.length < 2 || name.length > 60) return json({ error: "Enter a valid team name." }, 400);
        const id = `team-${crypto.randomUUID()}`;
        state.teams.push({ id, name, designer_ids: [] });
        await saveState(context, state);
        return json({ ok: true, teams: state.teams });
      }
      const teamId = String(body.teamId || "");
      const team = state.teams.find(t => t.id === teamId);
      if (action === "rename") {
        if (!team) return json({ error: "Team not found." }, 404);
        const name = cleanName(body.name);
        if (name.length < 2 || name.length > 60) return json({ error: "Enter a valid team name." }, 400);
        team.name = name;
        await saveState(context, state);
        return json({ ok: true, teams: state.teams });
      }
      if (action === "move") {
        const designerId = Number(body.designerId);
        if (!state.designers.some(d => d.id === designerId)) return json({ error: "Designer not found." }, 404);
        for (const t of state.teams) t.designer_ids = (t.designer_ids || []).filter(id => id !== designerId);
        if (teamId) {
          if (!team) return json({ error: "Team not found." }, 404);
          team.designer_ids.push(designerId);
        }
        await saveState(context, state);
        return json({ ok: true, teams: state.teams });
      }
      if (action === "delete") {
        if (!team) return json({ error: "Team not found." }, 404);
        state.teams = state.teams.filter(t => t.id !== teamId);
        await saveState(context, state);
        return json({ ok: true, teams: state.teams });
      }
      return json({ error: "Unknown team action." }, 400);
    }

    if (path === "/api/holidays") {
      const designer = cleanName(body.designer);
      const date = String(body.date || "");
      const holiday = !!body.holiday;
      if (!state.designers.some(d => d.name.toLowerCase() === designer.toLowerCase())) return json({ error: "Designer not found." }, 404);
      if (!validDate(date)) return json({ error: "Invalid holiday date." }, 400);
      state.holidays = state.holidays.filter(h => !(h.designer.toLowerCase() === designer.toLowerCase() && h.date === date));
      if (holiday) state.holidays.push({ designer: state.designers.find(d => d.name.toLowerCase() === designer.toLowerCase())!.name, date });
      state.holidays.sort((a,b) => a.date.localeCompare(b.date) || a.designer.localeCompare(b.designer));
      await saveState(context, state);
      return json({ ok: true, holiday });
    }

    const week = String(body.week || "");
    const sourceFile = String(body.sourceFile || "Excel import").slice(0,255);
    const records = Array.isArray(body.records) ? body.records.slice(0,100) : [];
    if (!validDate(week)) return json({ error: "Invalid week commencing date." }, 400);
    if (new Date(`${week}T12:00:00Z`).getUTCDay() !== 1) return json({ error: "Week commencing date must be a Monday." }, 400);
    if (!records.length) return json({ error: "No designer totals were supplied." }, 400);

    const byName = new Map(state.designers.map(d => [d.name.toLowerCase(), d]));
    const imported: {name:string;hours:number}[] = [], skipped: string[] = [];
    const now = new Date().toISOString();
    for (const r of records) {
      const name = cleanName(r?.name), hours = Number(r?.hours);
      if (!name || !Number.isFinite(hours) || hours < 0 || hours > 1000) continue;
      const designer = byName.get(name.toLowerCase());
      if (!designer) { skipped.push(name); continue; }
      state.hours = state.hours.filter(h => !(h.designer.toLowerCase() === designer.name.toLowerCase() && h.week === week));
      state.hours.push({ designer: designer.name, week, hours, source_file: sourceFile, imported_at: now });
      imported.push({ name: designer.name, hours });
    }
    state.hours.sort((a,b) => a.week.localeCompare(b.week) || a.designer.localeCompare(b.designer));
    const nextId = Math.max(0, ...state.importLog.map(x => x.id || 0)) + 1;
    state.importLog.unshift({ id: nextId, week, source_file: sourceFile, imported_at: now, rows_imported: imported.length });
    state.importLog = state.importLog.slice(0,50);
    await saveState(context, state);
    return json({ imported, skipped });
  } catch (error: any) {
    console.error(error);
    return json({ error: error?.message || "Unexpected server error." }, 500);
  }
};

export const config: Config = {
  path: ["/api/data", "/api/health", "/api/auth", "/api/designers", "/api/import", "/api/holidays", "/api/teams"]
};

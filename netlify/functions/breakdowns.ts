import type { Config } from "@netlify/functions";
import { getStore, getDeployStore } from "@netlify/blobs";

type Counts = Record<string, number>;
type BreakdownRecord = { designer: string; week: string; products: Counts; statuses: Counts; product_hours?: Counts; status_hours?: Counts; source_file?: string; imported_at?: string };
type BreakdownState = { records: BreakdownRecord[] };

function json(data: unknown, status = 200) {
  return new Response(JSON.stringify(data), { status, headers: { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" } });
}
function clean(value: unknown) { return String(value ?? "").replace(/\s+/g, " ").trim(); }
function validWeek(value: string) { return /^\d{4}-\d{2}-\d{2}$/.test(value); }
function cleanCounts(value: unknown): Counts {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  const out: Counts = {};
  for (const [key, raw] of Object.entries(value as Record<string, unknown>)) {
    const label = clean(key); const count = Number(raw);
    if (label && Number.isFinite(count) && count >= 0) out[label] = Math.round(count);
  }
  return out;
}
function cleanHours(value: unknown): Counts {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  const out: Counts = {};
  for (const [key, raw] of Object.entries(value as Record<string, unknown>)) {
    const label = clean(key); const hours = Number(raw);
    if (label && Number.isFinite(hours) && hours >= 0) out[label] = Math.round(hours * 100) / 100;
  }
  return out;
}
function storeForContext() {
  const production = Netlify.context?.deploy?.context === "production";
  return production ? getStore("design-hours-breakdowns", { consistency: "strong" }) : getDeployStore("design-hours-breakdowns", { consistency: "strong" });
}
async function readState(): Promise<BreakdownState> {
  const store = storeForContext();
  const stored = await store.get("state-v1", { type: "json" }) as BreakdownState | null;
  const records = Array.isArray(stored?.records) ? stored.records : [];
  let changed = false;
  for (const record of records) {
    if (record.designer.toLowerCase() === "jon wilson") {
      record.designer = "Jonathan Wilson";
      changed = true;
    }
  }
  if (changed) await store.setJSON("state-v1", { records });
  return { records };
}

export default async (req: Request) => {
  if (req.method === "GET") return json(await readState());
  if (req.method !== "POST") return json({ error: "Method not allowed" }, 405);
  const body = await req.json().catch(() => ({})) as { records?: unknown[] };
  if (!Array.isArray(body.records)) return json({ error: "records must be an array" }, 400);
  const incoming: BreakdownRecord[] = [];
  for (const raw of body.records as any[]) {
    const designer = clean(raw?.designer), week = clean(raw?.week);
    if (!designer || !validWeek(week)) continue;
    incoming.push({
      designer,
      week,
      products: cleanCounts(raw?.products),
      statuses: cleanCounts(raw?.statuses),
      product_hours: cleanHours(raw?.product_hours),
      status_hours: cleanHours(raw?.status_hours),
      source_file: clean(raw?.source_file),
      imported_at: new Date().toISOString()
    });
  }
  const state = await readState();
  for (const record of incoming) {
    const idx = state.records.findIndex(x => x.week === record.week && x.designer.toLowerCase() === record.designer.toLowerCase());
    if (idx >= 0) state.records[idx] = record; else state.records.push(record);
  }
  state.records.sort((a,b) => a.week.localeCompare(b.week) || a.designer.localeCompare(b.designer));
  await storeForContext().setJSON("state-v1", state);
  return json({ ok: true, saved: incoming.length, records: state.records });
};

export const config: Config = { path: "/api/breakdowns" };

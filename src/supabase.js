import { createClient } from "@supabase/supabase-js";
import { SUPABASE_URL, SUPABASE_ANON_KEY } from "./config.js";

export { SUPABASE_URL };

// 只有兩個值都填咗先建立 client；否則 cloudEnabled = false，app 用本機儲存
export const cloudEnabled = Boolean(SUPABASE_URL && SUPABASE_ANON_KEY);
export const supa = cloudEnabled ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY) : null;

// 2026-11 拆分：資料改為逐筆記錄存喺 app_records（每位教練、每個預約時段、每筆記錄各自一行）。
// app_state.main 變成由資料庫 trigger 自動砌出嚟嘅快照（俾日曆 feed 用），app 唔會再寫佢。
const RECORDS = "app_records";
const BACKUPS = "app_backups";
const PAGE = 1000;   // PostgREST 每次最多回傳 1000 行，要分頁
const UPSERT_CHUNK = 400;
const DELETE_CHUNK = 100;

// 讀取全部記錄。任何一頁失敗都當成「讀取失敗」（{ ok:false }），絕不可以當成雲端空白。
export async function loadAllRecords() {
  if (!supa) return { ok: false, rows: [] };
  const rows = [];
  try {
    for (let from = 0; ; from += PAGE) {
      const { data, error } = await supa.from(RECORDS).select("collection,rid,data")
        .order("collection", { ascending: true }).order("rid", { ascending: true })
        .range(from, from + PAGE - 1);
      if (error) { console.error("loadAllRecords", error); return { ok: false, rows: [] }; }
      rows.push(...(data || []));
      if (!data || data.length < PAGE) break;
    }
    return { ok: true, rows };
  } catch (e) { console.error("loadAllRecords", e); return { ok: false, rows: [] }; }
}

// 新增／更新記錄（分批）。回傳成功寫入咗嘅記錄（失敗嗰批之後嘅唔算）。
export async function upsertRecords(rows) {
  const done = [];
  if (!supa) return { ok: false, done };
  const now = new Date().toISOString();
  for (let i = 0; i < rows.length; i += UPSERT_CHUNK) {
    const batch = rows.slice(i, i + UPSERT_CHUNK);
    try {
      const { error } = await supa.from(RECORDS)
        .upsert(batch.map((r) => ({ collection: r.collection, rid: r.rid, data: r.data, updated_at: now })), { onConflict: "collection,rid" });
      if (error) { console.error("upsertRecords", error); return { ok: false, done }; }
    } catch (e) { console.error("upsertRecords", e); return { ok: false, done }; }
    done.push(...batch);
  }
  return { ok: true, done };
}

// 刪除記錄（按 collection 分組、分批）
export async function deleteRecords(rows) {
  const done = [];
  if (!supa) return { ok: false, done };
  const byCol = {};
  rows.forEach((r) => { (byCol[r.collection] = byCol[r.collection] || []).push(r); });
  for (const col of Object.keys(byCol)) {
    const list = byCol[col];
    for (let i = 0; i < list.length; i += DELETE_CHUNK) {
      const batch = list.slice(i, i + DELETE_CHUNK);
      try {
        const { error } = await supa.from(RECORDS).delete().eq("collection", col).in("rid", batch.map((r) => r.rid));
        if (error) { console.error("deleteRecords", error); return { ok: false, done }; }
      } catch (e) { console.error("deleteRecords", e); return { ok: false, done }; }
      done.push(...batch);
    }
  }
  return { ok: true, done };
}

// 訂閱其他裝置嘅改動。短時間內嘅多個改動會一齊交出（80ms）。
// onChanges([{ type: "upsert" | "delete", collection, rid, data }])；onStatus("SUBSCRIBED" | "CHANNEL_ERROR" | ...)
export function subscribeRecords(onChanges, onStatus) {
  if (!supa) return () => {};
  let queue = [], timer = null;
  const flush = () => { const q = queue; queue = []; timer = null; if (q.length) onChanges(q); };
  const ch = supa
    .channel("app_records_changes")
    .on("postgres_changes", { event: "*", schema: "public", table: RECORDS }, (payload) => {
      if (payload.eventType === "DELETE") {
        const o = payload.old || {};
        if (o.collection !== undefined && o.rid !== undefined) queue.push({ type: "delete", collection: o.collection, rid: o.rid });
      } else {
        const n = payload.new || {};
        if (n.collection !== undefined && n.rid !== undefined) queue.push({ type: "upsert", collection: n.collection, rid: n.rid, data: n.data });
      }
      if (!timer) timer = setTimeout(flush, 80);
    })
    .subscribe((status) => { if (onStatus) onStatus(status); });
  return () => { try { if (timer) clearTimeout(timer); supa.removeChannel(ch); } catch (e) { /* ignore */ } };
}

// 每日自動備份：id = backup-0 … backup-6（按星期幾輪流覆蓋，保留最近7日）。
// 同一日已經備份過就唔再上載（慳手機流量）；資料係空白嘅就唔好叫呢個 function。
export async function cloudSaveDailyBackup(payload) {
  if (!supa) return false;
  try {
    const id = "backup-" + new Date().getDay();
    const today = new Date().toDateString();
    const { data: existing } = await supa.from(BACKUPS).select("updated_at").eq("id", id).maybeSingle();
    if (existing && new Date(existing.updated_at).toDateString() === today) return true;
    const { error } = await supa.from(BACKUPS).upsert({ id, data: payload, updated_at: new Date().toISOString() });
    if (error) { console.error("cloudSaveDailyBackup", error); return false; }
    return true;
  } catch (e) { return false; }
}

// 列出雲端自動備份（id, data, updated_at），最新排最先
export async function cloudListBackups() {
  if (!supa) return [];
  try {
    const { data, error } = await supa.from(BACKUPS).select("id,data,updated_at");
    if (error) { console.error("cloudListBackups", error); return []; }
    return (data || []).sort((a, b) => String(b.updated_at).localeCompare(String(a.updated_at)));
  } catch (e) { return []; }
}

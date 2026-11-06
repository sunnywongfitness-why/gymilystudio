import { createClient } from "@supabase/supabase-js";
import { SUPABASE_URL, SUPABASE_ANON_KEY } from "./config.js";

export { SUPABASE_URL };

// 只有兩個值都填咗先建立 client；否則 cloudEnabled = false，app 用本機儲存
export const cloudEnabled = Boolean(SUPABASE_URL && SUPABASE_ANON_KEY);
export const supa = cloudEnabled ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY) : null;

const ROW_ID = "main";

// 讀取雲端資料（回傳 data 物件，或 null）
export async function cloudLoad() {
  if (!supa) return null;
  const { data, error } = await supa.from("app_state").select("data").eq("id", ROW_ID).maybeSingle();
  if (error) { console.error("cloudLoad", error); return null; }
  return data ? data.data : null;
}

// 讀取雲端資料，並分清「讀取出錯」同「真係冇資料」：
//   { ok:false }            → 讀取失敗（斷網、專案暫停、權限等），絕不可以當作雲端空白
//   { ok:true, data:null }  → 讀取成功，但雲端真係未有呢個 row（首次使用）
//   { ok:true, data:{...} } → 正常
export async function cloudLoadStatus() {
  if (!supa) return { ok: false, data: null };
  try {
    const { data, error } = await supa.from("app_state").select("data").eq("id", ROW_ID).maybeSingle();
    if (error) { console.error("cloudLoad", error); return { ok: false, data: null }; }
    return { ok: true, data: data ? data.data : null };
  } catch (e) { console.error("cloudLoad", e); return { ok: false, data: null }; }
}

// 每日自動備份：喺同一個 table 用 id = backup-0 … backup-6（按星期幾輪流覆蓋，永遠保留最近7日）。
// 只會喺「資料唔係空白」嗰陣先寫，避免空白資料洗走好嘅備份。失敗唔影響正常運作。
export async function cloudSaveBackup(payload) {
  if (!supa) return false;
  try {
    const id = "backup-" + new Date().getDay();
    const { error } = await supa.from("app_state").upsert({ id, data: payload, updated_at: new Date().toISOString() });
    if (error) { console.error("cloudSaveBackup", error); return false; }
    return true;
  } catch (e) { return false; }
}

// 列出雲端自動備份（id, data, updated_at），最新排最先
export async function cloudListBackups() {
  if (!supa) return [];
  try {
    const { data, error } = await supa.from("app_state").select("id,data,updated_at").like("id", "backup-%");
    if (error) { console.error("cloudListBackups", error); return []; }
    return (data || []).sort((a, b) => String(b.updated_at).localeCompare(String(a.updated_at)));
  } catch (e) { return []; }
}

// 寫入雲端資料（整份覆蓋）
export async function cloudSave(payload) {
  if (!supa) return false;
  const { error } = await supa.from("app_state").upsert({ id: ROW_ID, data: payload, updated_at: new Date().toISOString() });
  if (error) { console.error("cloudSave", error); return false; }
  return true;
}

// 訂閱雲端變更（其他裝置改咗會即時收到）
export function cloudSubscribe(onChange) {
  if (!supa) return () => {};
  const ch = supa
    .channel("app_state_changes")
    .on("postgres_changes", { event: "*", schema: "public", table: "app_state", filter: `id=eq.${ROW_ID}` },
      (payload) => { if (payload.new && payload.new.data) onChange(payload.new.data); })
    .subscribe();
  return () => { try { supa.removeChannel(ch); } catch (e) { /* ignore */ } };
}

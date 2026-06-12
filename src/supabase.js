import { createClient } from "@supabase/supabase-js";
import { SUPABASE_URL, SUPABASE_ANON_KEY } from "./config.js";

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

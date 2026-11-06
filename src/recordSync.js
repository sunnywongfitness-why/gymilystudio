// ===== 逐筆記錄同步（2026-11 拆分）=====
// 以前成個 app 嘅資料放喺 Supabase 同一行（app_state.main），每次改動都成份覆蓋——一次錯誤寫入就會冚走晒所有資料。
// 而家改做 app_records 表：每位教練、每個預約時段、每筆流水帳／支出……各自一行。
// app 入面嘅資料結構（bundle）完全唔變，呢個檔案只負責「bundle ⇄ 一行行記錄」嘅轉換、比較同合併，全部係純函數，方便測試。
//
// 記錄格式：{ collection, rid, data }
//   collection "s"        → 單一設定值（例如 adminPassword），rid = key
//   collection "k"        → 每個 list / dict 嘅登記：{ kind:"list", order:[rid...] } 或 { kind:"dict" }
//   collection "l:<key>"  → list 入面每一項，rid = "i:<id>"（有 id）或 "h:<內容hash>"（冇 id）
//   collection "d:<key>"  → dict 入面每一個 entry，rid = 原本嘅 key（例如預約時段 "2026-11-06_15:30"）
// 同 helpers.js 嘅 stableStringify 完全一樣（object key 排序後先 stringify），抄一份喺度令呢個檔案唔使依賴其他模組
export function stableStringify(obj) {
  if (obj === null || typeof obj !== "object") return JSON.stringify(obj);
  if (Array.isArray(obj)) return "[" + obj.map(stableStringify).join(",") + "]";
  const keys = Object.keys(obj).sort();
  return "{" + keys.map((k) => JSON.stringify(k) + ":" + stableStringify(obj[k])).join(",") + "}";
}

export const rowKey = (collection, rid) => collection + "\u0000" + rid;

// cyrb53：細、快、足夠分散嘅 53-bit hash，用嚟幫冇 id 嘅記錄（例如取消記錄）整一個穩定 rid
function cyrb53(str, seed = 0) {
  let h1 = 0xdeadbeef ^ seed, h2 = 0x41c6ce57 ^ seed;
  for (let i = 0; i < str.length; i++) {
    const ch = str.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
}

// list 每一項嘅 rid：有 id 用 id；冇 id 用內容 hash。逐項獨立決定，所以新增一項唔會令其他項嘅 rid 改變。
// 重複（同 id 或者內容完全一樣）就加 #2、#3……
export function listRids(arr) {
  const seen = {};
  return arr.map((item) => {
    const base = item && typeof item === "object" && !Array.isArray(item) && item.id !== undefined && item.id !== null
      ? "i:" + String(item.id)
      : "h:" + cyrb53(stableStringify(item));
    seen[base] = (seen[base] || 0) + 1;
    return seen[base] > 1 ? base + "#" + seen[base] : base;
  });
}

const isPlainObject = (v) => v !== null && typeof v === "object" && !Array.isArray(v);

// bundle → Map(rowKey → { collection, rid, data, json })
export function explode(bundle) {
  const map = new Map();
  const put = (collection, rid, data) => map.set(rowKey(collection, rid), { collection, rid, data, json: stableStringify(data) });
  Object.keys(bundle).forEach((key) => {
    const v = bundle[key];
    if (v === undefined) return;
    if (Array.isArray(v)) {
      const rids = listRids(v);
      v.forEach((item, i) => put("l:" + key, rids[i], item));
      put("k", key, { kind: "list", order: rids });
    } else if (isPlainObject(v)) {
      Object.keys(v).forEach((k) => { if (v[k] !== undefined) put("d:" + key, k, v[k]); });
      put("k", key, { kind: "dict" });
    } else {
      put("s", key, v);
    }
  });
  return map;
}

// 一行行記錄 → bundle。只砌 keys 入面指定嘅 top-level key（唔畀就全部）。
// list：按登記嘅 order 排；有記錄但唔喺 order 入面（兩部機同時新增造成）就放最前，按 rid 排，同資料庫快照嘅排法一致。
export function implode(map, keys) {
  const regs = {}, settings = {}, lists = {}, dicts = {};
  map.forEach((r) => {
    if (r.collection === "k") regs[r.rid] = r.data;
    else if (r.collection === "s") settings[r.rid] = r.data;
    else if (r.collection.startsWith("l:")) { const k = r.collection.slice(2); (lists[k] = lists[k] || {})[r.rid] = r.data; }
    else if (r.collection.startsWith("d:")) { const k = r.collection.slice(2); (dicts[k] = dicts[k] || {})[r.rid] = r.data; }
  });
  const want = keys ? new Set(keys) : null;
  const out = {};
  Object.keys(settings).forEach((k) => { if (!want || want.has(k)) out[k] = settings[k]; });
  Object.keys(regs).forEach((k) => {
    if (want && !want.has(k)) return;
    const reg = regs[k] || {};
    if (reg.kind === "list") {
      const items = lists[k] || {};
      const order = Array.isArray(reg.order) ? reg.order : [];
      const inOrder = new Set(order);
      const extra = Object.keys(items).filter((rid) => !inOrder.has(rid)).sort();
      out[k] = [...extra.map((rid) => items[rid]), ...order.filter((rid) => rid in items).map((rid) => items[rid])];
    } else if (reg.kind === "dict") {
      const entries = dicts[k] || {};
      const o = {};
      Object.keys(entries).sort().forEach((rid) => { o[rid] = entries[rid]; });
      out[k] = o;
    }
  });
  return out;
}

// 由記錄嘅 collection 推返佢屬於邊個 top-level key
export const topKeyOf = (collection, rid) => (collection === "k" || collection === "s" ? rid : collection.slice(2));

// 本機 vs 上次同步：要上載邊啲、要刪邊啲
export function diff(localMap, syncedMap) {
  const upserts = [], deletes = [];
  localMap.forEach((r, k) => { const s = syncedMap.get(k); if (!s || s.json !== r.json) upserts.push(r); });
  syncedMap.forEach((r, k) => { if (!localMap.has(k)) deletes.push(r); });
  return { upserts, deletes };
}

// 收到雲端（其他裝置）嘅改動：逐行 3-way 合併。
// 本機嗰行同「上次同步」一樣（即本機冇改過）→ 用雲端新版本；本機改緊（未上載）→ 保留本機。
// changes: [{ type: "upsert" | "delete", collection, rid, data }]
// 回傳 { localMap, syncedMap, touchedKeys }（全部係新 Map，唔改原本）
export function mergeRemote(localMap, syncedMap, changes) {
  const local = new Map(localMap), synced = new Map(syncedMap);
  const touched = new Set();
  changes.forEach((c) => {
    const k = rowKey(c.collection, c.rid);
    const prev = synced.get(k), mine = local.get(k);
    const localUnchanged = (prev === undefined && mine === undefined) || (prev && mine && prev.json === mine.json);
    if (c.type === "delete") {
      synced.delete(k);
      if (localUnchanged && mine !== undefined) { local.delete(k); touched.add(topKeyOf(c.collection, c.rid)); }
    } else {
      const row = { collection: c.collection, rid: c.rid, data: c.data, json: stableStringify(c.data) };
      synced.set(k, row);
      if (localUnchanged && (!mine || mine.json !== row.json)) { local.set(k, row); touched.add(topKeyOf(c.collection, c.rid)); }
    }
  });
  return { localMap: local, syncedMap: synced, touchedKeys: [...touched] };
}

// 整批雲端記錄（重新載入）→ 轉做 changes，交俾 mergeRemote（包括雲端已經冇咗嘅行）
export function fullRefreshChanges(remoteRows, syncedMap) {
  const changes = [];
  const seen = new Set();
  remoteRows.forEach((r) => {
    const k = rowKey(r.collection, r.rid);
    seen.add(k);
    const s = syncedMap.get(k);
    if (!s || s.json !== stableStringify(r.data)) changes.push({ type: "upsert", collection: r.collection, rid: r.rid, data: r.data });
  });
  syncedMap.forEach((r, k) => { if (!seen.has(k)) changes.push({ type: "delete", collection: r.collection, rid: r.rid }); });
  return changes;
}

export const rowsToMap = (rows) => {
  const m = new Map();
  rows.forEach((r) => m.set(rowKey(r.collection, r.rid), { collection: r.collection, rid: r.rid, data: r.data, json: stableStringify(r.data) }));
  return m;
};

// 安全檢查：一次過刪走大量記錄，或者將有內容嘅資料變成空白，都要用戶明確確認先准（重設／還原／清理歷史）
export function massChangeProblem(d, syncedMap) {
  const dataDeletes = d.deletes.filter((r) => r.collection !== "k" && r.collection !== "s");
  const total = [...syncedMap.values()].filter((r) => r.collection !== "k" && r.collection !== "s").length;
  if (dataDeletes.length > 30 && dataDeletes.length > total * 0.2) return `一次過刪除 ${dataDeletes.length} 筆記錄（共 ${total} 筆）`;
  return null;
}

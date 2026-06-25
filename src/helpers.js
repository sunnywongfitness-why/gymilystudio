// 純函數 helper：日期/時間/堂數計算、登入狀態、雲端同步比對等
import { LS_KEY, SESSION_KEY, CALSCALE_KEY, DEFAULT_COACHES, DEFAULT_SUBADMINS, ADMIN_TAB_KEYS, CLOSED_DAYS } from "./constants.js";
import { S } from "./styles.js";

export const loadStore = () => {
  try { return JSON.parse(localStorage.getItem(LS_KEY)) || {}; } catch (e) { return {}; }
};
export const initStore = loadStore();
export const persisted = (key, fallback) => (initStore[key] !== undefined ? initStore[key] : fallback);

// ---- 登入狀態（記住密碼）：獨立一個 key，淨係存喺呢部裝置，唔會同其他資料一齊上雲端同步 ----
export const loadSession = () => { try { return JSON.parse(localStorage.getItem(SESSION_KEY)); } catch (e) { return null; } };
export const saveSession = (s) => { try { localStorage.setItem(SESSION_KEY, JSON.stringify(s)); } catch (e) { /* ignore */ } };
export const clearSession = () => { try { localStorage.removeItem(SESSION_KEY); } catch (e) { /* ignore */ } };

// ---- 日曆縮放偏好：純粹顯示設定，淨係存喺呢部裝置，唔同步 ----
export const loadCalScale = () => { try { const v = parseFloat(localStorage.getItem(CALSCALE_KEY)); return v && v >= 0.5 && v <= 1 ? v : 1; } catch (e) { return 1; } };
export const saveCalScale = (v) => { try { localStorage.setItem(CALSCALE_KEY, String(v)); } catch (e) { /* ignore */ } };

// 穩定序列化：將物件 key 依字母排序，令同步比對唔受 Supabase jsonb 重排 key 影響
export function stableStringify(obj) {
  if (obj === null || typeof obj !== "object") return JSON.stringify(obj);
  if (Array.isArray(obj)) return "[" + obj.map(stableStringify).join(",") + "]";
  const keys = Object.keys(obj).sort();
  return "{" + keys.map((k) => JSON.stringify(k) + ":" + stableStringify(obj[k])).join(",") + "}";
}
export function resolveSession() {
  const session = loadSession();
  if (!session) return null;
  if (session.role === "admin") return { user: { id: 0, name: "管理員", role: "admin" }, view: "admin", adminTab: "overview" };
  if (session.role === "subadmin") {
    const subs = persisted("subAdmins", DEFAULT_SUBADMINS);
    const sub = subs.find((s) => s.id === session.id);
    if (!sub) { clearSession(); return null; }
    const firstAllowed = ADMIN_TAB_KEYS.find((k) => sub.permissions?.[k]);
    return { user: { ...sub, role: "subadmin" }, view: "admin", adminTab: firstAllowed || "settings" };
  }
  if (session.role === "coach") {
    const coaches = persisted("coaches", DEFAULT_COACHES);
    const coach = coaches.find((c) => c.id === session.id);
    if (!coach) { clearSession(); return null; }
    return { user: { ...coach, role: "coach" }, view: "calendar", adminTab: "overview" };
  }
  return null;
}
export const initialSession = resolveSession();
export function duoPrice(hours) { return DUO_BASE + Math.round((hours - 1) / 0.5) * DUO_HALF_HOUR_ADD; }
export const isWholeVenue = (e) => e.type === "charter" && e.charterType !== "trial";
export const rentalShort = (ct) => ct === "group" ? "小組" : ct === "trial" ? "試堂" : "包場";
export const rentalFull = (ct) => ct === "group" ? "小組訓練" : ct === "trial" ? "試堂" : "私人包場";
export const isClosedDay = (date) => CLOSED_DAYS.includes(new Date(`${date}T00:00:00`).getDay());

// 15-min grid 07:00–22:00
export function getDaysOfWeek(offset = 0) {
  const days = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  today.setDate(today.getDate() + offset);
  for (let i = 0; i < 7; i++) { const d = new Date(today); d.setDate(today.getDate() + i); days.push(d); }
  return days;
}
export const pad2 = (n) => String(n).padStart(2, "0");
export const formatDate = (date) => `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
export const isTodayDate = (date) => formatDate(date) === formatDate(new Date());
export const formatDay = (date) => `周${["日", "一", "二", "三", "四", "五", "六"][date.getDay()]}`;
export const monthKey = (dateStr) => dateStr.slice(0, 7);
export const hoursUntil = (date, time) => (new Date(`${date}T${time}:00`) - new Date()) / (1000 * 60 * 60);
export function addMinutes(time, mins) {
  const [h, m] = time.split(":").map(Number);
  const total = h * 60 + m + mins;
  return `${String(Math.floor(total / 60)).padStart(2, "0")}:${String(total % 60).padStart(2, "0")}`;
}
// list of 15-min slot keys for a booking of given hours starting at time
export function slotsFor(time, hours) {
  const n = Math.round(hours * 4);
  return Array.from({ length: n }, (_, i) => addMinutes(time, i * 15));
}
// 將 "HH:MM" 轉做由 07:00 起計嘅第幾個 15 分鐘格（方便計算跨度中點）
export function slotIndex(time) {
  const [h, m] = time.split(":").map(Number);
  return (h - 7) * 4 + m / 15;
}
// 將一個 booking 嘅顯示內容拆做幾行：教練名/類型、（學生名）、開始時間 —— 每行會分配落唔同嘅實際格仔
export function buildEntryLines(v, isTrial, coachObj, isOwner) {
  const lines = [];
  if (isTrial) {
    lines.push({ text: `試堂${v.coachName ? " · " + v.coachName : ""}`, style: S.slotNameFull });
  } else if (v.type === "charter") {
    lines.push({ text: `${rentalShort(v.charterType)}${v.coachName ? " · " + v.coachName : ""}`, style: S.slotNameFull });
  } else {
    lines.push({ text: `${coachObj?.name || ""}${v.type === "duo" ? " · 1對2" : " · 1對1"}`, style: S.slotNameFull });
    if (isOwner && v.students && v.students.length > 0) lines.push({ text: v.students.join("、"), style: S.slotStudentsFull });
  }
  lines.push({ text: v.start, style: S.slotTimeFull });
  return lines;
}

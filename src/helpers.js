// 純函數 helper：日期/時間/堂數計算、登入狀態、雲端同步比對等
import { LS_KEY, SESSION_KEY, CALSCALE_KEY, ADMIN_TAB_KEYS } from "./constants.js";
import { DEFAULT_COACHES, DEFAULT_SUBADMINS, CLOSED_DAYS, BRAND_NAME } from "./brand.js";
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
    // 有「總覽」權限就強制去總覽（同主admin睇齊，唔記住上次留低嗰個tab）；冇嗰個權限先fallback去佢第一個有權限嘅tab
    const adminTab = sub.permissions?.overview ? "overview" : (firstAllowed || "settings");
    return { user: { ...sub, role: "subadmin" }, view: "admin", adminTab };
  }
  if (session.role === "coach") {
    const coaches = persisted("coaches", DEFAULT_COACHES);
    const coach = coaches.find((c) => c.id === session.id);
    if (!coach) { clearSession(); return null; }
    // 教練每次開app（包括淨係重新載入、未重新輸入密碼嘅情況）都要見返首頁，等佢即刻知道自己有冇book堂，
    // 唔好記住佢上次留低喺邊個view——呢個係刻意行為，唔係bug（見2026-07 對話定案，方法A）
    return { user: { ...coach, role: "coach" }, view: "home", adminTab: "overview" };
  }
  return null;
}
export const initialSession = resolveSession();
export const isWholeVenue = (e) => e.type === "charter" && e.charterType !== "trial";
export const rentalShort = (ct) => ct === "group" ? "小組" : ct === "trial" ? "試堂" : ct === "clean" ? "清潔" : ct === "filming" ? "拍片" : "包場";
export const rentalFull = (ct) => ct === "group" ? "小組訓練" : ct === "trial" ? "試堂" : ct === "clean" ? "封場清潔" : ct === "filming" ? "拍片" : "私人包場";
// 第12項：將 cancelledBy/addedBy 呢類操作者標記，轉做人睇得明嘅文字（subadmin 會顯示返實際姓名）
export function actorLabel(v) {
  if (!v) return "";
  if (v === "coach") return "教練自行取消";
  if (v === "admin") return "管理員代取消";
  if (typeof v === "string" && v.startsWith("subadmin:")) return `副管理員代取消（${v.slice(9)}）`;
  return v;
}
// 落單方式：呢筆booking係教練自己book，定係admin／副管理員代book（同 actorLabel 同一套 tag 格式，但文字係「book」唔係「取消」）
export function bookedByLabel(v) {
  if (!v) return "—（舊記錄）";
  if (v === "coach") return "教練自己 book";
  if (v === "admin") return "管理員代 book";
  if (typeof v === "string" && v.startsWith("subadmin:")) return `副管理員代 book（${v.slice(9)}）`;
  return v;
}
export const isClosedDay = (date) => CLOSED_DAYS.includes(new Date(`${date}T00:00:00`).getDay());
// 動態生成休息日提示文字，根據 brand.js 嘅 CLOSED_DAYS 自動配，唔使寫死星期幾（開新公司唔使搵埋呢句改）
const DAY_NAMES = ["日", "一", "二", "三", "四", "五", "六"];
export const closedDayMessage = () => CLOSED_DAYS.length === 0 ? "" : `星期${CLOSED_DAYS.map((d) => DAY_NAMES[d]).join("、")}休息，不開放預約`;
// 教練顏色自動派：根據 coachId 用黃金角 hash 一個色相，固定飽和度/明度，可以無限擴展、相鄰ID易分辨、零人手介入
// 注意：全app好多地方用緊 c.color + "33" 呢種hex尾加透明度嘅寫法（例如格仔背景），呢個function必須輸出 #rrggbb 格式，
// 唔可以用 hsl()/rgb()，唔係會令個透明度寫法變成invalid CSS，背景就會消失變返黑色（呢個bug以前中過招，唔好再犯）。
function hslToHex(h, s, l) {
  s /= 100; l /= 100;
  const k = (n) => (n + h / 30) % 12;
  const a = s * Math.min(l, 1 - l);
  const f = (n) => l - a * Math.max(-1, Math.min(k(n) - 3, Math.min(9 - k(n), 1)));
  const toHex = (n) => Math.round(255 * f(n)).toString(16).padStart(2, "0");
  return `#${toHex(0)}${toHex(8)}${toHex(4)}`;
}
export function coachColorFromId(id) {
  const hue = (Number(id) * 137.508) % 360;
  return hslToHex(hue, 65, 60);
}

// 固定一週：星期日=第一日、星期六=第七日（唔再係「今日之後7日」嘅 rolling window）
// mode: "fixed"（固定星期一至日，offset 為 7 嘅倍數） | "rolling"（以今日做第一日，offset 為任意日數）
export function getDaysOfWeek(offset = 0, mode = "fixed") {
  const days = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  if (mode === "rolling") {
    // Rolling 模式：每次「上週/下週」都係 ±7 日，但永遠以「今日 + offset」做第一日
    const start = new Date(today);
    start.setDate(today.getDate() + offset);
    for (let i = 0; i < 7; i++) { const d = new Date(start); d.setDate(start.getDate() + i); days.push(d); }
    return days;
  }
  // Fixed 模式：固定星期一到星期日
  const monday = new Date(today);
  const dow = today.getDay(); // 0=日 1=一 ... 6=六
  const diffToMonday = dow === 0 ? -6 : 1 - dow; // 今日距返本週一嘅日數
  monday.setDate(today.getDate() + diffToMonday);
  monday.setDate(monday.getDate() + offset); // offset 已經係 weekOffset*7，指定去第幾個禮拜
  for (let i = 0; i < 7; i++) { const d = new Date(monday); d.setDate(monday.getDate() + i); days.push(d); }
  return days;
}
export const pad2 = (n) => String(n).padStart(2, "0");
export const formatDate = (date) => `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
// 「呢一刻」嘅日期時間戳（YYYY-MM-DD HH:MM），用裝置本地時間（唔係toISOString嗰種UTC，唔係唔會慢咗8個鐘）
export const nowStamp = (date = new Date()) => `${formatDate(date)} ${pad2(date.getHours())}:${pad2(date.getMinutes())}`;
export const addDaysToDate = (dateStr, days) => { const d = new Date(`${dateStr}T00:00:00`); d.setDate(d.getDate() + days); return formatDate(d); };
export const addMonthsToDate = (dateStr, months) => { const d = new Date(`${dateStr}T00:00:00`); d.setMonth(d.getMonth() + months); return formatDate(d); };
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
// ---- 財政年度（4月1日至3月31日，跟審計師/稅局個basis period，唔係calendar year）----
// 用「開始年份」表示一個財政年度：例如 2025-04-01～2026-03-31 呢個年度，開始年份係 2025
export function fiscalYearOf(dateStr) {
  const [y, m] = dateStr.split("-").map(Number);
  return m >= 4 ? y : y - 1;
}
export const fiscalYearLabel = (startYear) => `${startYear}/${String(startYear + 1).slice(2)}`;
export const fiscalYearRange = (startYear) => ({ start: `${startYear}-04-01`, end: `${startYear + 1}-03-31` });

// ---- 支出憑證編號：EXP-YYYYMMDD-序號，同日第二筆起編02、03...（跟返purchaseLog同日排序嘅做法，見§7）----
// 淨係喺新增嗰刻計一次，之後編輯呢筆記錄唔會再變號；刪除舊記錄會留低缺口，屬預期行為（唔重編號）
export function nextVoucherNo(expenseLog, dateStr) {
  const ymd = dateStr.replaceAll("-", "");
  const sameDay = (expenseLog || []).filter((r) => r.voucherNo && r.voucherNo.startsWith(`EXP-${ymd}-`));
  return `EXP-${ymd}-${String(sameDay.length + 1).padStart(2, "0")}`;
}

// ---- 支出憑證圖：Canvas疊字（§3.2定案版）----
// 用瀏覽器原生 Canvas 2D 文字渲染，唔使好似Python mockup環境咁手動載入CJK字體——
// 現代瀏覽器嘅 fillText 本身就會用系統字體做per-glyph fallback，中文一樣render得到，唔使特別處理
function roundRectPath(ctx, x, y, w, h, r) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r);
  ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r);
  ctx.arcTo(x, y, x + w, y, r);
  ctx.closePath();
}
// measureOnly=true 淨係計算需要嘅高度，唔真係畫——等call方可以先知卡片幾高，先啱啱好貼喺底部
function layoutVoucherCard(ctx, record, x, yTop, w, measureOnly) {
  const draw = !measureOnly;
  const pad = Math.round(w * 0.07);
  const titleSize = Math.max(13, Math.round(w * 0.052));
  const amountSize = Math.max(20, Math.round(w * 0.11));
  const lineSize = Math.max(11, Math.round(w * 0.042));
  let cy = yTop + pad;
  const cx = x + pad;
  const innerW = w - pad * 2;

  if (draw) {
    ctx.fillStyle = "#4ECDC4";
    ctx.beginPath(); ctx.arc(cx + titleSize * 0.35, cy + titleSize * 0.35, titleSize * 0.22, 0, Math.PI * 2); ctx.fill();
    ctx.fillStyle = "#fff"; ctx.textBaseline = "top"; ctx.font = `700 ${titleSize}px sans-serif`;
    ctx.fillText(`${BRAND_NAME} 支出憑證`, cx + titleSize * 0.9, cy);
  }
  cy += titleSize * 1.7;

  if (draw) { ctx.fillStyle = "#4ECDC4"; ctx.font = `800 ${amountSize}px sans-serif`; ctx.fillText(`HKD $${record.amount.toLocaleString()}`, cx, cy); }
  cy += amountSize * 1.35;

  if (draw) { ctx.font = `400 ${lineSize}px sans-serif`; ctx.fillStyle = "#ddd"; }
  const line = (label, value) => { if (draw) ctx.fillText(`${label}：${value}`, cx, cy); cy += lineSize * 1.6; };
  line("類別", record.category);
  line("購買日期", record.date);
  line("代付人", record.payer);
  // ⚠️歸還狀態（未歸還/已歸還）刻意唔顯示喺憑證圖度（見§3.2）——呢張圖淨係做「呢筆錢使咗」嘅單據，唔記錄內部欠款狀態

  if (record.items.length > 1) {
    cy += lineSize * 0.3;
    if (draw) { ctx.strokeStyle = "rgba(255,255,255,0.25)"; ctx.beginPath(); ctx.moveTo(cx, cy); ctx.lineTo(cx + innerW, cy); ctx.stroke(); }
    cy += lineSize * 0.7;
    if (draw) { ctx.fillStyle = "#aaa"; ctx.font = `600 ${Math.round(lineSize * 0.9)}px sans-serif`; ctx.fillText("物品明細", cx, cy); }
    cy += lineSize * 1.4;
    if (draw) { ctx.font = `400 ${Math.round(lineSize * 0.92)}px sans-serif`; ctx.fillStyle = "#ccc"; }
    record.items.forEach((it) => { if (draw) ctx.fillText(`${it.name}　$${it.amount.toLocaleString()}`, cx, cy); cy += lineSize * 1.4; });
  }

  cy += lineSize * 0.5;
  if (draw) { ctx.strokeStyle = "rgba(255,255,255,0.15)"; ctx.beginPath(); ctx.moveTo(cx, cy); ctx.lineTo(cx + innerW, cy); ctx.stroke(); }
  cy += lineSize * 0.7;
  if (draw) { ctx.fillStyle = "#888"; ctx.font = `400 ${Math.round(lineSize * 0.8)}px sans-serif`; ctx.fillText(record.voucherNo, cx, cy); }
  cy += lineSize * 1.3;

  return (cy - yTop) + pad;
}
// 生成單張支出憑證圖（JPEG blob）：有相就疊喺相右下角，冇相就用返個深色底做成張獨立嘅卡片圖。
// 呢個function純粹喺瀏覽器本機運算輸出blob，唔會寫入任何App狀態，call方負責觸發download，全程唔會上傳/存落Supabase（見§3.1）
export function buildExpenseVoucherBlob(record, receiptFile) {
  const render = (receiptImg) => new Promise((resolve, reject) => {
    const canvas = document.createElement("canvas");
    const hasPhoto = !!receiptImg;
    let cw, ch;
    if (hasPhoto) {
      const MAXLONG = 1600;
      const { naturalWidth: iw, naturalHeight: ih } = receiptImg;
      const scale = Math.min(1, MAXLONG / Math.max(iw, ih));
      cw = Math.round(iw * scale); ch = Math.round(ih * scale);
    } else { cw = 900; ch = 1150; }
    canvas.width = cw; canvas.height = ch;
    const ctx = canvas.getContext("2d");
    if (hasPhoto) ctx.drawImage(receiptImg, 0, 0, cw, ch);
    else { ctx.fillStyle = "#0f0f0f"; ctx.fillRect(0, 0, cw, ch); }

    const margin = Math.round(cw * (hasPhoto ? 0.035 : 0.06));
    const cardW = hasPhoto ? Math.max(220, Math.min(cw * 0.5, cw - margin * 2)) : cw - margin * 2;
    const cardH = layoutVoucherCard(ctx, record, 0, 0, cardW, true);
    const cardX = cw - margin - cardW;
    const cardY = ch - margin - cardH;
    ctx.fillStyle = "rgba(15,25,26,0.86)";
    roundRectPath(ctx, cardX, cardY, cardW, cardH, Math.round(cardW * 0.03));
    ctx.fill();
    layoutVoucherCard(ctx, record, cardX, cardY, cardW, false);

    canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error("toBlob失敗")), "image/jpeg", 0.9);
  });

  if (!receiptFile) return render(null);
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (ev) => {
      const img = new Image();
      img.onload = () => render(img).then(resolve, reject);
      img.onerror = () => render(null).then(resolve, reject); // 相讀取失敗就後備做空白憑證，唔好整個功能卡死
      img.src = ev.target.result;
    };
    reader.onerror = () => render(null).then(resolve, reject);
    reader.readAsDataURL(receiptFile);
  });
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

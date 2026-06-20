import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { cloudEnabled, cloudLoad, cloudSave, cloudSubscribe } from "./supabase.js";

// ---- 本機儲存（localStorage）：部署到 Vercel 後，資料會喺呢部裝置保存，重新整頁唔會消失 ----
const LS_KEY = "gymily_data_v1";
const loadStore = () => {
  try { return JSON.parse(localStorage.getItem(LS_KEY)) || {}; } catch (e) { return {}; }
};
const initStore = loadStore();
const persisted = (key, fallback) => (initStore[key] !== undefined ? initStore[key] : fallback);

// 穩定序列化：將物件 key 依字母排序，令同步比對唔受 Supabase jsonb 重排 key 影響
function stableStringify(obj) {
  if (obj === null || typeof obj !== "object") return JSON.stringify(obj);
  if (Array.isArray(obj)) return "[" + obj.map(stableStringify).join(",") + "]";
  const keys = Object.keys(obj).sort();
  return "{" + keys.map((k) => JSON.stringify(k) + ":" + stableStringify(obj[k])).join(",") + "}";
}

const LOGO = "data:image/png;base64,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";

const DEFAULT_COACHES = [];


const COLORS = ["#FF6B6B", "#4ECDC4", "#FFE66D", "#A8E6CF", "#C9B1FF", "#FFB347", "#FF8FA3", "#6BCB77"];

const MAX_CONCURRENT = 2; // 場地同時最多 2 名教練

// 一對二定價：1小時 $150，每加0.5小時 +$50
const DUO_BASE = 150;
const DUO_HALF_HOUR_ADD = 50;
function duoPrice(hours) { return DUO_BASE + Math.round((hours - 1) / 0.5) * DUO_HALF_HOUR_ADD; }

const CHARTER_PRICE = 300; // 包場 $300 一節（時長由管理員自選），佔全場

// 其他租場類型：包場/小組佔全場（2位），試堂只佔1位（可同教練並存）
const isWholeVenue = (e) => e.type === "charter" && e.charterType !== "trial";
const rentalShort = (ct) => ct === "group" ? "小組" : ct === "trial" ? "試堂" : "包場";
const rentalFull = (ct) => ct === "group" ? "小組訓練" : ct === "trial" ? "試堂" : "私人包場";
const ASSIST_CANCEL_LIMIT = 1; // 每位教練每月「24小時內由管理員代取消」額度
const LOW_CREDIT_THRESHOLD = 2; // 剩餘堂數 ≤ 此數視為「快用完」

// 休息日設定（空 = 全週開放）。如需設休息日，例如週五：改成 [5]。getDay(): 日0 一1 二2 三3 四4 五5 六6
const CLOSED_DAYS = [];
const isClosedDay = (date) => CLOSED_DAYS.includes(new Date(`${date}T00:00:00`).getDay());

// 15-min grid 07:00–22:00
const TIME_SLOTS = [];
for (let h = 7; h < 22; h++) for (let m = 0; m < 60; m += 15) TIME_SLOTS.push(`${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`);

function getDaysOfWeek(offset = 0) {
  const days = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  today.setDate(today.getDate() + offset);
  for (let i = 0; i < 7; i++) { const d = new Date(today); d.setDate(today.getDate() + i); days.push(d); }
  return days;
}
const pad2 = (n) => String(n).padStart(2, "0");
const formatDate = (date) => `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
const isTodayDate = (date) => formatDate(date) === formatDate(new Date());
const formatDay = (date) => `周${["日", "一", "二", "三", "四", "五", "六"][date.getDay()]}`;
const monthKey = (dateStr) => dateStr.slice(0, 7);
const hoursUntil = (date, time) => (new Date(`${date}T${time}:00`) - new Date()) / (1000 * 60 * 60);
function addMinutes(time, mins) {
  const [h, m] = time.split(":").map(Number);
  const total = h * 60 + m + mins;
  return `${String(Math.floor(total / 60)).padStart(2, "0")}:${String(total % 60).padStart(2, "0")}`;
}
// list of 15-min slot keys for a booking of given hours starting at time
function slotsFor(time, hours) {
  const n = Math.round(hours * 4);
  return Array.from({ length: n }, (_, i) => addMinutes(time, i * 15));
}
// 將 "HH:MM" 轉做由 07:00 起計嘅第幾個 15 分鐘格（方便計算跨度中點）
function slotIndex(time) {
  const [h, m] = time.split(":").map(Number);
  return (h - 7) * 4 + m / 15;
}
// 將一個 booking 嘅顯示內容拆做幾行：教練名/類型、（學生名）、開始時間 —— 每行會分配落唔同嘅實際格仔
function buildEntryLines(v, isTrial, coachObj, isOwner) {
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

export default function App() {
  const [coaches, setCoaches] = useState(() => persisted("coaches", DEFAULT_COACHES));
  const [currentUser, setCurrentUser] = useState(null);
  const [adminPassword, setAdminPassword] = useState(() => persisted("adminPassword", "admin123"));
  const [subAdmins, setSubAdmins] = useState(() => persisted("subAdmins", [
    { id: 1, username: "subadmin1", password: "1234", name: "副管理員1", permissions: { overview: true, schedule: true, coaches: true, ledger: true, records: true, settings: true } },
    { id: 2, username: "subadmin2", password: "1234", name: "副管理員2", permissions: { overview: true, schedule: true, coaches: true, ledger: true, records: true, settings: true } },
  ]));
  const [view, setView] = useState("login");
  // bookings: key date_time(15min) -> { coachId, start, hours, type }  (type: 'solo' | 'duo')
  const [bookings, setBookings] = useState(() => persisted("bookings", {}));
  const [purchaseLog, setPurchaseLog] = useState(() => persisted("purchaseLog", []));
  const [ledgerFilter, setLedgerFilter] = useState("all"); // coachId or "all"
  const [ledgerMonth, setLedgerMonth] = useState("all"); // "all" or "YYYY-MM"
  const [editDateRec, setEditDateRec] = useState(null); // {id, date}
  const [weekOffset, setWeekOffset] = useState(0);
  const [bookModal, setBookModal] = useState(null);   // { date, time }
  const [charterModal, setCharterModal] = useState(null); // admin charter { date, time, hours }
  const [charterLog, setCharterLog] = useState(() => persisted("charterLog", [])); // {date, bookDate, start, hours, amount}
  const [assistCancelLog, setAssistCancelLog] = useState(() => persisted("assistCancelLog", [])); // {coachId, month, date, start}
  const [cancelLog, setCancelLog] = useState(() => persisted("cancelLog", [])); // {date, start, hours, type, charterType, coachId, coachName, price, cancelledBy, cancelledAt}
  const [syncState, setSyncState] = useState(cloudEnabled ? "connecting" : "local"); // connecting | synced | local | error

  // 同步用：記住最後一次「已儲存／已收到」嘅內容，避免回音造成無限迴圈
  const lastSyncedRef = useRef(null);
  const readyRef = useRef(!cloudEnabled); // 雲端模式要等首次載入完成先準許寫入
  const saveTimer = useRef(null);

  const applyBundle = (d) => {
    if (!d) return;
    if (d.coaches !== undefined) setCoaches(d.coaches);
    if (d.adminPassword !== undefined) setAdminPassword(d.adminPassword);
    if (d.subAdmins !== undefined) setSubAdmins(d.subAdmins);
    if (d.bookings !== undefined) setBookings(d.bookings);
    if (d.purchaseLog !== undefined) setPurchaseLog(d.purchaseLog);
    if (d.charterLog !== undefined) setCharterLog(d.charterLog);
    if (d.assistCancelLog !== undefined) setAssistCancelLog(d.assistCancelLog);
    if (d.cancelLog !== undefined) setCancelLog(d.cancelLog);
  };

  // 首次載入：雲端模式由雲端讀取（若雲端空白則上載目前本機資料），並訂閱即時變更
  useEffect(() => {
    if (!cloudEnabled) return;
    let unsub = () => {};
    (async () => {
      const remote = await cloudLoad();
      if (remote && Object.keys(remote).length) {
        lastSyncedRef.current = stableStringify(remote);
        applyBundle(remote);
      } else {
        // 雲端未有資料：將目前（本機／預設）資料推上去做初始
        const seed = { coaches, adminPassword, subAdmins, bookings, purchaseLog, charterLog, assistCancelLog, cancelLog };
        lastSyncedRef.current = stableStringify(seed);
        await cloudSave(seed);
      }
      readyRef.current = true;
      setSyncState("synced");
      unsub = cloudSubscribe((d) => {
        const s = stableStringify(d);
        if (s === lastSyncedRef.current) return; // 自己嘅更新，略過
        lastSyncedRef.current = s;
        applyBundle(d);
      });
    })();
    return () => unsub();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // 任何資料變更時儲存（雲端 or 本機）
  useEffect(() => {
    const bundle = { coaches, adminPassword, subAdmins, bookings, purchaseLog, charterLog, assistCancelLog, cancelLog };
    // 本機永遠都存一份（離線後備）
    try { localStorage.setItem(LS_KEY, JSON.stringify(bundle)); } catch (e) { /* ignore */ }

    if (!cloudEnabled) return;
    if (!readyRef.current) return; // 首次雲端載入未完成，唔好覆蓋
    const s = stableStringify(bundle);
    if (s === lastSyncedRef.current) return; // 同雲端一樣，唔使寫

    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      lastSyncedRef.current = s;
      setSyncState("connecting");
      const ok = await cloudSave(bundle);
      setSyncState(ok ? "synced" : "error");
    }, 500);
  }, [coaches, adminPassword, subAdmins, bookings, purchaseLog, charterLog, assistCancelLog, cancelLog]);

  const [cancelModal, setCancelModal] = useState(null);
  const [signModal, setSignModal] = useState(null); // {date,start,coachId,type,studentName}
  const [adminCancelModal, setAdminCancelModal] = useState(null); // {date,start,coachId,type}
  const [delLedgerModal, setDelLedgerModal] = useState(null); // ledger record to delete
  const [toast, setToast] = useState(null);
  const [loginForm, setLoginForm] = useState({ id: "", password: "" });
  const [pwForm, setPwForm] = useState({ old: "", new1: "", new2: "" });
  const [editCoach, setEditCoach] = useState(null);
  const [addCreditModal, setAddCreditModal] = useState(null);
  const [adminTab, setAdminTab] = useState("overview");
  const [recordsView, setRecordsView] = useState("bookings"); // bookings | cancelled
  const [recCoach, setRecCoach] = useState("all");
  const [recType, setRecType] = useState("all"); // all|solo|duo|private|group|trial
  const [recRange, setRecRange] = useState("upcoming"); // upcoming|past|month|all
  const [recMonth, setRecMonth] = useState(() => monthKey(formatDate(new Date())));
  const [recExpanded, setRecExpanded] = useState(null);
  const [coachSort, setCoachSort] = useState("remain"); // remain|paid|name
  const [expandedCoachId, setExpandedCoachId] = useState(null);
  const [viewMonth, setViewMonth] = useState(() => monthKey(formatDate(new Date())));
  const [monthsExpanded, setMonthsExpanded] = useState(false);
  const [newStudentName, setNewStudentName] = useState("");
  const [addStudentCreditModal, setAddStudentCreditModal] = useState(null); // {name, qty}
  const [studentLogOpen, setStudentLogOpen] = useState(null);
  const [resetModal, setResetModal] = useState(false);
  const [delCoachModal, setDelCoachModal] = useState(null); // coach pending deletion
  const [showPasswords, setShowPasswords] = useState(false);

  const days = getDaysOfWeek(weekOffset * 7);
  const showToast = (msg, type = "success") => { setToast({ msg, type }); setTimeout(() => setToast(null), 3000); };
  const getCoach = (id) => coaches.find((c) => c.id === id);

  const handleLogin = () => {
    const uid = loginForm.id.trim().toLowerCase();
    if (uid === "admin") {
      if (loginForm.password === adminPassword) { setCurrentUser({ id: 0, name: "管理員", role: "admin" }); setView("admin"); }
      else showToast("管理員密碼錯誤", "error");
      return;
    }
    const sub = subAdmins.find((s) => (s.username || "").toLowerCase() === uid && s.password === loginForm.password);
    if (sub) {
      setCurrentUser({ ...sub, role: "subadmin" });
      setView("admin");
      const firstAllowed = ["overview", "schedule", "coaches", "ledger", "records", "settings"].find((k) => sub.permissions?.[k]);
      setAdminTab(firstAllowed || "settings");
      return;
    }
    const coach = coaches.find((c) => (c.username || "").toLowerCase() === uid && c.password === loginForm.password);
    if (coach) { setCurrentUser({ ...coach, role: "coach" }); setView("calendar"); }
    else showToast("帳號或密碼錯誤", "error");
  };
  const logout = () => { setCurrentUser(null); setView("login"); setLoginForm({ id: "", password: "" }); };

  const isCoach = currentUser?.role === "coach";
  const liveUser = isCoach ? getCoach(currentUser.id) : currentUser;
  const remaining = isCoach && liveUser ? liveUser.credits - liveUser.used : 0;
  const soldOut = isCoach && remaining <= 0;

  // bookings[key] is an ARRAY of entries; each entry occupies "seats" (包場/小組=2, 其他=1)
  const cellArr = (date, slot) => bookings[`${date}_${slot}`] || [];
  const seats = (entry) => (isWholeVenue(entry) ? MAX_CONCURRENT : 1);
  const occupancy = (date, slot) => cellArr(date, slot).reduce((n, e) => n + seats(e), 0);

  // can we place a booking of `hours` at date/time? need = 需要幾多個位
  const canPlace = (date, time, hours, need = 1) => {
    const slots = slotsFor(time, hours);
    for (const s of slots) {
      const [hh] = s.split(":").map(Number);
      if (hh >= 22) return "超出營業時間";
      if (occupancy(date, s) + need > MAX_CONCURRENT)
        return need >= MAX_CONCURRENT ? "呢個時段唔夠空（包場／小組需全場）" : "呢個時段已滿（最多2名）";
    }
    return null;
  };

  const openBook = (date, time) => {
    if (isClosedDay(date)) return showToast("星期四、五休息，不開放預約", "error");
    if (soldOut) return showToast("你已用晒購買堂數，請聯絡管理員增購", "error");
    const allowSolo = liveUser.allowSolo !== false;
    const allowDuo = liveUser.allowDuo !== false;
    if (!allowSolo && !allowDuo) return showToast("你冇任何可用嘅預約類型，請聯絡管理員設定", "error");
    setBookModal({ date, time, sessionType: allowSolo ? "solo" : "duo", hours: 1, students: [], studentOther: "" });
  };

  const confirmBook = () => {
    const { date, time, sessionType, hours } = bookModal;
    if (sessionType === "solo" && liveUser.allowSolo === false) { showToast("你冇一對一預約權限", "error"); return; }
    if (sessionType === "duo" && liveUser.allowDuo === false) { showToast("你冇一對二預約權限", "error"); return; }
    const creditCost = hours; // 1hr = 1堂, 1.5hr = 1.5堂
    if (creditCost > remaining) { showToast("剩餘堂數不足", "error"); return; }
    const err = canPlace(date, time, hours);
    if (err) { showToast(err, "error"); return; }
    const price = sessionType === "duo" ? duoPrice(hours) : liveUser.rate * hours;
    const rentalCost = liveUser.rate * hours; // 租場費用：用「落單嗰刻」嘅租金snapshot，日後改租金唔會影響舊紀錄
    const slots = slotsFor(time, hours);
    const selected = Array.isArray(bookModal.students) ? bookModal.students : [];
    const extra = (bookModal.studentOther || "").trim();
    const studentList = [...selected, ...(extra ? [extra] : [])].filter(Boolean).slice(0, 4);
    // 將每位學生「當時」嘅收費 snapshot 落呢張記錄（roster 有嘅用 roster 收費；自填嘅新名暫定 $0，教練可以之後再幫佈設收費）
    const studentCharges = {};
    studentList.forEach((n) => { const s = myRoster.find((x) => x.name === n); studentCharges[n] = s ? (s.rate || 0) : 0; });
    const entry = { coachId: currentUser.id, start: time, hours, type: sessionType, price, rentalCost, students: studentList, studentCharges, createdAt: new Date().toISOString().slice(0, 16).replace("T", " ") };
    setBookings((prev) => {
      const u = { ...prev };
      slots.forEach((s) => { u[`${date}_${s}`] = [...(u[`${date}_${s}`] || []), entry]; });
      return u;
    });
    setCoaches((prev) => prev.map((c) => c.id === currentUser.id ? { ...c, used: c.used + creditCost } : c));
    showToast("預約成功！");
    setBookModal(null);
  };

  // ADMIN: place a rental (包場/小組=全場2位, 試堂=1位), price editable
  const confirmCharter = () => {
    const { date, time, hours, charterType, price, coachName } = charterModal;
    if (isClosedDay(date)) { showToast("休息日", "error"); return; }
    const need = charterType === "trial" ? 1 : MAX_CONCURRENT;
    const err = canPlace(date, time, hours, need);
    if (err) { showToast(err, "error"); return; }
    const amt = charterType === "trial" ? 0 : (parseInt(price) || 0);
    const slots = slotsFor(time, hours);
    const entry = { coachId: 0, start: time, hours, type: "charter", charterType, price: amt, coachName: coachName || "", createdAt: new Date().toISOString().slice(0, 16).replace("T", " ") };
    setBookings((prev) => {
      const u = { ...prev };
      slots.forEach((s) => { u[`${date}_${s}`] = [...(u[`${date}_${s}`] || []), entry]; });
      return u;
    });
    setCharterLog((prev) => [{ date: new Date().toISOString().slice(0, 16).replace("T", " "), bookDate: date, start: time, hours, charterType, amount: amt, coachName: coachName || "" }, ...prev]);
    showToast(`已落${rentalFull(charterType)}（$${amt}）`);
    setCharterModal(null);
  };

  const openCancel = (date, start, coachId, type) => {
    const hrs = hoursUntil(date, start);
    const win = getCoach(coachId)?.cancelWindowHours ?? 24;
    if (currentUser.role === "coach" && hrs < win) return showToast(`${win}小時內取消需要管理員協助`, "error");
    setCancelModal({ date, start, coachId, type });
  };

  const doCancel = (date, start, coachId, type, byAdmin = false) => {
    // locate one matching entry to read its hours
    const startArr = cellArr(date, start);
    const meta = startArr.find((e) => e.coachId === coachId && e.start === start && (type ? e.type === type : true));
    if (!meta) { setCancelModal(null); return; }
    const slots = slotsFor(start, meta.hours);
    setBookings((prev) => {
      const u = { ...prev };
      slots.forEach((s) => {
        const arr = (u[`${date}_${s}`] || []).filter((e) => !(e.coachId === coachId && e.start === start && e.type === meta.type));
        if (arr.length) u[`${date}_${s}`] = arr; else delete u[`${date}_${s}`];
      });
      return u;
    });
    // 留底：取消記錄（先記低先删，等日後可以查到呢個時段點解空咗）
    setCancelLog((prev) => [{
      date, start, hours: meta.hours, type: meta.type, charterType: meta.charterType || null,
      coachId: meta.type === "charter" ? null : coachId,
      coachName: meta.type === "charter" ? (meta.coachName || "") : (getCoach(coachId)?.name || ""),
      price: meta.price || 0,
      cancelledBy: byAdmin ? "admin" : "coach",
      cancelledAt: new Date().toISOString().slice(0, 16).replace("T", " "),
    }, ...prev]);
    if (meta.type !== "charter") {
      setCoaches((prev) => prev.map((c) => c.id === coachId ? { ...c, used: Math.max(0, c.used - meta.hours) } : c));
      // 由管理員協助、而且係該教練設定嘅通知時數內取消，計入本月額度
      const win = getCoach(coachId)?.cancelWindowHours ?? 24;
      if (byAdmin && hoursUntil(date, start) < win) {
        setAssistCancelLog((prev) => [{ coachId, month: monthKey(formatDate(new Date())), date, start }, ...prev]);
      }
    } else {
      setCharterLog((prev) => prev.filter((r) => !(r.bookDate === date && r.start === start)));
    }
    showToast(byAdmin ? "已協助取消" : "已取消預約");
    setCancelModal(null);
  };

  // 學生簽到：將一個學生嘅簽名（base64 圖）寫入呢個 booking 嘅所有 15 分鐘格副本
  const signIn = (date, start, coachId, type, studentName, dataUrl) => {
    const startArr = cellArr(date, start);
    const meta = startArr.find((e) => e.coachId === coachId && e.start === start && e.type === type);
    if (!meta) return;
    const slots = slotsFor(start, meta.hours);
    setBookings((prev) => {
      const u = { ...prev };
      slots.forEach((s) => {
        const arr = (u[`${date}_${s}`] || []).map((e) => {
          if (e.coachId === coachId && e.start === start && e.type === type) {
            return { ...e, signatures: { ...(e.signatures || {}), [studentName]: { dataUrl, signedAt: new Date().toISOString().slice(0, 16).replace("T", " ") } } };
          }
          return e;
        });
        u[`${date}_${s}`] = arr;
      });
      return u;
    });
    // 簽到先確實上堂 → 先扣呢個學生嘅堂數（取消咗嘅堂冇人簽，自然唔會扣錯）
    // 注意：唔可以靠 setCoaches 嘅 updater 嚎計提醒文字，因為 updater 唔保證即刻執行（React 會 batch），
    // 跟住嗰行 showToast 好多時讀到嘅仲係舊值。所以要喺呼叫 setCoaches 之前，用現有資料直接計好。
    const roster = getStudentRoster(coachId);
    const sIdx = roster.findIndex((s) => s.name === studentName);
    let remainText = "";
    if (sIdx !== -1) {
      const newUsed = (roster[sIdx].used || 0) + meta.hours;
      const remain = (roster[sIdx].credits || 0) - newUsed;
      if (remain <= LOW_CREDIT_THRESHOLD) remainText = `　⚠️ 剩返 ${Math.max(0, remain)} 堂`;
      setCoaches((prev) => prev.map((c) => {
        if (c.id !== coachId) return c;
        const r = (c.studentRoster || []).map(normStudent);
        const i = r.findIndex((s) => s.name === studentName);
        if (i === -1) return c;
        const newRoster = [...r]; newRoster[i] = { ...r[i], used: newUsed };
        return { ...c, studentRoster: newRoster };
      }));
    }
    showToast(`${studentName} 已簽到${remainText}`);
  };

  // 教練本月已用 / 剩餘代取消額度
  const assistUsedThisMonth = (coachId) => {
    const m = monthKey(formatDate(new Date()));
    return assistCancelLog.filter((r) => r.coachId === coachId && r.month === m).length;
  };

  // 學生名單由舊版「淨係名」升級做完整 record；呢個 helper 兩種格式都食得（向後兼容舊資料）
  const normStudent = (s) => (typeof s === "string" ? { name: s, rate: 0, credits: 0, used: 0 } : { rate: 0, credits: 0, used: 0, ...s });
  const myRoster = (liveUser?.studentRoster || []).map(normStudent);
  const getStudentRoster = (coachId) => (getCoach(coachId)?.studentRoster || []).map(normStudent);

  // 教練新增學生入自己嘅名單（去重、淨係自己改到自己嗰份）
  const addStudentToRoster = () => {
    const name = newStudentName.trim();
    if (!name) return;
    if (!isCoach) return;
    if (myRoster.some((s) => s.name === name)) { showToast("呢個名已經喺名單度", "error"); return; }
    setCoaches((prev) => prev.map((c) => c.id === currentUser.id ? { ...c, studentRoster: [...myRoster, { name, rate: 0, credits: 0, used: 0 }] } : c));
    setNewStudentName("");
  };

  const updateStudentField = (name, field, value) => {
    setCoaches((prev) => prev.map((c) => {
      if (c.id !== currentUser.id) return c;
      const roster = (c.studentRoster || []).map(normStudent);
      return { ...c, studentRoster: roster.map((s) => s.name === name ? { ...s, [field]: value } : s) };
    }));
  };

  const removeStudent = (name) => {
    setCoaches((prev) => prev.map((c) => c.id === currentUser.id ? { ...c, studentRoster: myRoster.filter((s) => s.name !== name) } : c));
  };

  // 幫學生開堂數（教練自己用，類似 admin 幫教練「+ 堂」嗰個概念）
  const addStudentCredits = (name, qty) => {
    const s = myRoster.find((x) => x.name === name);
    if (!s) return;
    updateStudentField(name, "credits", (s.credits || 0) + qty);
    showToast(`已為 ${name} 增加 ${qty} 堂`);
  };

  // 直接修改「剩餘堂數」：保留「已用」（簽到歷史唔變），改返「已開」嚎夾返新嘅剩餘數
  const setStudentRemain = (name, newRemain) => {
    const s = myRoster.find((x) => x.name === name);
    if (!s) return;
    const used = s.used || 0;
    const newCredits = Math.max(0, used + newRemain);
    updateStudentField(name, "credits", newCredits);
  };

  const changePassword = () => {
    if (pwForm.new1 !== pwForm.new2) return showToast("兩次新密碼唔一致", "error");
    if (pwForm.new1.length < 4) return showToast("密碼至少4位", "error");
    if (currentUser.role === "admin") {
      if (pwForm.old !== adminPassword) return showToast("舊密碼錯誤", "error");
      setAdminPassword(pwForm.new1);
    } else if (currentUser.role === "subadmin") {
      const me = subAdmins.find((s) => s.id === currentUser.id);
      if (!me || pwForm.old !== me.password) return showToast("舊密碼錯誤", "error");
      setSubAdmins((prev) => prev.map((s) => s.id === currentUser.id ? { ...s, password: pwForm.new1 } : s));
    } else {
      if (pwForm.old !== getCoach(currentUser.id).password) return showToast("舊密碼錯誤", "error");
      setCoaches((prev) => prev.map((c) => c.id === currentUser.id ? { ...c, password: pwForm.new1 } : c));
    }
    setPwForm({ old: "", new1: "", new2: "" });
    showToast("密碼已更新");
  };

  const addCredits = (coachId, qty, expiryDate) => {
    const coach = getCoach(coachId);
    const amount = qty * coach.rate;
    setCoaches((prev) => prev.map((c) => c.id === coachId ? { ...c, credits: c.credits + qty } : c));
    setPurchaseLog((prev) => [{ id: "p" + Date.now() + "-" + Math.random().toString(36).slice(2), date: new Date().toISOString().slice(0, 10), coachId, coachName: coach.name, qty, amount, rate: coach.rate, expiryDate: expiryDate || null }, ...prev]);
    showToast(`已為 ${coach.name} 增加 ${qty} 堂（$${amount}）${expiryDate ? `，失效日：${expiryDate}` : ""}`);
  };

  // FIFO：將某教練嘅 used 堂數，依購買時間順序分配到每筆購買記錄，計出每筆嘅「已用／剩餘」
  const purchaseFifoStatus = (coachId) => {
    const batches = purchaseLog.filter((r) => r.coachId === coachId).slice().sort((a, b) => a.date.localeCompare(b.date));
    const coach = getCoach(coachId);
    let remainingToAllocate = coach ? coach.used : 0;
    return batches.map((r) => {
      const consumed = Math.min(r.qty, remainingToAllocate);
      remainingToAllocate -= consumed;
      return { ...r, consumed, remaining: r.qty - consumed };
    });
  };

  // 邊位教練有「未用完、已過期或快過期」嘅堂數（用嚎提醒管理員／教練）
  const EXPIRY_WARN_DAYS = 14;
  const expiringBatchesOf = (coachId) => {
    const today = formatDate(new Date());
    return purchaseFifoStatus(coachId).filter((b) => b.remaining > 0 && b.expiryDate &&
      (new Date(`${b.expiryDate}T00:00:00`) - new Date(`${today}T00:00:00`)) / 86400000 <= EXPIRY_WARN_DAYS);
  };

  // sanitize sheet names (Excel: <=31 chars, no : \ / ? * [ ])
  const sheetName = (s) => (s || "").replace(/[:\\/?*[\]]/g, " ").slice(0, 28).trim() || "Sheet";

  const exportExcel = () => {
    try {
      const wb = XLSX.utils.book_new();

      // 1) 教練總覽
      const coachRows = coaches.map((c) => ({
        教練: c.name, 帳號: c.username, 已購買堂數: c.credits, 已用堂數: c.used,
        剩餘堂數: c.credits - c.used, 每堂租金: c.rate,
        總付款: purchaseLog.filter((r) => r.coachId === c.id).reduce((a, r) => a + r.amount, 0)
          + Math.max(0, c.credits - purchaseLog.filter((r) => r.coachId === c.id).reduce((a, r) => a + r.qty, 0)) * c.rate,
      }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(coachRows), "教練總覽");

      // 2) 流水帳（全部）
      const ledgerRows = purchaseLog.map((r) => ({ 日期: r.date, 教練: r.coachName, 增加堂數: r.qty, 每堂: r.rate, 金額: r.amount }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ledgerRows.length ? ledgerRows : [{ 日期: "", 教練: "", 增加堂數: "", 每堂: "", 金額: "" }]), "流水帳-全部");

      // 3) 每個教練獨立流水帳
      coaches.forEach((c) => {
        const rows = purchaseLog.filter((r) => r.coachId === c.id).map((r) => ({ 日期: r.date, 增加堂數: r.qty, 每堂: r.rate, 金額: r.amount }));
        const total = rows.reduce((a, r) => a + r.金額, 0);
        const body = rows.length ? [...rows, { 日期: "小計", 增加堂數: "", 每堂: "", 金額: total }] : [{ 日期: "（無記錄）", 增加堂數: "", 每堂: "", 金額: "" }];
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(body), sheetName("流水-" + c.name));
      });

      // 4) 上堂記錄
      const bkRows = [];
      Object.entries(bookings).forEach(([k, arr]) => {
        const date = k.split("_")[0];
        arr.forEach((v) => { if (k === `${date}_${v.start}`) bkRows.push({
          日期: date, 開始: v.start, 時長小時: v.hours,
          類型: v.type === "charter" ? rentalFull(v.charterType) : v.type === "duo" ? "一對二" : "一對一",
          教練: v.type === "charter" ? (v.coachName || "") : (getCoach(v.coachId)?.name || ""),
          收費: v.price || 0,
        }); });
      });
      bkRows.sort((a, b) => `${a.日期}${a.開始}`.localeCompare(`${b.日期}${b.開始}`));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(bkRows.length ? bkRows : [{ 日期: "", 開始: "", 時長小時: "", 類型: "", 教練: "", 收費: "" }]), "上堂記錄");

      // 5) 包場/小組
      const activeCh = charterLog.map((r) => ({ 落單時間: r.date, 預約日期: r.bookDate, 開始: r.start, 時長小時: r.hours, 類型: rentalFull(r.charterType), 負責教練: r.coachName || "", 收費: r.amount, 已取消: "否", 取消時間: "" }));
      const cancelledCh = cancelLog.filter((r) => r.type === "charter").map((r) => ({ 落單時間: "", 預約日期: r.date, 開始: r.start, 時長小時: r.hours, 類型: rentalFull(r.charterType), 負責教練: r.coachName || "", 收費: r.price || 0, 已取消: "是", 取消時間: r.cancelledAt || "" }));
      const chRows = [...activeCh, ...cancelledCh];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(chRows.length ? chRows : [{ 落單時間: "", 預約日期: "", 開始: "", 時長小時: "", 類型: "", 負責教練: "", 收費: "", 已取消: "", 取消時間: "" }]), "包場小組");

      // 取消記錄
      const cxRows = cancelLog.map((r) => ({ 原定日期: r.date, 開始: r.start, 時長小時: r.hours, 類型: r.type === "charter" ? rentalFull(r.charterType) : r.type === "duo" ? "一對二" : "一對一", 教練: r.coachName || "", 收費: r.price || 0, 取消方式: r.cancelledBy === "admin" ? "管理員代取消" : "教練自行取消", 取消時間: r.cancelledAt || "" }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cxRows.length ? cxRows : [{ 原定日期: "", 開始: "", 時長小時: "", 類型: "", 教練: "", 收費: "", 取消方式: "", 取消時間: "" }]), "取消記錄");

      const today = new Date().toISOString().slice(0, 10);
      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `Gymily_資料備份_${today}.xlsx`;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1000);
      showToast("已匯出 Excel（請查看下載）");
    } catch (e) {
      showToast("下載被阻擋，請改用「複製 CSV」", "error");
    }
  };

  // 後備：複製流水帳做 CSV，貼入 Excel / Google Sheets
  const copyLedgerCsv = async () => {
    const header = ["日期", "教練", "增加堂數", "每堂", "金額"];
    const lines = [header.join(",")];
    purchaseLog.forEach((r) => lines.push([r.date, r.coachName, r.qty, r.rate, r.amount].join(",")));
    const csv = lines.join("\n");
    try {
      await navigator.clipboard.writeText(csv);
      showToast("流水帳已複製，可貼入 Excel");
    } catch (e) {
      // 再後備：用 textarea + execCommand
      try {
        const ta = document.createElement("textarea");
        ta.value = csv; ta.style.position = "fixed"; ta.style.opacity = "0";
        document.body.appendChild(ta); ta.focus(); ta.select();
        document.execCommand("copy");
        document.body.removeChild(ta);
        showToast("流水帳已複製，可貼入 Excel");
      } catch (e2) {
        showToast("複製失敗，請改用電腦版", "error");
      }
    }
  };

  // distinct bookings of current user (one entry per start)
  const myBookings = [];
  Object.entries(bookings).forEach(([k, arr]) => {
    const date = k.split("_")[0];
    arr.forEach((v) => {
      if (v.coachId === currentUser?.id && k === `${date}_${v.start}`)
        myBookings.push({ date, start: v.start, hours: v.hours, type: v.type, price: v.price || 0, rentalCost: v.rentalCost ?? (v.price || 0), students: v.students || [], studentCharges: v.studentCharges || {}, signatures: v.signatures || {} });
    });
  });
  myBookings.sort((a, b) => `${a.date}${a.start}`.localeCompare(`${b.date}${b.start}`));

  // 教練近3個月實際收入（只計有填學生名嘅堂，用 snapshot 收費；扣除租場費用）+ 各學生上堂紀錄（近3個月）
  const myIncomeReport = (() => {
    if (!isCoach) return { months: [], studentLog: {} };
    const now = new Date();
    const cutoffs = [0, 1, 2].map((i) => monthKey(formatDate(new Date(now.getFullYear(), now.getMonth() - i, 1))));
    const months = cutoffs.map((tm) => {
      const bs = myBookings.filter((b) => monthKey(b.date) === tm && b.students.length > 0);
      const gross = bs.reduce((s, b) => s + Object.values(b.studentCharges || {}).reduce((a, v) => a + v, 0), 0);
      const rentalCost = bs.reduce((s, b) => s + (b.rentalCost || 0), 0);
      return { month: tm, gross, rentalCost, net: gross - rentalCost, count: bs.length };
    });
    const studentLog = {};
    myBookings.filter((b) => cutoffs.includes(monthKey(b.date))).forEach((b) => {
      (b.students || []).forEach((n) => {
        if (!studentLog[n]) studentLog[n] = [];
        studentLog[n].push({ date: b.date, start: b.start, hours: b.hours, type: b.type, charge: (b.studentCharges || {})[n] || 0 });
      });
    });
    Object.values(studentLog).forEach((arr) => arr.sort((a, b) => `${b.date}${b.start}`.localeCompare(`${a.date}${a.start}`)));
    return { months, studentLog };
  })();

  // ---------- LOGIN ----------
  if (view === "login") {
    return (
      <div style={S.loginBg}>
        <div style={S.loginCard}>
          <img src={LOGO} alt="Gymily Studio" style={S.loginLogoImg} />
          <p style={S.loginSub}>場地預約系統</p>
          <Field label="帳號名稱"><input style={S.input} placeholder=""
            value={loginForm.id} onChange={(e) => setLoginForm({ ...loginForm, id: e.target.value })}
            onKeyDown={(e) => e.key === "Enter" && handleLogin()} /></Field>
          <Field label="密碼"><input style={S.input} type="password" placeholder="密碼"
            value={loginForm.password} onChange={(e) => setLoginForm({ ...loginForm, password: e.target.value })}
            onKeyDown={(e) => e.key === "Enter" && handleLogin()} /></Field>
          <button style={S.loginBtn} onClick={handleLogin}>登入</button>
        </div>
        {toast && <Toast toast={toast} />}
      </div>
    );
  }

  // ---------- ADMIN ----------
  if (view === "admin") {
    const totalSold = coaches.reduce((s, c) => s + c.credits, 0);
    const totalUsed = coaches.reduce((s, c) => s + c.used, 0);

    // helper: a coach's initial credits not represented in the purchase log
    const initialCreditsOf = (c) => {
      const fromLog = purchaseLog.filter((r) => r.coachId === c.id).reduce((a, r) => a + r.qty, 0);
      return Math.max(0, c.credits - fromLog);
    };

    // 一筆過租金（買堂）收入，按入數月份
    const purchaseByMonth = {};
    purchaseLog.forEach((r) => { const m = monthKey(r.date); purchaseByMonth[m] = (purchaseByMonth[m] || 0) + r.amount; });
    coaches.forEach((c) => { const init = initialCreditsOf(c); if (init > 0) purchaseByMonth["初始"] = (purchaseByMonth["初始"] || 0) + init * c.rate; });

    // 所有 booking（去重，每節一條），附帶實際收費
    const allBookings = [];
    Object.entries(bookings).forEach(([k, arr]) => {
      const date = k.split("_")[0];
      arr.forEach((v) => {
        if (k === `${date}_${v.start}`)
          allBookings.push({ date, start: v.start, hours: v.hours, type: v.type, charterType: v.charterType, price: v.price || 0, coachName: v.coachName || "", coach: v.type === "charter" ? null : getCoach(v.coachId), coachId: v.coachId, createdAt: v.createdAt || null });
      });
    });
    allBookings.sort((a, b) => `${a.date}${a.start}`.localeCompare(`${b.date}${b.start}`));

    // 實際堂數收入，按 booking 月份
    const classByMonth = {};
    allBookings.forEach((b) => { const m = monthKey(b.date); classByMonth[m] = (classByMonth[m] || 0) + (b.price || 0); });

    // 月份清單（兩個來源合併）
    const allMonths = Array.from(new Set([...Object.keys(purchaseByMonth), ...Object.keys(classByMonth)]))
      .sort((a, b) => b.localeCompare(a));

    const totalPurchase = Object.values(purchaseByMonth).reduce((a, b) => a + b, 0);
    const totalCharter = charterLog.reduce((s, r) => s + r.amount, 0);
    const totalClassRev = Object.values(classByMonth).reduce((a, b) => a + b, 0);
    const totalRevenue = totalPurchase + totalCharter; // 實收現金：買堂 + 包場/小組

    // 指定月份總收入：該月買堂 + 該月包場/小組（試堂 $0 自動唔計）
    const thisMonth = monthKey(formatDate(new Date()));
    const monthPurchase = purchaseLog.filter((r) => monthKey(r.date) === viewMonth).reduce((a, r) => a + r.amount, 0);
    const monthCharter = charterLog.filter((r) => monthKey(r.bookDate) === viewMonth).reduce((a, r) => a + r.amount, 0);
    const monthRevenue = monthPurchase + monthCharter;

    // 各教練總付款（買堂 + 初始）
    const coachPaid = {};
    coaches.forEach((c) => { coachPaid[c.id] = purchaseLog.filter((r) => r.coachId === c.id).reduce((a, r) => a + r.amount, 0) + initialCreditsOf(c) * c.rate; });

    const isSubAdmin = currentUser.role === "subadmin";
    const visibleTabs = [["overview", "📊 總覽"], ["schedule", "📅 課表"], ["coaches", "👥 教練"], ["ledger", "💰 流水帳"], ["records", "📋 記錄"], ["settings", "⚙️ 設定"]]
      .filter(([k]) => !isSubAdmin || currentUser.permissions?.[k]);
    return (
      <div style={S.appBg}>
        <Header title={isSubAdmin ? `副管理員 · ${currentUser.name}` : "管理員"} onLogout={logout} syncState={syncState} />
        <div style={S.tabRow}>
          {visibleTabs.map(([k, label]) => (
            <button key={k} style={adminTab === k ? S.tabActive : S.tab} onClick={() => setAdminTab(k)}>{label}</button>
          ))}
        </div>

        {isSubAdmin && !visibleTabs.some(([k]) => k === adminTab) && (
          <div style={S.container}><p style={S.emptyText}>你呢個帳戶暫時冇任何已啟用嘅功能，請聯絡管理員開通。</p></div>
        )}

        {adminTab === "overview" && (
          <div style={S.container}>
            <div style={{ ...S.flexBetween, marginBottom: 10 }}>
              <span style={{ fontSize: 12, color: "#888" }}>查看月份</span>
              <input style={S.select} type="month" value={viewMonth} onChange={(e) => setViewMonth(e.target.value)} />
            </div>
            <div style={S.kpiRow}>
              <div style={S.kpiCard}><div style={S.kpiLabel}>{viewMonth === thisMonth ? "本月總收入" : `${viewMonth} 收入`}</div><div style={S.kpiBig}>${monthRevenue.toLocaleString()}</div></div>
              <div style={S.kpiCard}><div style={S.kpiLabel}>已上堂數</div><div style={S.kpiBig}>{totalUsed}</div></div>
              <div style={S.kpiCard}><div style={S.kpiLabel}>已售堂數</div><div style={S.kpiBig}>{totalSold}</div></div>
            </div>
            <p style={S.assistHint}>本月＝{thisMonth}　｜　累計總收入 ${totalRevenue.toLocaleString()}</p>

            <div style={{ ...S.flexBetween, marginBottom: 0 }}>
              <h2 style={S.sectionTitle}>每月收入</h2>
              {allMonths.length > 6 && (
                <button style={S.linkBtn} onClick={() => setMonthsExpanded((v) => !v)}>{monthsExpanded ? "收埋" : `顯示全部（${allMonths.length}）`}</button>
              )}
            </div>
            <div style={S.bookingList}>
              {allMonths.length === 0 ? <p style={S.emptyText}>暫無收入</p> : (monthsExpanded ? allMonths : allMonths.slice(0, 6)).map((m) => (
                <div key={m} style={S.monthCard}>
                  <div style={S.monthHead}>{m === "初始" ? "初始已售堂數" : m}</div>
                  <div style={S.monthRow}><span style={S.monthLabel}>一筆過租金（買堂）</span><span style={S.revenueNum}>${(purchaseByMonth[m] || 0).toLocaleString()}</span></div>
                  <div style={S.monthRow}><span style={S.monthLabel}>實際堂數收入</span><span style={S.classNum}>${(classByMonth[m] || 0).toLocaleString()}</span></div>
                </div>
              ))}
            </div>
            {!monthsExpanded && allMonths.length > 6 && <p style={S.assistHint}>顯示最近 6 個月，撳上面「顯示全部」睇齊歷史。</p>}
            <p style={S.assistHint}>「一筆過租金」= 教練買堂時實收現金；「實際堂數收入」= 當月實際 book 咗嘅堂（一對一／一對二／包場）價值。</p>

            <div style={{ ...S.flexBetween, marginTop: 24 }}>
              <h2 style={{ ...S.sectionTitle, marginBottom: 0 }}>各教練統計</h2>
              <select style={S.select} value={coachSort} onChange={(e) => setCoachSort(e.target.value)}>
                <option value="remain">剩餘堂數（少→多）</option>
                <option value="paid">總付款（多→少）</option>
                <option value="name">名稱</option>
              </select>
            </div>
            {(() => {
              const lowList = coaches.filter((c) => (c.credits - c.used) <= LOW_CREDIT_THRESHOLD);
              const expiringCoaches = coaches.filter((c) => expiringBatchesOf(c.id).length > 0);
              const sorted = [...coaches].sort((a, b) => {
                if (coachSort === "paid") return (coachPaid[b.id] || 0) - (coachPaid[a.id] || 0);
                if (coachSort === "name") return a.name.localeCompare(b.name);
                return (a.credits - a.used) - (b.credits - b.used); // remain asc
              });
              return (
              <>
                {lowList.length > 0 && (
                  <div style={S.lowWarnBox}>
                    ⚠️ 堂數快用完（剩 ≤ {LOW_CREDIT_THRESHOLD}）：{lowList.map((c) => `${c.name}（剩${c.credits - c.used}）`).join("、")}　— 可提早提醒增購
                  </div>
                )}
                {expiringCoaches.length > 0 && (
                  <div style={{ ...S.lowWarnBox, background: "#332a0f", color: "#FFB347" }}>
                    ⏰ 堂數即將／已經過期：{expiringCoaches.map((c) => {
                      const batches = expiringBatchesOf(c.id);
                      return `${c.name}（${batches.map((b) => `${b.remaining}堂@${b.expiryDate}`).join("、")}）`;
                    }).join("　")}
                  </div>
                )}
                <div style={{ ...S.bookingList, marginTop: 12 }}>
                  {sorted.map((c) => {
                    const remain = c.credits - c.used;
                    const low = remain <= LOW_CREDIT_THRESHOLD;
                    const expanded = expandedCoachId === c.id;
                    const fifo = expanded ? purchaseFifoStatus(c.id) : [];
                    return (
                      <div key={c.id}>
                        <div style={{ ...S.coachStatRow, cursor: "pointer" }} onClick={() => setExpandedCoachId(expanded ? null : c.id)}>
                          <div style={{ ...S.avatar, background: c.color }}>{c.initials}</div>
                          <div style={{ flex: 1 }}>
                            <div style={S.bookingCoach}>{c.name}{low && <span style={S.lowPill}>低</span>}</div>
                            <div style={S.bookingTime}>每堂 ${c.rate}　已用 {c.used}/{c.credits} 堂</div>
                          </div>
                          <div style={{ textAlign: "right" }}>
                            <div style={S.revenueNum}>${(coachPaid[c.id] || 0).toLocaleString()}</div>
                            <div style={S.bookingTime}>總付款　剩 <span style={{ color: low ? "#FF8FA3" : "#aaa", fontWeight: low ? 700 : 400 }}>{remain}</span> 堂</div>
                          </div>
                        </div>
                        {expanded && (
                          <div style={S.purchaseBreakdown}>
                            {fifo.length === 0 ? <p style={S.emptyText}>暫無購買記錄</p> : fifo.map((b) => {
                              const isExpired = b.expiryDate && b.remaining > 0 && b.expiryDate < formatDate(new Date());
                              const isExpiring = !isExpired && b.expiryDate && b.remaining > 0 && expiringBatchesOf(c.id).some((x) => x.id === b.id);
                              return (
                                <div key={b.id} style={S.purchaseRow}>
                                  <div>
                                    <div style={S.bookingTime}>{b.date}　+{b.qty} 堂　$@{b.rate}</div>
                                    {b.expiryDate && <div style={{ fontSize: 11, color: isExpired ? "#FF6B6B" : isExpiring ? "#FFB347" : "#666" }}>失效日：{b.expiryDate}{isExpired ? "（已過期）" : isExpiring ? "（快到期）" : ""}</div>}
                                  </div>
                                  <div style={{ textAlign: "right", fontSize: 12 }}>
                                    <div style={{ color: "#6BCB77" }}>已用 {b.consumed}</div>
                                    <div style={{ color: b.remaining > 0 ? "#4ECDC4" : "#555" }}>剩 {b.remaining}</div>
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </>
              );
            })()}
          </div>
        )}

        {adminTab === "schedule" && (
          <div style={S.calContainer}>
            <h2 style={S.sectionTitle}>全部教練課表</h2>
            <p style={S.gridHint}>一覽所有教練同其他租場嘅預約。撳已預約嘅格協助取消；撳空格可落其他租場（包場／小組／試堂）。</p>
            <div style={S.weekNav}>
              <button style={S.navBtn} onClick={() => setWeekOffset((w) => w - 1)}>‹ 上週</button>
              <span style={S.weekLabel}>{formatDate(days[0])} – {formatDate(days[6])}</span>
              <button style={S.navBtn} onClick={() => setWeekOffset((w) => w + 1)}>下週 ›</button>
            </div>
            <div style={S.calScroll}>
              <table style={S.table}>
                <thead><tr><th style={S.thTime}></th>
                  {days.map((d) => { const closed = CLOSED_DAYS.includes(d.getDay()); const today = isTodayDate(d); return <th key={d} style={{ ...S.th, background: today ? "#13302e" : undefined }}><div style={{ ...S.dayLabel, color: closed ? "#5a3030" : undefined }}>{formatDay(d)}</div><div style={{ ...S.dateLabel, color: closed ? "#555" : today ? "#4ECDC4" : undefined }}>{d.getDate()}</div>{today ? <div style={S.todayTag}>今日</div> : closed ? <div style={S.closedTag}>休息</div> : null}</th>; })}
                </tr></thead>
                <tbody>
                  {TIME_SLOTS.map((time) => {
                    const isHourStart = time.endsWith(":00");
                    return (
                      <tr key={time}>
                        <td style={{ ...S.tdTime, color: isHourStart ? "#aaa" : "#3a3a3a" }}>{time}</td>
                        {days.map((d) => {
                          const date = formatDate(d);
                          const here = cellArr(date, time);
                          const occ = occupancy(date, time);
                          const whole = here.find(isWholeVenue);
                          const isPast = hoursUntil(date, time) < 0;
                          const closed = isClosedDay(date);
                          const canAdd = occ < MAX_CONCURRENT && !isPast && !closed;
                          return (
                            <td key={date} style={{ ...S.td, borderTop: isHourStart ? "1px solid #2a2a2a" : "1px solid #161616", background: closed && here.length === 0 ? "#0c0c0c" : undefined }}>
                              {whole ? (() => {
                                const span = Math.round(whole.hours * 4);
                                const relRow = slotIndex(time) - slotIndex(whole.start);
                                const lines = buildEntryLines(whole, false, null, false);
                                let node = null;
                                if (span === 1) node = <span style={lines[0].style}>{lines[0].text}</span>;
                                else if (relRow < span - 1 && relRow < lines.length) node = <span style={lines[relRow].style}>{lines[relRow].text}</span>;
                                else if (relRow === span - 1) node = <span style={S.slotBottomTime}>{addMinutes(whole.start, whole.hours * 60)}</span>;
                                return (
                                  <div style={{ ...S.slotChip, background: "#ffffff22", borderLeft: "3px solid #fff" }}>
                                    {node}
                                    {relRow === 0 && <button style={S.cancelSlotBtn} onClick={() => setAdminCancelModal({ date, start: whole.start, coachId: 0, type: "charter" })}>✕</button>}
                                  </div>
                                );
                              })() : here.length > 0 ? (
                                <div style={S.slotMulti}>
                                  {here.map((v, idx) => {
                                    const isTrial = v.type === "charter";
                                    const c = isTrial ? null : getCoach(v.coachId);
                                    const span = Math.round(v.hours * 4);
                                    const relRow = slotIndex(time) - slotIndex(v.start);
                                    const lines = buildEntryLines(v, isTrial, c, false);
                                    let node = null;
                                    if (span === 1) node = <span style={lines[0].style}>{lines[0].text}</span>;
                                    else if (relRow < span - 1 && relRow < lines.length) node = <span style={lines[relRow].style}>{lines[relRow].text}</span>;
                                    else if (relRow === span - 1) node = <span style={S.slotBottomTime}>{addMinutes(v.start, v.hours * 60)}</span>;
                                    return (
                                      <div key={idx} style={{ ...S.slotChip, background: isTrial ? "#ffffff22" : c?.color + "33", borderLeft: `3px solid ${isTrial ? "#fff" : c?.color}` }}>
                                        {node}
                                        {relRow === 0 && <button style={S.cancelSlotBtn} onClick={() => setAdminCancelModal({ date, start: v.start, coachId: v.coachId, type: v.type })}>✕</button>}
                                      </div>
                                    );
                                  })}
                                  {canAdd && <button style={S.slotAdd} onClick={() => setCharterModal({ date, time, charterType: "trial", hours: 1, price: 0, coachName: "" })}>+</button>}
                                </div>
                              ) : closed ? <div style={S.slotClosed} />
                                : isPast ? <div style={S.slotPast} />
                                : <button style={S.slotEmpty} onClick={() => setCharterModal({ date, time, charterType: "private", hours: 1, price: CHARTER_PRICE, coachName: "" })}>+</button>}
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <p style={S.assistHint}>² = 1對2　｜　白色 = 其他租場（包場／小組／試堂）</p>
          </div>
        )}

        {adminTab === "coaches" && (
          <div style={S.container}>
            <div style={S.flexBetween}>
              <h2 style={S.sectionTitle}>教練帳戶</h2>
              <button style={S.addBtn} onClick={() => setEditCoach({ id: null, username: "", name: "", credits: 0, rate: 200, password: "1234" })}>+ 新增教練</button>
            </div>
            <div style={{ ...S.filterRow, justifyContent: "flex-end" }}>
              <button style={S.linkBtn} onClick={() => setShowPasswords((v) => !v)}>{showPasswords ? "🙈 隱藏密碼" : "👁️ 顯示密碼"}</button>
            </div>
            <div style={S.bookingList}>
              {coaches.map((c) => (
                <div key={c.id} style={S.coachStatRow}>
                  <div style={{ ...S.avatar, background: c.color }}>{c.initials}</div>
                  <div style={{ flex: 1 }}>
                    <div style={S.bookingCoach}>{c.name} <span style={S.idTag}>@{c.username}</span></div>
                    <div style={S.bookingTime}>堂數 {c.used}/{c.credits}　每堂 ${c.rate}　密碼 {showPasswords ? c.password : "••••"}</div>
                  </div>
                  <button style={S.creditBtn} onClick={() => setAddCreditModal({ coachId: c.id, qty: 1, expiryDate: "" })}>+ 堂</button>
                  <button style={S.smallBtn} onClick={() => setEditCoach(c)}>編輯</button>
                  <button style={S.delBtn} onClick={() => setDelCoachModal(c)}>刪</button>
                </div>
              ))}
            </div>
          </div>
        )}

        {adminTab === "ledger" && (() => {
          let filtered = ledgerFilter === "all" ? purchaseLog : purchaseLog.filter((r) => String(r.coachId) === String(ledgerFilter));
          if (ledgerMonth !== "all") filtered = filtered.filter((r) => monthKey(r.date) === ledgerMonth);
          const filteredTotal = filtered.reduce((s, r) => s + r.amount, 0);
          return (
          <div style={S.container}>
            <h2 style={S.sectionTitle}>購買堂數流水帳</h2>
            <div style={S.filterRow}>
              <span style={S.filterLabel}>教練：</span>
              <select style={S.select} value={ledgerFilter} onChange={(e) => setLedgerFilter(e.target.value)}>
                <option value="all">全部</option>
                {coaches.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
              </select>
              <span style={{ ...S.filterLabel, marginLeft: 10 }}>月份：</span>
              {ledgerMonth === "all"
                ? <button style={S.smallBtn} onClick={() => setLedgerMonth(monthKey(formatDate(new Date())))}>選擇月份</button>
                : <><input style={S.select} type="month" value={ledgerMonth} onChange={(e) => setLedgerMonth(e.target.value)} />
                   <button style={S.linkBtn} onClick={() => setLedgerMonth("all")}>清除</button></>}
            </div>
            {filtered.length === 0 ? <p style={S.emptyText}>暫無購買記錄</p> : (
              <div style={S.bookingList}>
                {filtered.map((r) => (
                  <div key={r.id} style={S.bookingItem}>
                    <div style={{ ...S.dot, background: getCoach(r.coachId)?.color || "#666" }} />
                    <div style={{ flex: 1 }}>
                      <div style={S.bookingCoach}>{r.coachName} <span style={S.plusTag}>+{r.qty} 堂</span></div>
                      <div style={S.bookingTime}>{r.date}　@${r.rate}/堂</div>
                    </div>
                    <div style={{ textAlign: "right" }}>
                      <div style={S.revenueNum}>+${r.amount.toLocaleString()}</div>
                      <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
                        <button style={S.linkBtn} onClick={() => setEditDateRec({ id: r.id, date: r.date })}>改日期</button>
                        <button style={{ ...S.linkBtn, color: "#FF6B6B" }} onClick={() => setDelLedgerModal(r)}>剷除</button>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}
            {filtered.length > 0 && <div style={S.ledgerTotal}>{ledgerFilter === "all" ? "流水帳總額" : `${getCoach(ledgerFilter)?.name} 小計`}{ledgerMonth !== "all" ? `（${ledgerMonth}）` : ""}：${filteredTotal.toLocaleString()}</div>}
          </div>
          );
        })()}

        {adminTab === "records" && (
          <div style={S.container}>
            <h2 style={S.sectionTitle}>記錄</h2>
            <div style={S.segRow}>
              <button style={recordsView === "bookings" ? S.segActive : S.seg} onClick={() => setRecordsView("bookings")}>上堂記錄</button>
              <button style={recordsView === "cancelled" ? S.segActive : S.seg} onClick={() => setRecordsView("cancelled")}>取消記錄 {cancelLog.length > 0 && <span style={S.badge}>{cancelLog.length}</span>}</button>
            </div>

            {recordsView === "bookings" ? (() => {
              const now = new Date();
              const typeOf = (b) => b.type === "charter" ? (b.charterType || "private") : b.type;
              let list = allBookings.filter((b) => {
                if (recCoach !== "all") { if (b.type === "charter") return false; if (String(b.coachId) !== String(recCoach)) return false; }
                if (recType !== "all" && typeOf(b) !== recType) return false;
                const isPast = new Date(`${b.date}T${b.start}:00`) < now;
                if (recRange === "upcoming" && isPast) return false;
                if (recRange === "past" && !isPast) return false;
                if (recRange === "month" && monthKey(b.date) !== recMonth) return false;
                return true;
              });
              list.sort((a, b) => recRange === "upcoming"
                ? `${a.date}${a.start}`.localeCompare(`${b.date}${b.start}`)
                : `${b.date}${b.start}`.localeCompare(`${a.date}${a.start}`));
              const sumRevenue = list.reduce((s, b) => s + (b.price || 0), 0);
              return (
              <>
                <div style={{ ...S.filterWrap, marginTop: 14 }}>
                  <select style={S.select} value={recCoach} onChange={(e) => setRecCoach(e.target.value)}>
                    <option value="all">全部教練</option>
                    {coaches.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
                  </select>
                  <select style={S.select} value={recType} onChange={(e) => setRecType(e.target.value)}>
                    <option value="all">全部類型</option>
                    <option value="solo">一對一</option>
                    <option value="duo">一對二</option>
                    <option value="private">私人包場</option>
                    <option value="group">小組訓練</option>
                    <option value="trial">試堂</option>
                  </select>
                  <select style={S.select} value={recRange} onChange={(e) => setRecRange(e.target.value)}>
                    <option value="upcoming">即將</option>
                    <option value="past">已完成</option>
                    <option value="month">指定年月</option>
                    <option value="all">全部</option>
                  </select>
                  {recRange === "month" && <input style={S.select} type="month" value={recMonth} onChange={(e) => setRecMonth(e.target.value)} />}
                </div>
                <div style={S.recSummary}>共 {list.length} 項　｜　收入 ${sumRevenue.toLocaleString()}</div>
                <div style={S.bookingList}>
                  {list.length === 0 ? <p style={S.emptyText}>冇符合條件嘅記錄</p> : list.map((b, i) => {
                    const { date, start, hours, type, charterType, price, coachName, coach, coachId } = b;
                    const key = `${date}_${start}_${coachId}_${type}`;
                    const open = recExpanded === key;
                    const isPast = new Date(`${date}T${start}:00`) < now;
                    return (
                    <div key={i} style={S.bookingItem}>
                      <div style={{ ...S.dot, background: type === "charter" ? "#fff" : coach?.color }} />
                      <div style={{ flex: 1, cursor: "pointer" }} onClick={() => setRecExpanded(open ? null : key)}>
                        <div style={S.bookingCoach}>
                          {type === "charter" ? rentalFull(charterType) : coach?.name}{" "}
                          <span style={type === "duo" ? S.duoTag : type === "charter" ? S.charterTag : S.soloTag}>
                            {type === "duo" ? "1對2" : type === "charter" ? (charterType === "trial" ? "試堂" : `$${price}`) : "1對1"}
                          </span>
                          {isPast && <span style={S.donePill}>已完成</span>}
                        </div>
                        <div style={S.bookingTime}>{date}（{formatDay(new Date(`${date}T00:00:00`))}） · {start}–{addMinutes(start, hours * 60)}（{hours}小時）{type === "charter" && coachName ? `　負責：${coachName}` : ""}</div>
                        {open && (
                          <div style={S.recDetail}>
                            <div>類型：{type === "charter" ? rentalFull(charterType) : type === "duo" ? "一對二" : "一對一"}</div>
                            <div>收費：{type === "charter" && charterType === "trial" ? "免費" : `$${price}`}</div>
                            {type !== "charter" && <div>扣堂數：{hours} 堂</div>}
                            <div>落單時間：{b.createdAt || "—（舊記錄）"}</div>
                          </div>
                        )}
                      </div>
                      {!isPast && <button style={S.delBtn} onClick={() => setAdminCancelModal({ date, start, coachId, type })}>取消</button>}
                    </div>
                    );
                  })}
                </div>
                <p style={S.assistHint}>※ 撳記錄可展開詳情；管理員可協助取消未開始嘅時段（包括24小時內）。</p>
              </>
              );
            })() : (
              <>
                <div style={{ ...S.bookingList, marginTop: 16 }}>
                  {cancelLog.length === 0 ? <p style={S.emptyText}>暫無取消記錄</p> : cancelLog.map((r, i) => (
                    <div key={i} style={S.bookingItem}>
                      <div style={{ ...S.dot, background: "#FF6B6B" }} />
                      <div style={{ flex: 1 }}>
                        <div style={S.bookingCoach}>
                          {r.type === "charter" ? rentalFull(r.charterType) : r.coachName}{" "}
                          <span style={S.cancelledTag}>{r.cancelledBy === "admin" ? "管理員代取消" : "教練自行取消"}</span>
                        </div>
                        <div style={S.bookingTime}>原定 {r.date} · {r.start}–{addMinutes(r.start, r.hours * 60)}（{r.hours}小時）{r.price ? `　$${r.price}` : ""}</div>
                        <div style={S.bookingTime}>取消於 {r.cancelledAt}</div>
                      </div>
                    </div>
                  ))}
                </div>
                <p style={S.assistHint}>※ 留底紀錄，方便日後查核某時段點解空出，唔可以還原。</p>
              </>
            )}
          </div>
        )}

        {adminTab === "settings" && (
          <div style={S.container}>
            <h2 style={S.sectionTitle}>修改{isSubAdmin ? "我的" : "管理員"}密碼</h2>
            <div style={S.formCard}>
              <Field label="舊密碼"><input style={S.input} type="password" value={pwForm.old} onChange={(e) => setPwForm({ ...pwForm, old: e.target.value })} /></Field>
              <Field label="新密碼"><input style={S.input} type="password" value={pwForm.new1} onChange={(e) => setPwForm({ ...pwForm, new1: e.target.value })} /></Field>
              <Field label="確認新密碼"><input style={S.input} type="password" value={pwForm.new2} onChange={(e) => setPwForm({ ...pwForm, new2: e.target.value })} /></Field>
              <button style={S.loginBtn} onClick={changePassword}>更新密碼</button>
            </div>
            <p style={S.assistHint}>※ 收費：1對1 用教練每堂租金；1對2 $150/小時，每加0.5小時 +$50</p>

            <h2 style={{ ...S.sectionTitle, marginTop: 28 }}>匯出資料備份</h2>
            <div style={S.formCard}>
              <p style={{ ...S.bookingTime, marginBottom: 14, lineHeight: 1.6 }}>匯出 Excel，包含教練總覽、全部流水帳、每個教練獨立流水帳、上堂記錄、包場小組記錄。建議定期備份。</p>
              <button style={{ ...S.loginBtn, background: "#6BCB77" }} onClick={exportExcel}>⬇️ 匯出 Excel 備份</button>
              <button style={{ ...S.loginBtn, background: "#2a2a2a", color: "#fff", marginTop: 10 }} onClick={copyLedgerCsv}>📋 複製流水帳 (CSV)</button>
              <p style={{ ...S.assistHint, marginTop: 10 }}>※ 若下載冇反應（手機 app 常見），可改按「複製流水帳」再貼入 Excel / Google Sheets；或喺電腦瀏覽器開啟再匯出。</p>
            </div>

            {currentUser.role === "admin" && (
              <>
                <h2 style={{ ...S.sectionTitle, marginTop: 28 }}>副管理員帳戶</h2>
                <p style={S.gridHint}>副管理員可登入並使用下面開啟咗嘅分頁，但唔可以管理副管理員帳戶本身或重設資料。</p>
                <div style={S.bookingList}>
                  {subAdmins.map((s) => (
                    <div key={s.id} style={S.formCard}>
                      <Field label="顯示名稱"><input style={S.input} value={s.name} onChange={(e) => setSubAdmins((prev) => prev.map((x) => x.id === s.id ? { ...x, name: e.target.value } : x))} /></Field>
                      <Field label="登入帳號名稱"><input style={S.input} value={s.username} onChange={(e) => {
                        const v = e.target.value.trim().toLowerCase();
                        if (v === "admin") { showToast("帳號名稱不可用 admin", "error"); return; }
                        const dup = coaches.some((c) => (c.username || "").toLowerCase() === v) || subAdmins.some((x) => x.id !== s.id && (x.username || "").toLowerCase() === v);
                        if (dup) { showToast("帳號名稱已被使用", "error"); return; }
                        setSubAdmins((prev) => prev.map((x) => x.id === s.id ? { ...x, username: e.target.value } : x));
                      }} /></Field>
                      <Field label="密碼"><input style={S.input} value={s.password} onChange={(e) => setSubAdmins((prev) => prev.map((x) => x.id === s.id ? { ...x, password: e.target.value } : x))} /></Field>
                      <label style={S.label}>可使用分頁</label>
                      <div style={{ ...S.checkRow, flexWrap: "wrap", rowGap: 8 }}>
                        {[["overview", "總覽"], ["schedule", "課表"], ["coaches", "教練"], ["ledger", "流水帳"], ["records", "記錄"], ["settings", "設定"]].map(([k, label]) => (
                          <label key={k} style={S.checkLabel}>
                            <input type="checkbox" checked={!!s.permissions?.[k]}
                              onChange={(e) => setSubAdmins((prev) => prev.map((x) => x.id === s.id ? { ...x, permissions: { ...x.permissions, [k]: e.target.checked } } : x))} /> {label}
                          </label>
                        ))}
                      </div>
                      <button style={{ ...S.delBtn, marginTop: 12 }} onClick={() => { setSubAdmins((prev) => prev.filter((x) => x.id !== s.id)); showToast("已刪除副管理員"); }}>刪除呢個帳戶</button>
                    </div>
                  ))}
                </div>
                <button style={{ ...S.addBtn, marginTop: 12 }} onClick={() => {
                  const newId = Math.max(0, ...subAdmins.map((s) => s.id)) + 1;
                  setSubAdmins((prev) => [...prev, { id: newId, username: `subadmin${newId}`, password: "1234", name: `副管理員${newId}`, permissions: { overview: true, schedule: true, coaches: true, ledger: true, records: true, settings: true } }]);
                }}>+ 新增副管理員</button>
              </>
            )}

            {currentUser.role === "admin" && (
              <>
                <h2 style={{ ...S.sectionTitle, marginTop: 28 }}>重設資料</h2>
                <div style={S.formCard}>
                  <p style={{ ...S.bookingTime, marginBottom: 14, lineHeight: 1.6 }}>清除呢部裝置嘅所有資料，回復至初始狀態。建議先匯出備份。此動作無法復原。</p>
                  <button style={{ ...S.loginBtn, background: "#FF6B6B", color: "#fff" }} onClick={() => setResetModal(true)}>🗑️ 重設所有資料</button>
                </div>
              </>
            )}
          </div>
        )}

        {editCoach && (
          <EditCoachModal coach={editCoach} onClose={() => setEditCoach(null)}
            onSave={(data) => {
              const uname = (data.username || "").trim().toLowerCase();
              if (!uname) { showToast("請輸入帳號名稱", "error"); return; }
              if (uname === "admin") { showToast("帳號名稱不可用 admin", "error"); return; }
              const dup = coaches.some((c) => c.id !== data.id && (c.username || "").toLowerCase() === uname);
              if (dup) { showToast("帳號名稱已被使用", "error"); return; }
              const clean = { ...data, username: uname };
              if (clean.id) { setCoaches((prev) => prev.map((c) => c.id === clean.id ? { ...c, ...clean } : c)); showToast("已更新教練"); }
              else {
                const newId = Math.max(0, ...coaches.map((c) => c.id)) + 1;
                const color = COLORS[coaches.length % COLORS.length];
                const initials = clean.name.trim().split(/\s+/).map((w) => w[0]).join("").slice(0, 2).toUpperCase() || "NA";
                setCoaches((prev) => [...prev, { ...clean, id: newId, color, initials, used: 0 }]);
                if (clean.credits > 0) setPurchaseLog((prev) => [{ id: "p" + Date.now() + "-" + Math.random().toString(36).slice(2), date: new Date().toISOString().slice(0, 10), coachId: newId, coachName: clean.name, qty: clean.credits, amount: clean.credits * clean.rate, rate: clean.rate }, ...prev]);
                showToast(`已新增教練 ${clean.name}（@${uname}）`);
              }
              setEditCoach(null);
            }} />
        )}

        {addCreditModal && (
          <div style={S.modalOverlay}><div style={S.modal}>
            <h3 style={S.modalTitle}>增加堂數</h3>
            <p style={S.modalText}>{getCoach(addCreditModal.coachId)?.name}　每堂 ${getCoach(addCreditModal.coachId)?.rate}</p>
            <Field label="增加幾多堂"><input style={S.input} type="number" min="1" value={addCreditModal.qty} onChange={(e) => setAddCreditModal({ ...addCreditModal, qty: parseInt(e.target.value) || 1 })} /></Field>
            <Field label="失效日期（留空＝無限期）"><input style={S.input} type="date" value={addCreditModal.expiryDate || ""} onChange={(e) => setAddCreditModal({ ...addCreditModal, expiryDate: e.target.value })} /></Field>
            <p style={S.amountPreview}>金額：${((getCoach(addCreditModal.coachId)?.rate || 0) * addCreditModal.qty).toLocaleString()}</p>
            <div style={S.modalBtns}>
              <button style={S.modalCancel} onClick={() => setAddCreditModal(null)}>取消</button>
              <button style={S.modalConfirm} onClick={() => { addCredits(addCreditModal.coachId, addCreditModal.qty, addCreditModal.expiryDate); setAddCreditModal(null); }}>確認增加</button>
            </div>
          </div></div>
        )}

        {editDateRec && (
          <div style={S.modalOverlay}><div style={S.modal}>
            <h3 style={S.modalTitle}>修改入數日期</h3>
            <Field label="入數日期"><input style={S.input} type="date" value={editDateRec.date} onChange={(e) => setEditDateRec({ ...editDateRec, date: e.target.value })} /></Field>
            <div style={S.modalBtns}>
              <button style={S.modalCancel} onClick={() => setEditDateRec(null)}>取消</button>
              <button style={S.modalConfirm} onClick={() => {
                setPurchaseLog((prev) => prev.map((r) => r.id === editDateRec.id ? { ...r, date: editDateRec.date } : r));
                showToast("已更新入數日期");
                setEditDateRec(null);
              }}>儲存</button>
            </div>
          </div></div>
        )}

        {charterModal && (
          <div style={S.modalOverlay}><div style={{ ...S.modal, textAlign: "left" }}>
            <h3 style={{ ...S.modalTitle, textAlign: "center" }}>其他租場</h3>
            <p style={{ ...S.modalText, textAlign: "center" }}>{charterModal.date}　{charterModal.time}</p>

            <label style={S.label}>類型</label>
            <div style={S.segRow}>
              <button style={charterModal.charterType === "private" ? S.segActive : S.seg} onClick={() => setCharterModal({ ...charterModal, charterType: "private", price: charterModal.charterType === "trial" ? CHARTER_PRICE : charterModal.price })}>私人包場</button>
              <button style={charterModal.charterType === "group" ? S.segActive : S.seg} onClick={() => setCharterModal({ ...charterModal, charterType: "group", price: charterModal.charterType === "trial" ? CHARTER_PRICE : charterModal.price })}>小組訓練</button>
              <button style={charterModal.charterType === "trial" ? S.segActive : S.seg} onClick={() => setCharterModal({ ...charterModal, charterType: "trial", price: 0 })}>試堂</button>
            </div>
            {charterModal.charterType === "trial"
              ? <p style={{ ...S.assistHint, marginTop: 6 }}>試堂只佔 1 個位，同一時段仲可以有教練 book，唔收費。</p>
              : <p style={{ ...S.assistHint, marginTop: 6 }}>包場／小組會獨佔全場（2 位）。</p>}

            <label style={{ ...S.label, marginTop: 14 }}>時長</label>
            <div style={S.segRow}>
              {[1, 1.5, 2].map((h) => (
                <button key={h} style={charterModal.hours === h ? S.segActive : S.seg} onClick={() => setCharterModal({ ...charterModal, hours: h })}>{h} 小時</button>
              ))}
              <button style={![1, 1.5, 2].includes(charterModal.hours) ? S.segActive : S.seg} onClick={() => setCharterModal({ ...charterModal, hours: 3 })}>其他</button>
            </div>
            {![1, 1.5, 2].includes(charterModal.hours) && (
              <div style={{ marginTop: 8 }}>
                <label style={S.label}>自訂時長（小時，可 0.25 為一格）</label>
                <input style={S.input} type="number" step="0.25" min="0.25" value={charterModal.hours}
                  onChange={(e) => setCharterModal({ ...charterModal, hours: parseFloat(e.target.value) || 0.25 })} />
              </div>
            )}

            <label style={{ ...S.label, marginTop: 14 }}>負責教練</label>
            <select style={{ ...S.select, width: "100%", boxSizing: "border-box" }} value={charterModal.coachName} onChange={(e) => setCharterModal({ ...charterModal, coachName: e.target.value })}>
              <option value="">（未指定）</option>
              {coaches.map((c) => <option key={c.id} value={c.name}>{c.name}</option>)}
            </select>

            {charterModal.charterType === "trial" ? (
              <p style={{ ...S.amountPreview, color: "#999", marginTop: 14 }}>試堂不收費，唔會計入收入。</p>
            ) : (
              <>
                <label style={{ ...S.label, marginTop: 14 }}>收費 ($，可自由修改)</label>
                <input style={S.input} type="number" min="0" value={charterModal.price}
                  onChange={(e) => setCharterModal({ ...charterModal, price: e.target.value })} />
              </>
            )}

            <div style={S.priceBox}>
              <div style={S.priceRow}><span>時段</span><span>{charterModal.time} – {addMinutes(charterModal.time, charterModal.hours * 60)}</span></div>
              <div style={S.priceRow}><span>場地</span><span>{charterModal.charterType === "trial" ? "佔 1 位（可同教練並存）" : "全場獨佔"}</span></div>
              <div style={{ ...S.priceRow, color: "#4ECDC4", fontWeight: 700, fontSize: 16 }}><span>收費</span><span>{charterModal.charterType === "trial" ? "免費" : `$${parseInt(charterModal.price) || 0}`}</span></div>
            </div>
            <div style={S.modalBtns}>
              <button style={S.modalCancel} onClick={() => setCharterModal(null)}>返回</button>
              <button style={S.modalConfirm} onClick={confirmCharter}>確認落單</button>
            </div>
          </div></div>
        )}

        {adminCancelModal && (() => {
          const win = adminCancelModal.type !== "charter" ? (getCoach(adminCancelModal.coachId)?.cancelWindowHours ?? 24) : 24;
          const within24 = adminCancelModal.type !== "charter" && hoursUntil(adminCancelModal.date, adminCancelModal.start) < win;
          const used = adminCancelModal.type !== "charter" ? assistUsedThisMonth(adminCancelModal.coachId) : 0;
          const over = within24 && used >= ASSIST_CANCEL_LIMIT;
          return (
          <div style={S.modalOverlay}><div style={S.modal}>
            <h3 style={S.modalTitle}>確認取消</h3>
            <p style={S.modalText}>{adminCancelModal.date}　{adminCancelModal.start}<br />確定幫呢個時段取消？{adminCancelModal.type !== "charter" ? "（會退回對應堂數）" : ""}</p>
            {within24 && (
              <div style={{ ...S.quotaBox, borderColor: over ? "#5a2020" : "#1d3a2a", background: over ? "#2a1414" : "#13261c" }}>
                <div style={{ fontSize: 13, color: over ? "#FF8FA3" : "#6BCB77", fontWeight: 700 }}>
                  {over ? "⚠️ 已超出本月代取消額度" : "✓ 屬本月代取消額度範圍"}
                </div>
                <div style={{ fontSize: 12, color: "#999", marginTop: 4 }}>
                  {getCoach(adminCancelModal.coachId)?.name}　本月已用 {used} / {ASSIST_CANCEL_LIMIT} 次（通知時數 {win} 小時）
                  {over ? "，今次將超額。" : "。"}
                </div>
              </div>
            )}
            <div style={S.modalBtns}>
              <button style={S.modalCancel} onClick={() => setAdminCancelModal(null)}>返回</button>
              <button style={S.modalConfirm} onClick={() => { doCancel(adminCancelModal.date, adminCancelModal.start, adminCancelModal.coachId, adminCancelModal.type, true); setAdminCancelModal(null); }}>確認取消</button>
            </div>
          </div></div>
          );
        })()}

        {delLedgerModal && (
          <div style={S.modalOverlay}><div style={S.modal}>
            <h3 style={S.modalTitle}>剷除流水記錄</h3>
            <p style={S.modalText}>{delLedgerModal.coachName}　+{delLedgerModal.qty} 堂　${delLedgerModal.amount.toLocaleString()}<br />（{delLedgerModal.date}）<br /><br />確定剷除？教練堂數會相應扣減，此動作無法復原。</p>
            <div style={S.modalBtns}>
              <button style={S.modalCancel} onClick={() => setDelLedgerModal(null)}>返回</button>
              <button style={{ ...S.modalConfirm, background: "#FF6B6B" }} onClick={() => {
                const rec = delLedgerModal;
                setPurchaseLog((prev) => prev.filter((x) => x.id !== rec.id));
                setCoaches((prev) => prev.map((c) => c.id === rec.coachId ? { ...c, credits: Math.max(c.used, c.credits - rec.qty) } : c));
                showToast("已剷除流水記錄");
                setDelLedgerModal(null);
              }}>確認剷除</button>
            </div>
          </div></div>
        )}

        {delCoachModal && (() => {
          let bookingCount = 0;
          Object.entries(bookings).forEach(([k, arr]) => {
            const date = k.split("_")[0];
            arr.forEach((e) => { if (e.coachId === delCoachModal.id && e.type !== "charter" && k === `${date}_${e.start}`) bookingCount++; });
          });
          const purCount = purchaseLog.filter((r) => r.coachId === delCoachModal.id).length;
          return (
          <div style={S.modalOverlay}><div style={S.modal}>
            <h3 style={S.modalTitle}>刪除教練</h3>
            <p style={S.modalText}>
              確定刪除 <b style={{ color: "#fff" }}>{delCoachModal.name}</b>（@{delCoachModal.username}）？<br /><br />
              此教練名下仍有 <b style={{ color: "#FFB347" }}>{bookingCount}</b> 個已預約時段、<b style={{ color: "#FFB347" }}>{purCount}</b> 筆購買記錄。<br />
              刪除後帳號即時失效，歷史記錄仍會保留（顯示為空白教練）。此動作無法復原。
            </p>
            <div style={S.modalBtns}>
              <button style={S.modalCancel} onClick={() => setDelCoachModal(null)}>返回</button>
              <button style={{ ...S.modalConfirm, background: "#FF6B6B" }} onClick={() => {
                setCoaches((prev) => prev.filter((x) => x.id !== delCoachModal.id));
                showToast("已刪除教練");
                setDelCoachModal(null);
              }}>確認刪除</button>
            </div>
          </div></div>
          );
        })()}

        {resetModal && (
          <div style={S.modalOverlay}><div style={S.modal}>
            <h3 style={S.modalTitle}>重設所有資料</h3>
            <p style={S.modalText}>確定清除呢部裝置嘅所有資料（教練、預約、流水帳等），回復至初始狀態？<br /><br />此動作無法復原，建議先匯出備份。</p>
            <div style={S.modalBtns}>
              <button style={S.modalCancel} onClick={() => setResetModal(false)}>返回</button>
              <button style={{ ...S.modalConfirm, background: "#FF6B6B" }} onClick={() => {
                try { localStorage.removeItem(LS_KEY); } catch (e) { /* ignore */ }
                setCoaches(DEFAULT_COACHES); setAdminPassword("admin123"); setBookings({});
                setSubAdmins([
                  { id: 1, username: "subadmin1", password: "1234", name: "副管理員1", permissions: { overview: true, schedule: true, coaches: true, ledger: true, records: true, settings: true } },
                  { id: 2, username: "subadmin2", password: "1234", name: "副管理員2", permissions: { overview: true, schedule: true, coaches: true, ledger: true, records: true, settings: true } },
                ]);
                setPurchaseLog([]); setCharterLog([]); setAssistCancelLog([]); setCancelLog([]);
                setResetModal(false); showToast("已重設資料");
              }}>確認重設</button>
            </div>
          </div></div>
        )}
        {toast && <Toast toast={toast} />}
      </div>
    );
  }
  return (
    <div style={S.appBg}>
      <Header title={`你好，${liveUser.name}`} onLogout={logout} syncState={syncState} />
      <div style={S.creditBar}>
        <span>已購買 {liveUser.credits} 堂</span>
        <span style={{ color: remaining > 0 ? "#4ECDC4" : "#FF6B6B", fontWeight: 700 }}>剩餘 {remaining} 堂</span>
      </div>
      {(() => {
        const used = assistUsedThisMonth(currentUser.id);
        const left = Math.max(0, ASSIST_CANCEL_LIMIT - used);
        return (
          <div style={S.assistBar}>
            <span>本月 {liveUser.cancelWindowHours ?? 24} 小時內代取消額度</span>
            <span style={{ color: left > 0 ? "#4ECDC4" : "#FF6B6B", fontWeight: 700 }}>剩 {left} / {ASSIST_CANCEL_LIMIT} 次</span>
          </div>
        );
      })()}
      {soldOut && <div style={S.soldOutBanner}>⚠️ 你已用晒購買堂數，請聯絡管理員增購後再預約</div>}
      {isCoach && (() => {
        const exp = expiringBatchesOf(currentUser.id);
        if (exp.length === 0) return null;
        return (
          <div style={S.expiryBanner}>
            ⏰ 你有 {exp.reduce((a, b) => a + b.remaining, 0)} 堂將於 {exp.map((b) => b.expiryDate).join("、")} 失效，請盡快預約使用
          </div>
        );
      })()}
      <div style={S.tabRow}>
        <button style={view === "calendar" ? S.tabActive : S.tab} onClick={() => setView("calendar")}>📅 預約場地</button>
        <button style={view === "myBookings" ? S.tabActive : S.tab} onClick={() => setView("myBookings")}>📋 我的預約 {myBookings.length > 0 && <span style={S.badge}>{myBookings.length}</span>}</button>
        <button style={view === "income" ? S.tabActive : S.tab} onClick={() => setView("income")}>📈 上堂情況</button>
        <button style={view === "pw" ? S.tabActive : S.tab} onClick={() => setView("pw")}>🔑 改密碼</button>
      </div>

      {view === "calendar" && (
        <div style={S.calContainer}>
          <div style={S.weekNav}>
            <button style={S.navBtn} onClick={() => setWeekOffset((w) => w - 1)}>‹ 上週</button>
            <span style={S.weekLabel}>{formatDate(days[0])} – {formatDate(days[6])}</span>
            <button style={S.navBtn} onClick={() => setWeekOffset((w) => w + 1)}>下週 ›</button>
          </div>
          <p style={S.gridHint}>每格 15 分鐘　｜　同一時段最多 2 名教練　｜　按空格揀 1對1/1對2 同時長</p>
          <div style={S.calScroll}>
            <table style={S.table}>
              <thead><tr><th style={S.thTime}></th>
                {days.map((d) => { const closed = CLOSED_DAYS.includes(d.getDay()); const today = isTodayDate(d); return <th key={d} style={{ ...S.th, background: today ? "#13302e" : undefined }}><div style={{ ...S.dayLabel, color: closed ? "#5a3030" : undefined }}>{formatDay(d)}</div><div style={{ ...S.dateLabel, color: closed ? "#555" : today ? "#4ECDC4" : undefined }}>{d.getDate()}</div>{today ? <div style={S.todayTag}>今日</div> : closed ? <div style={S.closedTag}>休息</div> : null}</th>; })}
              </tr></thead>
              <tbody>
                {TIME_SLOTS.map((time) => {
                  const isHourStart = time.endsWith(":00");
                  return (
                    <tr key={time}>
                      <td style={{ ...S.tdTime, color: isHourStart ? "#aaa" : "#3a3a3a" }}>{time}</td>
                      {days.map((d) => {
                        const date = formatDate(d);
                        const here = cellArr(date, time);
                        const occ = occupancy(date, time);
                        const whole = here.find(isWholeVenue);
                        const isPast = hoursUntil(date, time) < 0;
                        const closed = isClosedDay(date);
                        const iAmHere = here.some((v) => v.coachId === currentUser.id && v.type !== "charter");
                        const canAddHere = !whole && occ < MAX_CONCURRENT && !iAmHere && !isPast && !soldOut && !closed;
                        return (
                          <td key={date} style={{ ...S.td, borderTop: isHourStart ? "1px solid #2a2a2a" : "1px solid #161616", background: closed && here.length === 0 ? "#0c0c0c" : undefined }}>
                            {whole ? (() => {
                              const span = Math.round(whole.hours * 4);
                              const relRow = slotIndex(time) - slotIndex(whole.start);
                              const lines = buildEntryLines(whole, false, null, false);
                              let node = null;
                              if (span === 1) node = <span style={lines[0].style}>{lines[0].text}</span>;
                              else if (relRow < span - 1 && relRow < lines.length) node = <span style={lines[relRow].style}>{lines[relRow].text}</span>;
                              else if (relRow === span - 1) node = <span style={S.slotBottomTime}>{addMinutes(whole.start, whole.hours * 60)}</span>;
                              return <div style={{ ...S.slotChip, background: "#ffffff22", borderLeft: "3px solid #fff" }}>{node}</div>;
                            })() : here.length > 0 ? (
                              <div style={S.slotMulti}>
                                {here.map((v, idx) => {
                                  const isTrial = v.type === "charter";
                                  const c = isTrial ? null : getCoach(v.coachId);
                                  const isOwner = currentUser.role === "coach" && v.coachId === currentUser.id;
                                  const span = Math.round(v.hours * 4);
                                  const relRow = slotIndex(time) - slotIndex(v.start);
                                  const lines = buildEntryLines(v, isTrial, c, isOwner);
                                  let node = null;
                                  if (span === 1) node = <span style={lines[0].style}>{lines[0].text}</span>;
                                  else if (relRow < span - 1 && relRow < lines.length) node = <span style={lines[relRow].style}>{lines[relRow].text}</span>;
                                  else if (relRow === span - 1) node = <span style={S.slotBottomTime}>{addMinutes(v.start, v.hours * 60)}</span>;
                                  const showCancel = relRow === 0 && !isTrial && v.coachId === currentUser.id && hoursUntil(date, v.start) >= (liveUser.cancelWindowHours ?? 24) && !isPast;
                                  return (
                                    <div key={idx} style={{ ...S.slotChip, background: isTrial ? "#ffffff22" : c?.color + "33", borderLeft: `3px solid ${isTrial ? "#fff" : c?.color}` }}>
                                      {node}
                                      {showCancel && <button style={S.cancelSlotBtn} onClick={() => openCancel(date, v.start, v.coachId, v.type)}>✕</button>}
                                    </div>
                                  );
                                })}
                                {canAddHere && <button style={S.slotAdd} onClick={() => openBook(date, time)}>+</button>}
                              </div>
                            ) : closed ? <div style={S.slotClosed} />
                              : isPast ? <div style={S.slotPast} />
                              : soldOut ? <div style={S.slotDisabled} />
                              : <button style={S.slotEmpty} onClick={() => openBook(date, time)}>+</button>}
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <p style={S.assistHint}>可預約任何時段；24小時內取消需管理員協助。² = 1對2</p>
        </div>
      )}

      {view === "myBookings" && (
        <div style={S.container}>
          <h2 style={S.sectionTitle}>我的預約記錄</h2>
          {myBookings.length === 0 ? <p style={S.emptyText}>你還未有預約</p> : (
            <div style={S.bookingList}>
              {myBookings.map(({ date, start, hours, type, students, signatures }, i) => {
                const hrs = hoursUntil(date, start);
                const isPast = hrs < 0;
                const locked = hrs >= 0 && hrs < (liveUser.cancelWindowHours ?? 24);
                return (
                  <div key={i} style={S.bookingItem}>
                    <div style={{ ...S.dot, background: liveUser.color }} />
                    <div style={{ flex: 1 }}>
                      <div style={S.bookingCoach}>{date} <span style={type === "duo" ? S.duoTag : S.soloTag}>{type === "duo" ? "1對2" : "1對1"}</span></div>
                      <div style={S.bookingTime}>{start} – {addMinutes(start, hours * 60)}（{hours}小時）</div>
                      {students && students.length > 0 && (
                        <div style={S.signRow}>
                          {students.map((name) => {
                            const signed = signatures && signatures[name];
                            return (
                              <button key={name} style={signed ? S.signedChip : S.signChip} onClick={() => !signed && setSignModal({ date, start, coachId: currentUser.id, type, studentName: name })}>
                                {signed ? `✓ ${name}` : `✍️ ${name}`}
                              </button>
                            );
                          })}
                        </div>
                      )}
                    </div>
                    {isPast ? <span style={S.pastTag}>已完成</span>
                      : locked ? <span style={S.lockTag}>🔒 取消需管理員</span>
                      : <button style={S.cancelBtn} onClick={() => openCancel(date, start, currentUser.id, type)}>取消</button>}
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}

      {view === "income" && (() => {
        const roster = myRoster;
        return (
        <div style={S.container}>
          <h2 style={S.sectionTitle}>我的收入（近3個月）</h2>
          <p style={S.assistHint}>只計有填學生名嘅堂（用學生當時收費計）；冇填學生名嘅堂唔計入呢個收入。實際收入＝學生收費總額 － 租場費用（落單嗰刻嘅租金 × 時長）。</p>
          <div style={S.bookingList}>
            {myIncomeReport.months.map((m) => (
              <div key={m.month} style={S.monthCard}>
                <div style={S.monthHead}>{m.month}{m.month === monthKey(formatDate(new Date())) ? "（本月）" : ""}</div>
                <div style={S.monthRow}><span style={S.monthLabel}>計入收入嘅堂數</span><span>{m.count} 堂</span></div>
                <div style={S.monthRow}><span style={S.monthLabel}>學生收費總額</span><span>${m.gross.toLocaleString()}</span></div>
                <div style={S.monthRow}><span style={S.monthLabel}>租場費用</span><span style={{ color: "#FF8FA3" }}>-${m.rentalCost.toLocaleString()}</span></div>
                <div style={S.monthRow}><span style={{ ...S.monthLabel, fontWeight: 700, color: "#fff" }}>實際收入</span><span style={{ color: m.net >= 0 ? "#6BCB77" : "#FF6B6B", fontWeight: 700 }}>${m.net.toLocaleString()}</span></div>
              </div>
            ))}
          </div>

          <h2 style={{ ...S.sectionTitle, marginTop: 28 }}>學生名單</h2>
          <p style={S.assistHint}>撳學生名展開資料同上課紀錄。預約場地時可以直接揀名單入面嘅學生。</p>
          {roster.length === 0 && <p style={S.emptyText}>仲未有學生，落面新增啦</p>}
          <div style={S.bookingList}>
            {roster.map((s) => {
              const remain = (s.credits || 0) - (s.used || 0);
              const low = remain <= LOW_CREDIT_THRESHOLD;
              const open = studentLogOpen === s.name;
              const log = myIncomeReport.studentLog[s.name] || [];
              return (
                <div key={s.name}>
                  <div style={{ ...S.coachStatRow, cursor: "pointer" }} onClick={() => setStudentLogOpen(open ? null : s.name)}>
                    <div style={{ ...S.avatar, background: liveUser.color }}>{s.name.slice(0, 2)}</div>
                    <div style={{ flex: 1 }}>
                      <div style={S.bookingCoach}>{s.name}{low && <span style={S.lowPill}>低</span>}</div>
                      <div style={S.bookingTime}>每堂 ${s.rate || 0}　近3個月 {log.length} 堂</div>
                    </div>
                    <div style={{ textAlign: "right" }}>
                      <div style={{ color: low ? "#FF8FA3" : "#4ECDC4", fontWeight: 700 }}>剩 {remain} 堂</div>
                    </div>
                  </div>
                  {open && (
                    <div style={S.purchaseBreakdown}>
                      <div style={{ ...S.flexBetween, marginBottom: 4 }}>
                        <span style={S.assistHint}>學生資料</span>
                        <button style={S.rosterRemoveBtn} onClick={() => removeStudent(s.name)}>刪除學生</button>
                      </div>
                      <Field label="每堂收費 ($)"><input style={S.input} type="number" min="0" value={s.rate}
                        onChange={(e) => updateStudentField(s.name, "rate", parseInt(e.target.value) || 0)} /></Field>
                      <div style={S.bookingTime}>已開 {s.credits || 0} 堂　已用 {s.used || 0} 堂</div>
                      <Field label="剩餘堂數">
                        <input style={{ ...S.input, borderColor: low ? "#5a2020" : undefined, color: low ? "#FF8FA3" : "#4ECDC4", fontWeight: 700 }}
                          type="number" value={remain}
                          onChange={(e) => setStudentRemain(s.name, parseInt(e.target.value) || 0)} />
                      </Field>
                      <button style={{ ...S.creditBtn, marginTop: 4, marginBottom: 14 }} onClick={() => setAddStudentCreditModal({ name: s.name, qty: 1 })}>+ 幫佈開堂數</button>

                      <div style={{ ...S.assistHint, marginBottom: 4 }}>上課紀錄（近3個月）</div>
                      {log.length === 0 ? <p style={S.emptyText}>暫無紀錄</p> : log.map((l, i) => (
                        <div key={i} style={S.purchaseRow}>
                          <div style={S.bookingTime}>{l.date} · {l.start}–{addMinutes(l.start, l.hours * 60)}</div>
                          <div style={{ textAlign: "right" }}>
                            <span style={l.type === "duo" ? S.duoTag : S.soloTag}>{l.type === "duo" ? "1對2" : "1對1"}</span>
                            <div style={{ fontSize: 12, color: "#4ECDC4", marginTop: 2 }}>${l.charge}</div>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
            <input style={S.input} value={newStudentName} placeholder="新學生名" onChange={(e) => setNewStudentName(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && addStudentToRoster()} />
            <button style={{ ...S.creditBtn, whiteSpace: "nowrap" }} onClick={addStudentToRoster}>新增</button>
          </div>
        </div>
        );
      })()}

      {view === "pw" && (
        <div style={S.container}>
          <h2 style={S.sectionTitle}>修改我的密碼</h2>
          <div style={S.formCard}>
            <Field label="舊密碼"><input style={S.input} type="password" value={pwForm.old} onChange={(e) => setPwForm({ ...pwForm, old: e.target.value })} /></Field>
            <Field label="新密碼"><input style={S.input} type="password" value={pwForm.new1} onChange={(e) => setPwForm({ ...pwForm, new1: e.target.value })} /></Field>
            <Field label="確認新密碼"><input style={S.input} type="password" value={pwForm.new2} onChange={(e) => setPwForm({ ...pwForm, new2: e.target.value })} /></Field>
            <button style={S.loginBtn} onClick={changePassword}>更新密碼</button>
          </div>
        </div>
      )}

      {bookModal && (() => {
        const isDuo = bookModal.sessionType === "duo";
        const price = isDuo ? duoPrice(bookModal.hours) : liveUser.rate * bookModal.hours;
        const allowSolo = liveUser.allowSolo !== false;
        const allowDuo = liveUser.allowDuo !== false;
        return (
          <div style={S.modalOverlay}><div style={{ ...S.modal, textAlign: "left" }}>
            <h3 style={{ ...S.modalTitle, textAlign: "center" }}>預約場地</h3>
            <p style={{ ...S.modalText, textAlign: "center" }}>{bookModal.date}　{bookModal.time}</p>
            <label style={S.label}>類型</label>
            <div style={S.segRow}>
              <button style={!allowSolo ? S.segDisabled : !isDuo ? S.segActive : S.seg} disabled={!allowSolo} onClick={() => allowSolo && setBookModal({ ...bookModal, sessionType: "solo" })}>1對1</button>
              <button style={!allowDuo ? S.segDisabled : isDuo ? S.segActive : S.seg} disabled={!allowDuo} onClick={() => allowDuo && setBookModal({ ...bookModal, sessionType: "duo" })}>1對2</button>
            </div>
            <label style={{ ...S.label, marginTop: 14 }}>時長</label>
            <div style={S.segRow}>
              <button style={bookModal.hours === 1 ? S.segActive : S.seg} onClick={() => setBookModal({ ...bookModal, hours: 1 })}>1 小時</button>
              <button style={bookModal.hours === 1.5 ? S.segActive : S.seg} onClick={() => setBookModal({ ...bookModal, hours: 1.5 })}>1.5 小時</button>
            </div>
            <label style={{ ...S.label, marginTop: 14 }}>學生（最多4位，只有你自己睇到）</label>
            {myRoster.length === 0 ? (
              <p style={S.assistHint}>你仲未有學生名單，可以喺「上堂情況」分頁新增。</p>
            ) : (
              <div style={S.studentChipWrap}>
                {myRoster.map(({ name }) => {
                  const sel = Array.isArray(bookModal.students) && bookModal.students.includes(name);
                  const atMax = !sel && (bookModal.students || []).length >= 4;
                  return (
                    <button key={name} disabled={atMax} style={sel ? S.studentChipActive : atMax ? S.studentChipDisabled : S.studentChip}
                      onClick={() => {
                        const cur = bookModal.students || [];
                        setBookModal({ ...bookModal, students: sel ? cur.filter((n) => n !== name) : [...cur, name] });
                      }}>{name}</button>
                  );
                })}
              </div>
            )}
            {(bookModal.students || []).length < 4 && (
              <input style={{ ...S.input, marginTop: 8 }} value={bookModal.studentOther || ""} placeholder="其他（唔在名單，打名就得）"
                onChange={(e) => setBookModal({ ...bookModal, studentOther: e.target.value })} />
            )}
            <div style={S.priceBox}>
              <div style={S.priceRow}><span>時段</span><span>{bookModal.time} – {addMinutes(bookModal.time, bookModal.hours * 60)}</span></div>
              <div style={S.priceRow}><span>扣堂數</span><span>{bookModal.hours} 堂</span></div>
              <div style={{ ...S.priceRow, color: "#4ECDC4", fontWeight: 700, fontSize: 16 }}><span>{isDuo ? "1對2 收費" : "1對1 收費"}</span><span>${price}</span></div>
            </div>
            <div style={S.modalBtns}>
              <button style={S.modalCancel} onClick={() => setBookModal(null)}>返回</button>
              <button style={S.modalConfirm} onClick={confirmBook}>確認預約</button>
            </div>
          </div></div>
        );
      })()}

      {cancelModal && (
        <div style={S.modalOverlay}><div style={S.modal}>
          <h3 style={S.modalTitle}>確認取消</h3>
          <p style={S.modalText}>{cancelModal.date}　{cancelModal.start}<br />取消後退回對應堂數</p>
          <div style={S.modalBtns}>
            <button style={S.modalCancel} onClick={() => setCancelModal(null)}>返回</button>
            <button style={S.modalConfirm} onClick={() => doCancel(cancelModal.date, cancelModal.start, cancelModal.coachId, cancelModal.type)}>確認取消</button>
          </div>
        </div></div>
      )}
      {signModal && (
        <SignaturePad studentName={signModal.studentName}
          onCancel={() => setSignModal(null)}
          onSave={(dataUrl) => { signIn(signModal.date, signModal.start, signModal.coachId, signModal.type, signModal.studentName, dataUrl); setSignModal(null); }} />
      )}
      {addStudentCreditModal && (
        <div style={S.modalOverlay}><div style={S.modal}>
          <h3 style={S.modalTitle}>幫學生開堂數</h3>
          <p style={S.modalText}>{addStudentCreditModal.name}</p>
          <Field label="增加幾多堂"><input style={S.input} type="number" min="1" value={addStudentCreditModal.qty}
            onChange={(e) => setAddStudentCreditModal({ ...addStudentCreditModal, qty: parseInt(e.target.value) || 1 })} /></Field>
          <div style={S.modalBtns}>
            <button style={S.modalCancel} onClick={() => setAddStudentCreditModal(null)}>取消</button>
            <button style={S.modalConfirm} onClick={() => { addStudentCredits(addStudentCreditModal.name, addStudentCreditModal.qty); setAddStudentCreditModal(null); }}>確認增加</button>
          </div>
        </div></div>
      )}
      {toast && <Toast toast={toast} />}
    </div>
  );
}

function EditCoachModal({ coach, onClose, onSave }) {
  const [form, setForm] = useState({
    id: coach.id, username: coach.username || "", name: coach.name, credits: coach.credits, rate: coach.rate, password: coach.password,
    allowSolo: coach.allowSolo !== false, allowDuo: coach.allowDuo !== false, cancelWindowHours: coach.cancelWindowHours || 24,
  });
  return (
    <div style={S.modalOverlay}><div style={{ ...S.modal, width: 320, textAlign: "left" }}>
      <h3 style={S.modalTitle}>{coach.id ? "編輯教練" : "新增教練"}</h3>
      <Field label="教練全名（顯示用）"><input style={S.input} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>
      <Field label="登入帳號名稱"><input style={S.input} value={form.username} placeholder="例如 alex（登入用，不分大小寫）" onChange={(e) => setForm({ ...form, username: e.target.value })} /></Field>
      <Field label={coach.id ? "總購買堂數" : "初始購買堂數"}><input style={S.input} type="number" value={form.credits} onChange={(e) => setForm({ ...form, credits: parseInt(e.target.value) || 0 })} /></Field>
      <Field label="一對一每堂租金 ($)"><input style={S.input} type="number" value={form.rate} onChange={(e) => setForm({ ...form, rate: parseInt(e.target.value) || 0 })} /></Field>
      <Field label="密碼"><input style={S.input} value={form.password} onChange={(e) => setForm({ ...form, password: e.target.value })} /></Field>
      <label style={S.label}>預約權限</label>
      <div style={S.checkRow}>
        <label style={S.checkLabel}><input type="checkbox" checked={form.allowSolo} onChange={(e) => setForm({ ...form, allowSolo: e.target.checked })} /> 允許一對一</label>
        <label style={S.checkLabel}><input type="checkbox" checked={form.allowDuo} onChange={(e) => setForm({ ...form, allowDuo: e.target.checked })} /> 允許一對二</label>
      </div>
      <Field label="取消需管理員協助嘅時數（小時）"><input style={S.input} type="number" min="0" value={form.cancelWindowHours} onChange={(e) => setForm({ ...form, cancelWindowHours: parseInt(e.target.value) || 0 })} /></Field>
      {!coach.id && form.credits > 0 && <p style={S.amountPreview}>初始堂數記入流水帳：${(form.credits * form.rate).toLocaleString()}</p>}
      <div style={S.modalBtns}>
        <button style={S.modalCancel} onClick={onClose}>取消</button>
        <button style={S.modalConfirm} onClick={() => { if (!form.name.trim()) return; onSave(form); }}>儲存</button>
      </div>
    </div></div>
  );
}

function Field({ label, children }) { return <div style={S.inputGroup}><label style={S.label}>{label}</label>{children}</div>; }

// 簽名板：支援滑鼠同觸控（手機）畫簽名，輸出 base64 PNG
function SignaturePad({ studentName, onSave, onCancel }) {
  const canvasRef = useRef(null);
  const drawingRef = useRef(false);
  const hasDrawnRef = useRef(false);
  const lastPos = useRef({ x: 0, y: 0 });

  const getCtx = () => canvasRef.current?.getContext("2d");

  const posFromEvent = (e) => {
    const canvas = canvasRef.current;
    const rect = canvas.getBoundingClientRect();
    const point = e.touches ? e.touches[0] : e;
    return { x: point.clientX - rect.left, y: point.clientY - rect.top };
  };

  const start = (e) => {
    e.preventDefault();
    drawingRef.current = true;
    hasDrawnRef.current = true;
    lastPos.current = posFromEvent(e);
  };
  const move = (e) => {
    if (!drawingRef.current) return;
    e.preventDefault();
    const ctx = getCtx();
    const pos = posFromEvent(e);
    ctx.strokeStyle = "#000"; ctx.lineWidth = 2.5; ctx.lineCap = "round";
    ctx.beginPath(); ctx.moveTo(lastPos.current.x, lastPos.current.y); ctx.lineTo(pos.x, pos.y); ctx.stroke();
    lastPos.current = pos;
  };
  const end = () => { drawingRef.current = false; };

  const clear = () => {
    const canvas = canvasRef.current;
    const ctx = getCtx();
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, canvas.width, canvas.height);
    hasDrawnRef.current = false;
  };

  useEffect(() => {
    const canvas = canvasRef.current;
    canvas.width = canvas.offsetWidth * 2; canvas.height = canvas.offsetHeight * 2;
    const ctx = canvas.getContext("2d");
    ctx.scale(2, 2);
    ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, canvas.offsetWidth, canvas.offsetHeight);
  }, []);

  return (
    <div style={S.modalOverlay}><div style={{ ...S.modal, width: 320 }}>
      <h3 style={S.modalTitle}>學生簽到</h3>
      <p style={S.modalText}>{studentName}　請喺下面簽名作實</p>
      <canvas ref={canvasRef} style={S.signCanvas}
        onMouseDown={start} onMouseMove={move} onMouseUp={end} onMouseLeave={end}
        onTouchStart={start} onTouchMove={move} onTouchEnd={end} />
      <div style={S.modalBtns}>
        <button style={S.modalCancel} onClick={clear}>清除</button>
        <button style={S.modalCancel} onClick={onCancel}>返回</button>
        <button style={S.modalConfirm} onClick={() => { if (!hasDrawnRef.current) return; onSave(canvasRef.current.toDataURL("image/png")); }}>確認簽到</button>
      </div>
    </div></div>
  );
}
function Header({ title, onLogout, syncState }) {
  const sync = {
    synced: { t: "☁️ 已同步", c: "#6BCB77" },
    connecting: { t: "⟳ 同步中", c: "#FFB347" },
    error: { t: "⚠️ 同步失敗", c: "#FF6B6B" },
    local: { t: "📱 本機", c: "#777" },
  }[syncState] || null;
  return <div style={S.header}><img src={LOGO} alt="Gymily" style={S.headerLogoImg} /><span style={S.headerUser}>{title}</span>{sync && <span style={{ ...S.syncBadge, color: sync.c }}>{sync.t}</span>}<button style={S.logoutBtn} onClick={onLogout}>登出</button></div>;
}
function Toast({ toast }) { return <div style={{ ...S.toast, background: toast.type === "error" ? "#FF6B6B" : "#4ECDC4" }}>{toast.msg}</div>; }

const S = {
  loginBg: { minHeight: "100vh", background: "#0f0f0f", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'DM Sans', sans-serif" },
  loginCard: { background: "#1a1a1a", borderRadius: 20, padding: "40px 40px", width: 380, boxShadow: "0 20px 60px rgba(0,0,0,0.5)" },
  loginLogoImg: { width: 200, display: "block", margin: "0 auto 8px", filter: "drop-shadow(0 2px 8px rgba(0,0,0,0.4))" },
  loginSub: { color: "#888", textAlign: "center", marginBottom: 28, marginTop: 4 },
  inputGroup: { marginBottom: 16 },
  label: { display: "block", color: "#aaa", fontSize: 13, marginBottom: 6 },
  input: { width: "100%", background: "#2a2a2a", border: "1px solid #333", borderRadius: 10, padding: "12px 14px", color: "#fff", fontSize: 14, outline: "none", boxSizing: "border-box" },
  loginBtn: { width: "100%", background: "#4ECDC4", color: "#000", border: "none", borderRadius: 10, padding: "14px", fontSize: 16, fontWeight: 700, cursor: "pointer", marginTop: 8 },
  hint: { color: "#555", fontSize: 11, textAlign: "center", marginTop: 16, lineHeight: 1.6 },
  appBg: { minHeight: "100vh", background: "#0f0f0f", fontFamily: "'DM Sans', sans-serif", color: "#fff" },
  header: { display: "flex", alignItems: "center", padding: "12px 20px", background: "#1a1a1a", borderBottom: "1px solid #222" },
  headerLogoImg: { height: 36, flex: "0 0 auto", marginRight: 16 },
  syncBadge: { fontSize: 11, marginRight: 12, whiteSpace: "nowrap" },
  headerUser: { color: "#aaa", fontSize: 14, marginRight: 16, flex: 1 },
  logoutBtn: { background: "transparent", border: "1px solid #444", color: "#aaa", borderRadius: 8, padding: "6px 12px", fontSize: 13, cursor: "pointer" },
  creditBar: { display: "flex", justifyContent: "space-between", padding: "10px 20px", background: "#151515", fontSize: 13, color: "#aaa", borderBottom: "1px solid #222" },
  assistBar: { display: "flex", justifyContent: "space-between", padding: "8px 20px", background: "#121212", fontSize: 12, color: "#888", borderBottom: "1px solid #222" },
  quotaBox: { border: "1px solid", borderRadius: 10, padding: "10px 12px", margin: "0 0 16px" },
  soldOutBanner: { background: "#3a1515", color: "#FF8FA3", padding: "10px 20px", fontSize: 13, textAlign: "center", borderBottom: "1px solid #5a2020" },
  expiryBanner: { background: "#332a0f", color: "#FFB347", padding: "10px 20px", fontSize: 13, textAlign: "center", borderBottom: "1px solid #5a4a1a" },
  tabRow: { display: "flex", borderBottom: "1px solid #222", background: "#151515", overflowX: "auto" },
  tab: { flex: 1, background: "transparent", border: "none", color: "#666", padding: "14px 6px", fontSize: 12, cursor: "pointer", whiteSpace: "nowrap", display: "flex", alignItems: "center", justifyContent: "center", gap: 4 },
  tabActive: { flex: 1, background: "transparent", border: "none", color: "#4ECDC4", padding: "14px 6px", fontSize: 12, cursor: "pointer", borderBottom: "2px solid #4ECDC4", fontWeight: 700, whiteSpace: "nowrap", display: "flex", alignItems: "center", justifyContent: "center", gap: 4 },
  badge: { background: "#4ECDC4", color: "#000", borderRadius: 10, padding: "1px 7px", fontSize: 11, fontWeight: 700 },
  calContainer: { padding: 16 },
  container: { padding: 20 },
  weekNav: { display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 },
  navBtn: { background: "#2a2a2a", border: "none", color: "#fff", borderRadius: 8, padding: "8px 14px", cursor: "pointer", fontSize: 14 },
  weekLabel: { color: "#aaa", fontSize: 13 },
  gridHint: { color: "#666", fontSize: 11, marginBottom: 10 },
  calScroll: { overflowX: "auto", borderRadius: 12, border: "1px solid #222", maxHeight: "60vh", overflowY: "auto" },
  table: { width: "100%", borderCollapse: "collapse", minWidth: 600 },
  thTime: { width: 50, position: "sticky", left: 0, top: 0, zIndex: 4, background: "#1a1a1a" },
  th: { padding: "10px 4px", textAlign: "center", background: "#1a1a1a", borderBottom: "1px solid #222", position: "sticky", top: 0, zIndex: 2 },
  dayLabel: { color: "#888", fontSize: 11 },
  dateLabel: { color: "#fff", fontWeight: 700, fontSize: 16 },
  tdTime: { fontSize: 11, padding: "2px 6px", whiteSpace: "nowrap", position: "sticky", left: 0, zIndex: 1, background: "#0f0f0f" },
  td: { padding: "1px", borderLeft: "1px solid #161616" },
  slotMulti: { display: "flex", gap: 1, minHeight: 30 },
  slotChip: { flex: 1, borderRadius: 3, padding: "2px 3px", display: "flex", alignItems: "center", justifyContent: "space-between", minHeight: 30, minWidth: 0 },
  slotName: { fontSize: 10, fontWeight: 700, color: "#fff" },
  slotNameFull: { fontSize: 9, fontWeight: 700, color: "#fff", lineHeight: 1.1, wordBreak: "break-word", overflow: "hidden" },
  slotLabelBlock: { display: "flex", flexDirection: "column", gap: 1, minWidth: 0, overflow: "hidden" },
  slotTimeFull: { fontSize: 8, color: "#ffffffcc", lineHeight: 1.1, whiteSpace: "nowrap", overflow: "hidden" },
  slotTypeFull: { fontSize: 8, color: "#ffffffaa", lineHeight: 1.1, whiteSpace: "nowrap", overflow: "hidden" },
  slotStudentsFull: { fontSize: 8, color: "#FFE66D", fontWeight: 700, lineHeight: 1.1, wordBreak: "break-word", overflow: "hidden" },
  slotBottomTime: { fontSize: 8, color: "#ffffffaa", lineHeight: 1.1, marginTop: "auto", whiteSpace: "nowrap", overflow: "hidden" },
  cancelSlotBtn: { background: "transparent", border: "none", color: "#fff", cursor: "pointer", fontSize: 10, padding: 0 },
  slotEmpty: { width: "100%", minHeight: 20, background: "#1a1a1a", border: "none", borderRadius: 3, color: "#3a3a3a", fontSize: 13, cursor: "pointer" },
  slotEmptyRO: { minHeight: 20, background: "#1a1a1a", borderRadius: 3 },
  slotAdd: { width: 18, minHeight: 30, background: "#202020", border: "1px dashed #3a3a3a", borderRadius: 3, color: "#6BCB77", fontSize: 12, cursor: "pointer", flexShrink: 0 },
  slotPast: { minHeight: 20, background: "#111", borderRadius: 3 },
  slotDisabled: { minHeight: 20, background: "#141414", borderRadius: 3 },
  slotClosed: { minHeight: 20, background: "#0c0c0c", borderRadius: 3 },
  closedTag: { fontSize: 9, color: "#7a4040", marginTop: 2 },
  todayTag: { fontSize: 9, color: "#4ECDC4", marginTop: 2, fontWeight: 700 },
  kpiRow: { display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 12, marginBottom: 24 },
  kpiCard: { background: "#1a1a1a", borderRadius: 12, padding: "18px 12px", textAlign: "center" },
  kpiLabel: { color: "#888", fontSize: 12, marginBottom: 6 },
  kpiBig: { fontSize: 22, fontWeight: 800, color: "#4ECDC4" },
  sectionTitle: { fontSize: 16, fontWeight: 700, marginBottom: 12, color: "#aaa" },
  emptyText: { color: "#555", textAlign: "center", padding: "40px 0" },
  bookingList: { display: "flex", flexDirection: "column", gap: 10 },
  bookingItem: { background: "#1a1a1a", borderRadius: 12, padding: "14px 16px", display: "flex", alignItems: "center", gap: 12 },
  coachStatRow: { background: "#1a1a1a", borderRadius: 12, padding: "12px 14px", display: "flex", alignItems: "center", gap: 10 },
  dot: { width: 10, height: 10, borderRadius: "50%", flexShrink: 0 },
  avatar: { width: 38, height: 38, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 13, color: "#000", flexShrink: 0 },
  bookingCoach: { fontWeight: 600, fontSize: 14 },
  bookingTime: { color: "#666", fontSize: 12, marginTop: 2 },
  revenueNum: { fontWeight: 800, fontSize: 15, color: "#4ECDC4" },
  idTag: { fontSize: 11, color: "#555", marginLeft: 6 },
  plusTag: { fontSize: 12, color: "#6BCB77", marginLeft: 6 },
  soloTag: { fontSize: 10, color: "#4ECDC4", background: "#13302e", padding: "1px 6px", borderRadius: 6, marginLeft: 4 },
  charterTag: { fontSize: 10, color: "#111", background: "#fff", padding: "1px 6px", borderRadius: 6, marginLeft: 4, fontWeight: 700 },
  duoTag: { fontSize: 10, color: "#FFB347", background: "#33260f", padding: "1px 6px", borderRadius: 6, marginLeft: 4 },
  cancelledTag: { fontSize: 10, color: "#FF8FA3", background: "#3a1515", padding: "1px 6px", borderRadius: 6, marginLeft: 4 },
  cancelBtn: { background: "#2a2a2a", border: "none", color: "#FF6B6B", borderRadius: 8, padding: "6px 12px", fontSize: 13, cursor: "pointer" },
  pastTag: { color: "#555", fontSize: 12 },
  lockTag: { color: "#FFB347", fontSize: 12 },
  addBtn: { background: "#4ECDC4", color: "#000", border: "none", borderRadius: 8, padding: "8px 14px", fontSize: 13, fontWeight: 700, cursor: "pointer" },
  creditBtn: { background: "#1d3a2a", border: "none", color: "#6BCB77", borderRadius: 8, padding: "6px 10px", fontSize: 12, cursor: "pointer", whiteSpace: "nowrap" },
  smallBtn: { background: "#2a2a2a", border: "none", color: "#4ECDC4", borderRadius: 8, padding: "6px 10px", fontSize: 12, cursor: "pointer" },
  delBtn: { background: "#2a2a2a", border: "none", color: "#FF6B6B", borderRadius: 8, padding: "6px 10px", fontSize: 12, cursor: "pointer" },
  flexBetween: { display: "flex", alignItems: "center", justifyContent: "space-between" },
  formCard: { background: "#1a1a1a", borderRadius: 12, padding: 20, maxWidth: 360 },
  ledgerTotal: { marginTop: 16, padding: "12px 16px", background: "#1a1a1a", borderRadius: 10, fontWeight: 700, color: "#4ECDC4", textAlign: "right" },
  monthCard: { background: "#1a1a1a", borderRadius: 12, padding: "12px 16px" },
  monthHead: { fontWeight: 700, fontSize: 14, marginBottom: 8, color: "#fff" },
  monthRow: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "3px 0" },
  monthLabel: { fontSize: 13, color: "#999" },
  classNum: { fontWeight: 700, fontSize: 15, color: "#FFB347" },
  filterRow: { display: "flex", alignItems: "center", gap: 8, marginBottom: 14 },
  filterWrap: { display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 },
  recSummary: { background: "#151515", borderRadius: 8, padding: "8px 12px", fontSize: 13, color: "#4ECDC4", fontWeight: 700, marginBottom: 12 },
  donePill: { fontSize: 10, color: "#888", background: "#222", padding: "1px 6px", borderRadius: 6, marginLeft: 6 },
  recDetail: { marginTop: 8, padding: "8px 10px", background: "#141414", borderRadius: 8, fontSize: 12, color: "#aaa", lineHeight: 1.7 },
  lowWarnBox: { background: "#3a1515", color: "#FF8FA3", borderRadius: 10, padding: "10px 12px", fontSize: 12, marginTop: 12, lineHeight: 1.6 },
  purchaseBreakdown: { background: "#141414", borderRadius: 10, padding: "8px 12px", marginTop: 4, marginBottom: 4, display: "flex", flexDirection: "column", gap: 8 },
  purchaseRow: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", borderBottom: "1px solid #222", paddingBottom: 8 },
  lowPill: { fontSize: 10, color: "#fff", background: "#FF6B6B", padding: "1px 6px", borderRadius: 6, marginLeft: 6, fontWeight: 700 },
  filterLabel: { fontSize: 13, color: "#aaa" },
  select: { background: "#2a2a2a", border: "1px solid #333", borderRadius: 8, padding: "8px 12px", color: "#fff", fontSize: 13, outline: "none" },
  linkBtn: { background: "transparent", border: "none", color: "#888", fontSize: 12, cursor: "pointer", textDecoration: "underline", padding: "2px 0", marginTop: 2 },
  assistHint: { color: "#666", fontSize: 12, marginTop: 16 },
  amountPreview: { color: "#6BCB77", fontSize: 13, fontWeight: 600, margin: "4px 0 12px" },
  segRow: { display: "flex", gap: 8 },
  checkRow: { display: "flex", gap: 16, marginBottom: 16 },
  studentChipWrap: { display: "flex", gap: 6, flexWrap: "wrap" },
  studentChip: { background: "#2a2a2a", border: "1px solid #333", color: "#ccc", borderRadius: 16, padding: "6px 12px", fontSize: 12, cursor: "pointer" },
  studentChipActive: { background: "#4ECDC4", border: "1px solid #4ECDC4", color: "#000", borderRadius: 16, padding: "6px 12px", fontSize: 12, cursor: "pointer", fontWeight: 700 },
  studentChipDisabled: { background: "#1a1a1a", border: "1px solid #222", color: "#444", borderRadius: 16, padding: "6px 12px", fontSize: 12, cursor: "not-allowed" },
  rosterChip: { display: "inline-flex", alignItems: "center", gap: 6, background: "#2a2a2a", border: "1px solid #333", color: "#ccc", borderRadius: 16, padding: "6px 8px 6px 12px", fontSize: 12 },
  rosterRemoveBtn: { background: "transparent", border: "none", color: "#FF8FA3", cursor: "pointer", fontSize: 12, padding: "0 2px" },
  checkLabel: { display: "flex", alignItems: "center", gap: 6, fontSize: 13, color: "#ccc", cursor: "pointer" },
  seg: { flex: 1, background: "#2a2a2a", border: "1px solid #333", color: "#aaa", borderRadius: 10, padding: "10px", cursor: "pointer", fontSize: 14 },
  segActive: { flex: 1, background: "#4ECDC4", border: "1px solid #4ECDC4", color: "#000", borderRadius: 10, padding: "10px", cursor: "pointer", fontSize: 14, fontWeight: 700 },
  segDisabled: { flex: 1, background: "#1a1a1a", border: "1px solid #2a2a2a", color: "#444", borderRadius: 10, padding: "10px", cursor: "not-allowed", fontSize: 14 },
  priceBox: { background: "#222", borderRadius: 10, padding: "12px 14px", margin: "16px 0" },
  priceRow: { display: "flex", justifyContent: "space-between", fontSize: 13, color: "#aaa", padding: "3px 0" },
  modalOverlay: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.7)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100, padding: 20 },
  signCanvas: { width: "100%", height: 160, background: "#fff", borderRadius: 10, touchAction: "none", display: "block" },
  signRow: { display: "flex", gap: 6, flexWrap: "wrap", marginTop: 6 },
  signChip: { background: "#2a2a2a", border: "1px solid #333", color: "#FFB347", borderRadius: 12, padding: "4px 10px", fontSize: 11, cursor: "pointer" },
  signedChip: { background: "#13302e", border: "1px solid #1d3a2a", color: "#6BCB77", borderRadius: 12, padding: "4px 10px", fontSize: 11, cursor: "default" },
  modal: { background: "#1a1a1a", borderRadius: 16, padding: "28px 24px", width: 320, textAlign: "center", maxHeight: "85vh", overflowY: "auto" },
  modalTitle: { fontSize: 18, fontWeight: 700, marginBottom: 12 },
  modalText: { color: "#aaa", marginBottom: 20, lineHeight: 1.6 },
  modalBtns: { display: "flex", gap: 12, marginTop: 8 },
  modalCancel: { flex: 1, background: "#2a2a2a", border: "none", color: "#aaa", borderRadius: 10, padding: 12, cursor: "pointer", fontSize: 14 },
  modalConfirm: { flex: 1, background: "#4ECDC4", border: "none", color: "#000", borderRadius: 10, padding: 12, cursor: "pointer", fontSize: 14, fontWeight: 700 },
  toast: { position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)", color: "#000", fontWeight: 700, padding: "12px 24px", borderRadius: 30, fontSize: 14, zIndex: 200, boxShadow: "0 4px 20px rgba(0,0,0,0.3)", maxWidth: "90%", textAlign: "center" },
};

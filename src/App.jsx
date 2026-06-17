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
const ASSIST_CANCEL_LIMIT = 1; // 每位教練每月「24小時內由管理員代取消」額度

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

export default function App() {
  const [coaches, setCoaches] = useState(() => persisted("coaches", DEFAULT_COACHES));
  const [currentUser, setCurrentUser] = useState(null);
  const [adminPassword, setAdminPassword] = useState(() => persisted("adminPassword", "admin123"));
  const [view, setView] = useState("login");
  // bookings: key date_time(15min) -> { coachId, start, hours, type }  (type: 'solo' | 'duo')
  const [bookings, setBookings] = useState(() => persisted("bookings", {}));
  const [purchaseLog, setPurchaseLog] = useState(() => persisted("purchaseLog", []));
  const [ledgerFilter, setLedgerFilter] = useState("all"); // coachId or "all"
  const [editDateRec, setEditDateRec] = useState(null); // {id, date}
  const [weekOffset, setWeekOffset] = useState(0);
  const [bookModal, setBookModal] = useState(null);   // { date, time }
  const [charterModal, setCharterModal] = useState(null); // admin charter { date, time, hours }
  const [charterLog, setCharterLog] = useState(() => persisted("charterLog", [])); // {date, bookDate, start, hours, amount}
  const [assistCancelLog, setAssistCancelLog] = useState(() => persisted("assistCancelLog", [])); // {coachId, month, date, start}
  const [syncState, setSyncState] = useState(cloudEnabled ? "connecting" : "local"); // connecting | synced | local | error

  // 同步用：記住最後一次「已儲存／已收到」嘅內容，避免回音造成無限迴圈
  const lastSyncedRef = useRef(null);
  const readyRef = useRef(!cloudEnabled); // 雲端模式要等首次載入完成先準許寫入
  const saveTimer = useRef(null);

  const applyBundle = (d) => {
    if (!d) return;
    if (d.coaches !== undefined) setCoaches(d.coaches);
    if (d.adminPassword !== undefined) setAdminPassword(d.adminPassword);
    if (d.bookings !== undefined) setBookings(d.bookings);
    if (d.purchaseLog !== undefined) setPurchaseLog(d.purchaseLog);
    if (d.charterLog !== undefined) setCharterLog(d.charterLog);
    if (d.assistCancelLog !== undefined) setAssistCancelLog(d.assistCancelLog);
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
        const seed = { coaches, adminPassword, bookings, purchaseLog, charterLog, assistCancelLog };
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
    const bundle = { coaches, adminPassword, bookings, purchaseLog, charterLog, assistCancelLog };
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
  }, [coaches, adminPassword, bookings, purchaseLog, charterLog, assistCancelLog]);

  const [cancelModal, setCancelModal] = useState(null);
  const [adminCancelModal, setAdminCancelModal] = useState(null); // {date,start,coachId,type}
  const [delLedgerModal, setDelLedgerModal] = useState(null); // ledger record to delete
  const [toast, setToast] = useState(null);
  const [loginForm, setLoginForm] = useState({ id: "", password: "" });
  const [pwForm, setPwForm] = useState({ old: "", new1: "", new2: "" });
  const [editCoach, setEditCoach] = useState(null);
  const [addCreditModal, setAddCreditModal] = useState(null);
  const [adminTab, setAdminTab] = useState("overview");
  const [resetModal, setResetModal] = useState(false);

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
    const coach = coaches.find((c) => (c.username || "").toLowerCase() === uid && c.password === loginForm.password);
    if (coach) { setCurrentUser({ ...coach, role: "coach" }); setView("calendar"); }
    else showToast("帳號或密碼錯誤", "error");
  };
  const logout = () => { setCurrentUser(null); setView("login"); setLoginForm({ id: "", password: "" }); };

  const isCoach = currentUser?.role === "coach";
  const liveUser = isCoach ? getCoach(currentUser.id) : currentUser;
  const remaining = isCoach && liveUser ? liveUser.credits - liveUser.used : 0;
  const soldOut = isCoach && remaining <= 0;

  // bookings[key] is an ARRAY of entries; each entry occupies "seats" (charter=2, others=1)
  const cellArr = (date, slot) => bookings[`${date}_${slot}`] || [];
  const seats = (entry) => (entry.type === "charter" ? MAX_CONCURRENT : 1);
  const occupancy = (date, slot) => cellArr(date, slot).reduce((n, e) => n + seats(e), 0);

  // can we place a booking of `hours` at date/time?
  const canPlace = (date, time, hours, isCharter = false) => {
    const slots = slotsFor(time, hours);
    const need = isCharter ? MAX_CONCURRENT : 1;
    for (const s of slots) {
      const [hh] = s.split(":").map(Number);
      if (hh >= 22) return "超出營業時間";
      if (occupancy(date, s) + need > MAX_CONCURRENT)
        return isCharter ? "呢個時段唔夠空（包場需全場）" : "呢個時段已滿（最多2名教練）";
    }
    return null;
  };

  const openBook = (date, time) => {
    if (isClosedDay(date)) return showToast("星期四、五休息，不開放預約", "error");
    if (soldOut) return showToast("你已用晒購買堂數，請聯絡管理員增購", "error");
    setBookModal({ date, time, sessionType: "solo", hours: 1 });
  };

  const confirmBook = () => {
    const { date, time, sessionType, hours } = bookModal;
    const creditCost = hours; // 1hr = 1堂, 1.5hr = 1.5堂
    if (creditCost > remaining) { showToast("剩餘堂數不足", "error"); return; }
    const err = canPlace(date, time, hours);
    if (err) { showToast(err, "error"); return; }
    const price = sessionType === "duo" ? duoPrice(hours) : liveUser.rate * hours;
    const slots = slotsFor(time, hours);
    const entry = { coachId: currentUser.id, start: time, hours, type: sessionType, price };
    setBookings((prev) => {
      const u = { ...prev };
      slots.forEach((s) => { u[`${date}_${s}`] = [...(u[`${date}_${s}`] || []), entry]; });
      return u;
    });
    setCoaches((prev) => prev.map((c) => c.id === currentUser.id ? { ...c, used: c.used + creditCost } : c));
    showToast("預約成功！");
    setBookModal(null);
  };

  // ADMIN: place a charter/group booking (occupies whole venue), price editable
  const confirmCharter = () => {
    const { date, time, hours, charterType, price, coachName } = charterModal;
    if (isClosedDay(date)) { showToast("休息日", "error"); return; }
    const err = canPlace(date, time, hours, true);
    if (err) { showToast(err, "error"); return; }
    const amt = parseInt(price) || 0;
    const slots = slotsFor(time, hours);
    const entry = { coachId: 0, start: time, hours, type: "charter", charterType, price: amt, coachName: coachName || "" };
    setBookings((prev) => {
      const u = { ...prev };
      slots.forEach((s) => { u[`${date}_${s}`] = [...(u[`${date}_${s}`] || []), entry]; });
      return u;
    });
    setCharterLog((prev) => [{ date: new Date().toISOString().slice(0, 16).replace("T", " "), bookDate: date, start: time, hours, charterType, amount: amt, coachName: coachName || "" }, ...prev]);
    showToast(`已落${charterType === "group" ? "小組訓練" : "私人包場"}（$${amt}）`);
    setCharterModal(null);
  };

  const openCancel = (date, start, coachId, type) => {
    const hrs = hoursUntil(date, start);
    if (currentUser.role === "coach" && hrs < 24) return showToast("24小時內取消需要管理員協助", "error");
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
    if (meta.type !== "charter") {
      setCoaches((prev) => prev.map((c) => c.id === coachId ? { ...c, used: Math.max(0, c.used - meta.hours) } : c));
      // 由管理員協助、而且係 24 小時內嘅取消，計入該教練本月額度
      if (byAdmin && hoursUntil(date, start) < 24) {
        setAssistCancelLog((prev) => [{ coachId, month: monthKey(formatDate(new Date())), date, start }, ...prev]);
      }
    } else {
      setCharterLog((prev) => prev.filter((r) => !(r.bookDate === date && r.start === start)));
    }
    showToast(byAdmin ? "已協助取消" : "已取消預約");
    setCancelModal(null);
  };

  // 教練本月已用 / 剩餘代取消額度
  const assistUsedThisMonth = (coachId) => {
    const m = monthKey(formatDate(new Date()));
    return assistCancelLog.filter((r) => r.coachId === coachId && r.month === m).length;
  };

  const changePassword = () => {
    if (pwForm.new1 !== pwForm.new2) return showToast("兩次新密碼唔一致", "error");
    if (pwForm.new1.length < 4) return showToast("密碼至少4位", "error");
    if (currentUser.role === "admin") {
      if (pwForm.old !== adminPassword) return showToast("舊密碼錯誤", "error");
      setAdminPassword(pwForm.new1);
    } else {
      if (pwForm.old !== getCoach(currentUser.id).password) return showToast("舊密碼錯誤", "error");
      setCoaches((prev) => prev.map((c) => c.id === currentUser.id ? { ...c, password: pwForm.new1 } : c));
    }
    setPwForm({ old: "", new1: "", new2: "" });
    showToast("密碼已更新");
  };

  const addCredits = (coachId, qty) => {
    const coach = getCoach(coachId);
    const amount = qty * coach.rate;
    setCoaches((prev) => prev.map((c) => c.id === coachId ? { ...c, credits: c.credits + qty } : c));
    setPurchaseLog((prev) => [{ id: "p" + Date.now() + "-" + Math.random().toString(36).slice(2), date: new Date().toISOString().slice(0, 10), coachId, coachName: coach.name, qty, amount, rate: coach.rate }, ...prev]);
    showToast(`已為 ${coach.name} 增加 ${qty} 堂（$${amount}）`);
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
          類型: v.type === "charter" ? (v.charterType === "group" ? "小組訓練" : "私人包場") : v.type === "duo" ? "一對二" : "一對一",
          教練: v.type === "charter" ? (v.coachName || "") : (getCoach(v.coachId)?.name || ""),
          收費: v.price || 0,
        }); });
      });
      bkRows.sort((a, b) => `${a.日期}${a.開始}`.localeCompare(`${b.日期}${b.開始}`));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(bkRows.length ? bkRows : [{ 日期: "", 開始: "", 時長小時: "", 類型: "", 教練: "", 收費: "" }]), "上堂記錄");

      // 5) 包場/小組
      const chRows = charterLog.map((r) => ({ 落單時間: r.date, 預約日期: r.bookDate, 開始: r.start, 時長小時: r.hours, 類型: r.charterType === "group" ? "小組訓練" : "私人包場", 負責教練: r.coachName || "", 收費: r.amount }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(chRows.length ? chRows : [{ 落單時間: "", 預約日期: "", 開始: "", 時長小時: "", 類型: "", 負責教練: "", 收費: "" }]), "包場小組");

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
        myBookings.push({ date, start: v.start, hours: v.hours, type: v.type });
    });
  });
  myBookings.sort((a, b) => `${a.date}${a.start}`.localeCompare(`${b.date}${b.start}`));

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
          allBookings.push({ date, start: v.start, hours: v.hours, type: v.type, charterType: v.charterType, price: v.price || 0, coachName: v.coachName || "", coach: v.type === "charter" ? null : getCoach(v.coachId), coachId: v.coachId });
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

    // 各教練總付款（買堂 + 初始）
    const coachPaid = {};
    coaches.forEach((c) => { coachPaid[c.id] = purchaseLog.filter((r) => r.coachId === c.id).reduce((a, r) => a + r.amount, 0) + initialCreditsOf(c) * c.rate; });

    return (
      <div style={S.appBg}>
        <Header title="管理員" onLogout={logout} syncState={syncState} />
        <div style={S.tabRow}>
          {[["overview", "📊 總覽"], ["coaches", "👥 教練"], ["charter", "🏟️ 包場"], ["ledger", "💰 流水帳"], ["records", "📋 記錄"], ["settings", "⚙️ 設定"]].map(([k, label]) => (
            <button key={k} style={adminTab === k ? S.tabActive : S.tab} onClick={() => setAdminTab(k)}>{label}</button>
          ))}
        </div>

        {adminTab === "overview" && (
          <div style={S.container}>
            <div style={S.kpiRow}>
              <div style={S.kpiCard}><div style={S.kpiLabel}>總收入</div><div style={S.kpiBig}>${totalRevenue.toLocaleString()}</div></div>
              <div style={S.kpiCard}><div style={S.kpiLabel}>已上堂數</div><div style={S.kpiBig}>{totalUsed}</div></div>
              <div style={S.kpiCard}><div style={S.kpiLabel}>已售堂數</div><div style={S.kpiBig}>{totalSold}</div></div>
            </div>

            <h2 style={S.sectionTitle}>每月收入</h2>
            <div style={S.bookingList}>
              {allMonths.length === 0 ? <p style={S.emptyText}>暫無收入</p> : allMonths.map((m) => (
                <div key={m} style={S.monthCard}>
                  <div style={S.monthHead}>{m === "初始" ? "初始已售堂數" : m}</div>
                  <div style={S.monthRow}><span style={S.monthLabel}>一筆過租金（買堂）</span><span style={S.revenueNum}>${(purchaseByMonth[m] || 0).toLocaleString()}</span></div>
                  <div style={S.monthRow}><span style={S.monthLabel}>實際堂數收入</span><span style={S.classNum}>${(classByMonth[m] || 0).toLocaleString()}</span></div>
                </div>
              ))}
            </div>
            <p style={S.assistHint}>「一筆過租金」= 教練買堂時實收現金；「實際堂數收入」= 當月實際 book 咗嘅堂（一對一／一對二／包場）價值。</p>

            <h2 style={{ ...S.sectionTitle, marginTop: 24 }}>各教練統計</h2>
            <div style={S.bookingList}>
              {coaches.map((c) => (
                <div key={c.id} style={S.coachStatRow}>
                  <div style={{ ...S.avatar, background: c.color }}>{c.initials}</div>
                  <div style={{ flex: 1 }}>
                    <div style={S.bookingCoach}>{c.name}</div>
                    <div style={S.bookingTime}>每堂 ${c.rate}　已用 {c.used}/{c.credits} 堂</div>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <div style={S.revenueNum}>${(coachPaid[c.id] || 0).toLocaleString()}</div>
                    <div style={S.bookingTime}>總付款　剩 {c.credits - c.used} 堂</div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {adminTab === "coaches" && (
          <div style={S.container}>
            <div style={S.flexBetween}>
              <h2 style={S.sectionTitle}>教練帳戶</h2>
              <button style={S.addBtn} onClick={() => setEditCoach({ id: null, username: "", name: "", credits: 0, rate: 200, password: "1234" })}>+ 新增教練</button>
            </div>
            <div style={S.bookingList}>
              {coaches.map((c) => (
                <div key={c.id} style={S.coachStatRow}>
                  <div style={{ ...S.avatar, background: c.color }}>{c.initials}</div>
                  <div style={{ flex: 1 }}>
                    <div style={S.bookingCoach}>{c.name} <span style={S.idTag}>@{c.username}</span></div>
                    <div style={S.bookingTime}>堂數 {c.used}/{c.credits}　每堂 ${c.rate}　密碼 {c.password}</div>
                  </div>
                  <button style={S.creditBtn} onClick={() => setAddCreditModal({ coachId: c.id, qty: 1 })}>+ 堂</button>
                  <button style={S.smallBtn} onClick={() => setEditCoach(c)}>編輯</button>
                  <button style={S.delBtn} onClick={() => { setCoaches((prev) => prev.filter((x) => x.id !== c.id)); showToast("已刪除教練"); }}>刪</button>
                </div>
              ))}
            </div>
          </div>
        )}

        {adminTab === "charter" && (
          <div style={S.calContainer}>
            <h2 style={S.sectionTitle}>包場 Booking（管理員專用）</h2>
            <p style={S.gridHint}>${CHARTER_PRICE} 一節，時長自選；包場期間全場獨佔（其他教練不能同時 book）。按空格落單。</p>
            <div style={S.weekNav}>
              <button style={S.navBtn} onClick={() => setWeekOffset((w) => w - 1)}>‹ 上週</button>
              <span style={S.weekLabel}>{formatDate(days[0])} – {formatDate(days[6])}</span>
              <button style={S.navBtn} onClick={() => setWeekOffset((w) => w + 1)}>下週 ›</button>
            </div>
            <div style={S.calScroll}>
              <table style={S.table}>
                <thead><tr><th style={S.thTime}></th>
                  {days.map((d) => { const closed = CLOSED_DAYS.includes(d.getDay()); return <th key={d} style={S.th}><div style={{ ...S.dayLabel, color: closed ? "#5a3030" : undefined }}>{formatDay(d)}</div><div style={{ ...S.dateLabel, color: closed ? "#555" : undefined }}>{d.getDate()}</div>{closed && <div style={S.closedTag}>休息</div>}</th>; })}
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
                          const charter = here.find((v) => v.type === "charter");
                          const isPast = hoursUntil(date, time) < 0;
                          const closed = isClosedDay(date);
                          return (
                            <td key={date} style={{ ...S.td, borderTop: isHourStart ? "1px solid #2a2a2a" : "1px solid #161616", background: closed && here.length === 0 ? "#0c0c0c" : undefined }}>
                              {charter ? (
                                <div style={{ ...S.slotChip, background: "#ffffff22", borderLeft: "3px solid #fff", alignItems: "flex-start" }}>
                                  {charter.start === time && (
                                    <span style={S.slotNameFull}>{charter.charterType === "group" ? "小組" : "包場"}{charter.coachName ? ` · ${charter.coachName}` : ""}</span>
                                  )}
                                  {charter.start === time && <button style={S.cancelSlotBtn} onClick={() => doCancel(date, charter.start, 0, "charter", true)}>✕</button>}
                                </div>
                              ) : occ > 0 ? <div style={S.slotDisabled} />
                                : closed ? <div style={S.slotClosed} />
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
          </div>
        )}

        {adminTab === "ledger" && (() => {
          const filtered = ledgerFilter === "all" ? purchaseLog : purchaseLog.filter((r) => String(r.coachId) === String(ledgerFilter));
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
            {filtered.length > 0 && <div style={S.ledgerTotal}>{ledgerFilter === "all" ? "流水帳總額" : `${getCoach(ledgerFilter)?.name} 小計`}：${filteredTotal.toLocaleString()}</div>}
          </div>
          );
        })()}

        {adminTab === "records" && (
          <div style={S.container}>
            <h2 style={S.sectionTitle}>所有上堂記錄</h2>
            {allBookings.length === 0 ? <p style={S.emptyText}>暫無記錄</p> : (
              <div style={S.bookingList}>
                {allBookings.map(({ date, start, hours, type, charterType, price, coachName, coach, coachId }, i) => (
                  <div key={i} style={S.bookingItem}>
                    <div style={{ ...S.dot, background: type === "charter" ? "#fff" : coach?.color }} />
                    <div style={{ flex: 1 }}>
                      <div style={S.bookingCoach}>
                        {type === "charter" ? (charterType === "group" ? "小組訓練" : "私人包場") : coach?.name}{" "}
                        <span style={type === "duo" ? S.duoTag : type === "charter" ? S.charterTag : S.soloTag}>
                          {type === "duo" ? "1對2" : type === "charter" ? `$${price}` : "1對1"}
                        </span>
                      </div>
                      <div style={S.bookingTime}>{date} · {start}–{addMinutes(start, hours * 60)}（{hours}小時）{type === "charter" && coachName ? `　負責：${coachName}` : ""}</div>
                    </div>
                    <button style={S.delBtn} onClick={() => setAdminCancelModal({ date, start, coachId, type })}>取消</button>
                  </div>
                ))}
              </div>
            )}
            <p style={S.assistHint}>※ 管理員可協助取消任何時段（包括24小時內）</p>
          </div>
        )}

        {adminTab === "settings" && (
          <div style={S.container}>
            <h2 style={S.sectionTitle}>修改管理員密碼</h2>
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

            <h2 style={{ ...S.sectionTitle, marginTop: 28 }}>重設資料</h2>
            <div style={S.formCard}>
              <p style={{ ...S.bookingTime, marginBottom: 14, lineHeight: 1.6 }}>清除呢部裝置嘅所有資料，回復至初始狀態。建議先匯出備份。此動作無法復原。</p>
              <button style={{ ...S.loginBtn, background: "#FF6B6B", color: "#fff" }} onClick={() => setResetModal(true)}>🗑️ 重設所有資料</button>
            </div>
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
            <p style={S.amountPreview}>金額：${((getCoach(addCreditModal.coachId)?.rate || 0) * addCreditModal.qty).toLocaleString()}</p>
            <div style={S.modalBtns}>
              <button style={S.modalCancel} onClick={() => setAddCreditModal(null)}>取消</button>
              <button style={S.modalConfirm} onClick={() => { addCredits(addCreditModal.coachId, addCreditModal.qty); setAddCreditModal(null); }}>確認增加</button>
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
            <h3 style={{ ...S.modalTitle, textAlign: "center" }}>包場 / 小組 Booking</h3>
            <p style={{ ...S.modalText, textAlign: "center" }}>{charterModal.date}　{charterModal.time}</p>

            <label style={S.label}>類型</label>
            <div style={S.segRow}>
              <button style={charterModal.charterType === "private" ? S.segActive : S.seg} onClick={() => setCharterModal({ ...charterModal, charterType: "private" })}>私人包場</button>
              <button style={charterModal.charterType === "group" ? S.segActive : S.seg} onClick={() => setCharterModal({ ...charterModal, charterType: "group" })}>小組訓練</button>
            </div>

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

            <label style={{ ...S.label, marginTop: 14 }}>收費 ($，可自由修改)</label>
            <input style={S.input} type="number" min="0" value={charterModal.price}
              onChange={(e) => setCharterModal({ ...charterModal, price: e.target.value })} />

            <div style={S.priceBox}>
              <div style={S.priceRow}><span>時段</span><span>{charterModal.time} – {addMinutes(charterModal.time, charterModal.hours * 60)}</span></div>
              <div style={S.priceRow}><span>場地</span><span>全場獨佔</span></div>
              <div style={{ ...S.priceRow, color: "#4ECDC4", fontWeight: 700, fontSize: 16 }}><span>收費</span><span>${parseInt(charterModal.price) || 0}</span></div>
            </div>
            <div style={S.modalBtns}>
              <button style={S.modalCancel} onClick={() => setCharterModal(null)}>返回</button>
              <button style={S.modalConfirm} onClick={confirmCharter}>確認落單</button>
            </div>
          </div></div>
        )}

        {adminCancelModal && (() => {
          const within24 = adminCancelModal.type !== "charter" && hoursUntil(adminCancelModal.date, adminCancelModal.start) < 24;
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
                  {getCoach(adminCancelModal.coachId)?.name}　本月已用 {used} / {ASSIST_CANCEL_LIMIT} 次
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

        {resetModal && (
          <div style={S.modalOverlay}><div style={S.modal}>
            <h3 style={S.modalTitle}>重設所有資料</h3>
            <p style={S.modalText}>確定清除呢部裝置嘅所有資料（教練、預約、流水帳等），回復至初始狀態？<br /><br />此動作無法復原，建議先匯出備份。</p>
            <div style={S.modalBtns}>
              <button style={S.modalCancel} onClick={() => setResetModal(false)}>返回</button>
              <button style={{ ...S.modalConfirm, background: "#FF6B6B" }} onClick={() => {
                try { localStorage.removeItem(LS_KEY); } catch (e) { /* ignore */ }
                setCoaches(DEFAULT_COACHES); setAdminPassword("admin123"); setBookings({});
                setPurchaseLog([]); setCharterLog([]); setAssistCancelLog([]);
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
            <span>本月 24 小時內代取消額度</span>
            <span style={{ color: left > 0 ? "#4ECDC4" : "#FF6B6B", fontWeight: 700 }}>剩 {left} / {ASSIST_CANCEL_LIMIT} 次</span>
          </div>
        );
      })()}
      {soldOut && <div style={S.soldOutBanner}>⚠️ 你已用晒購買堂數，請聯絡管理員增購後再預約</div>}
      <div style={S.tabRow}>
        <button style={view === "calendar" ? S.tabActive : S.tab} onClick={() => setView("calendar")}>📅 預約場地</button>
        <button style={view === "myBookings" ? S.tabActive : S.tab} onClick={() => setView("myBookings")}>📋 我的預約 {myBookings.length > 0 && <span style={S.badge}>{myBookings.length}</span>}</button>
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
                        const charter = here.find((v) => v.type === "charter");
                        const isPast = hoursUntil(date, time) < 0;
                        const closed = isClosedDay(date);
                        const full = occ >= MAX_CONCURRENT;
                        const iAmHere = here.some((v) => v.coachId === currentUser.id);
                        const canAddHere = !charter && occ < MAX_CONCURRENT && !iAmHere && !isPast && !soldOut && !closed;
                        return (
                          <td key={date} style={{ ...S.td, borderTop: isHourStart ? "1px solid #2a2a2a" : "1px solid #161616", background: closed && here.length === 0 ? "#0c0c0c" : undefined }}>
                            {charter ? (
                              <div style={{ ...S.slotChip, background: "#ffffff22", borderLeft: "3px solid #fff", alignItems: "flex-start" }}>
                                {charter.start === time && <span style={S.slotNameFull}>{charter.charterType === "group" ? "小組" : "包場"}{charter.coachName ? ` · ${charter.coachName}` : ""}</span>}
                              </div>
                            ) : here.length > 0 ? (
                              <div style={S.slotMulti}>
                                {here.map((v, idx) => {
                                  const c = getCoach(v.coachId);
                                  const showLabel = v.start === time;
                                  return (
                                    <div key={idx} style={{ ...S.slotChip, background: c?.color + "33", borderLeft: `3px solid ${c?.color}` }}>
                                      {showLabel && <span style={S.slotNameFull}>{c?.name}{v.type === "duo" ? " ²" : ""}</span>}
                                      {showLabel && v.coachId === currentUser.id && hoursUntil(date, v.start) >= 24 && !isPast &&
                                        <button style={S.cancelSlotBtn} onClick={() => openCancel(date, v.start, v.coachId, v.type)}>✕</button>}
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
              {myBookings.map(({ date, start, hours, type }, i) => {
                const hrs = hoursUntil(date, start);
                const isPast = hrs < 0;
                const locked = hrs >= 0 && hrs < 24;
                return (
                  <div key={i} style={S.bookingItem}>
                    <div style={{ ...S.dot, background: liveUser.color }} />
                    <div style={{ flex: 1 }}>
                      <div style={S.bookingCoach}>{date} <span style={type === "duo" ? S.duoTag : S.soloTag}>{type === "duo" ? "1對2" : "1對1"}</span></div>
                      <div style={S.bookingTime}>{start} – {addMinutes(start, hours * 60)}（{hours}小時）</div>
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
        return (
          <div style={S.modalOverlay}><div style={{ ...S.modal, textAlign: "left" }}>
            <h3 style={{ ...S.modalTitle, textAlign: "center" }}>預約場地</h3>
            <p style={{ ...S.modalText, textAlign: "center" }}>{bookModal.date}　{bookModal.time}</p>
            <label style={S.label}>類型</label>
            <div style={S.segRow}>
              <button style={!isDuo ? S.segActive : S.seg} onClick={() => setBookModal({ ...bookModal, sessionType: "solo" })}>1對1</button>
              <button style={isDuo ? S.segActive : S.seg} onClick={() => setBookModal({ ...bookModal, sessionType: "duo" })}>1對2</button>
            </div>
            <label style={{ ...S.label, marginTop: 14 }}>時長</label>
            <div style={S.segRow}>
              <button style={bookModal.hours === 1 ? S.segActive : S.seg} onClick={() => setBookModal({ ...bookModal, hours: 1 })}>1 小時</button>
              <button style={bookModal.hours === 1.5 ? S.segActive : S.seg} onClick={() => setBookModal({ ...bookModal, hours: 1.5 })}>1.5 小時</button>
            </div>
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
      {toast && <Toast toast={toast} />}
    </div>
  );
}

function EditCoachModal({ coach, onClose, onSave }) {
  const [form, setForm] = useState({ id: coach.id, username: coach.username || "", name: coach.name, credits: coach.credits, rate: coach.rate, password: coach.password });
  return (
    <div style={S.modalOverlay}><div style={{ ...S.modal, width: 320, textAlign: "left" }}>
      <h3 style={S.modalTitle}>{coach.id ? "編輯教練" : "新增教練"}</h3>
      <Field label="教練全名（顯示用）"><input style={S.input} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>
      <Field label="登入帳號名稱"><input style={S.input} value={form.username} placeholder="例如 alex（登入用，不分大小寫）" onChange={(e) => setForm({ ...form, username: e.target.value })} /></Field>
      <Field label={coach.id ? "總購買堂數" : "初始購買堂數"}><input style={S.input} type="number" value={form.credits} onChange={(e) => setForm({ ...form, credits: parseInt(e.target.value) || 0 })} /></Field>
      <Field label="一對一每堂租金 ($)"><input style={S.input} type="number" value={form.rate} onChange={(e) => setForm({ ...form, rate: parseInt(e.target.value) || 0 })} /></Field>
      <Field label="密碼"><input style={S.input} value={form.password} onChange={(e) => setForm({ ...form, password: e.target.value })} /></Field>
      {!coach.id && form.credits > 0 && <p style={S.amountPreview}>初始堂數記入流水帳：${(form.credits * form.rate).toLocaleString()}</p>}
      <div style={S.modalBtns}>
        <button style={S.modalCancel} onClick={onClose}>取消</button>
        <button style={S.modalConfirm} onClick={() => { if (!form.name.trim()) return; onSave(form); }}>儲存</button>
      </div>
    </div></div>
  );
}

function Field({ label, children }) { return <div style={S.inputGroup}><label style={S.label}>{label}</label>{children}</div>; }
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
  thTime: { width: 50 },
  th: { padding: "10px 4px", textAlign: "center", background: "#1a1a1a", borderBottom: "1px solid #222", position: "sticky", top: 0 },
  dayLabel: { color: "#888", fontSize: 11 },
  dateLabel: { color: "#fff", fontWeight: 700, fontSize: 16 },
  tdTime: { fontSize: 11, padding: "2px 6px", whiteSpace: "nowrap" },
  td: { padding: "1px", borderLeft: "1px solid #161616" },
  slotMulti: { display: "flex", gap: 1, minHeight: 30 },
  slotChip: { flex: 1, borderRadius: 3, padding: "2px 3px", display: "flex", alignItems: "center", justifyContent: "space-between", minHeight: 30, minWidth: 0 },
  slotName: { fontSize: 10, fontWeight: 700, color: "#fff" },
  slotNameFull: { fontSize: 9, fontWeight: 700, color: "#fff", lineHeight: 1.1, wordBreak: "break-word", overflow: "hidden" },
  cancelSlotBtn: { background: "transparent", border: "none", color: "#fff", cursor: "pointer", fontSize: 10, padding: 0 },
  slotEmpty: { width: "100%", minHeight: 20, background: "#1a1a1a", border: "none", borderRadius: 3, color: "#3a3a3a", fontSize: 13, cursor: "pointer" },
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
  filterLabel: { fontSize: 13, color: "#aaa" },
  select: { background: "#2a2a2a", border: "1px solid #333", borderRadius: 8, padding: "8px 12px", color: "#fff", fontSize: 13, outline: "none" },
  linkBtn: { background: "transparent", border: "none", color: "#888", fontSize: 12, cursor: "pointer", textDecoration: "underline", padding: "2px 0", marginTop: 2 },
  assistHint: { color: "#666", fontSize: 12, marginTop: 16 },
  amountPreview: { color: "#6BCB77", fontSize: 13, fontWeight: 600, margin: "4px 0 12px" },
  segRow: { display: "flex", gap: 8 },
  seg: { flex: 1, background: "#2a2a2a", border: "1px solid #333", color: "#aaa", borderRadius: 10, padding: "10px", cursor: "pointer", fontSize: 14 },
  segActive: { flex: 1, background: "#4ECDC4", border: "1px solid #4ECDC4", color: "#000", borderRadius: 10, padding: "10px", cursor: "pointer", fontSize: 14, fontWeight: 700 },
  priceBox: { background: "#222", borderRadius: 10, padding: "12px 14px", margin: "16px 0" },
  priceRow: { display: "flex", justifyContent: "space-between", fontSize: 13, color: "#aaa", padding: "3px 0" },
  modalOverlay: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.7)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100, padding: 20 },
  modal: { background: "#1a1a1a", borderRadius: 16, padding: "28px 24px", width: 320, textAlign: "center", maxHeight: "85vh", overflowY: "auto" },
  modalTitle: { fontSize: 18, fontWeight: 700, marginBottom: 12 },
  modalText: { color: "#aaa", marginBottom: 20, lineHeight: 1.6 },
  modalBtns: { display: "flex", gap: 12, marginTop: 8 },
  modalCancel: { flex: 1, background: "#2a2a2a", border: "none", color: "#aaa", borderRadius: 10, padding: 12, cursor: "pointer", fontSize: 14 },
  modalConfirm: { flex: 1, background: "#4ECDC4", border: "none", color: "#000", borderRadius: 10, padding: 12, cursor: "pointer", fontSize: 14, fontWeight: 700 },
  toast: { position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)", color: "#000", fontWeight: 700, padding: "12px 24px", borderRadius: 30, fontSize: 14, zIndex: 200, boxShadow: "0 4px 20px rgba(0,0,0,0.3)", maxWidth: "90%", textAlign: "center" },
};

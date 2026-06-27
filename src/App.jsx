import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { cloudEnabled, cloudLoad, cloudSave, cloudSubscribe, SUPABASE_URL } from "./supabase.js";
import { PDFDocument, rgb, StandardFonts } from "pdf-lib";
import {
  LOGO, STAMP_PNG, DEFAULT_COACHES, DEFAULT_SUBADMINS, ADMIN_TAB_KEYS,
  COLORS, MAX_CONCURRENT, DUO_BASE, DUO_HALF_HOUR_ADD, CHARTER_PRICE,
  ASSIST_CANCEL_LIMIT, LOW_CREDIT_THRESHOLD, CLOSED_DAYS, TIME_SLOTS, LS_KEY,
} from "./constants.js";
import {
  persisted, loadSession, saveSession, clearSession, loadCalScale, saveCalScale,
  stableStringify, initialSession, duoPrice, isWholeVenue, rentalShort, rentalFull,
  isClosedDay, getDaysOfWeek, formatDate, isTodayDate, formatDay, monthKey,
  hoursUntil, addMinutes, slotsFor, slotIndex, buildEntryLines, addDaysToDate,
} from "./helpers.js";
import { S } from "./styles.js";
import { EditCoachModal, Field, SignaturePad, Header, Toast } from "./components.jsx";

export default function App() {
  const [coaches, setCoaches] = useState(() => persisted("coaches", DEFAULT_COACHES));
  const [currentUser, setCurrentUser] = useState(() => initialSession?.user || null);
  const [adminPassword, setAdminPassword] = useState(() => persisted("adminPassword", "admin123"));
  const [whatsappNumber, setWhatsappNumber] = useState(() => persisted("whatsappNumber", ""));
  const [venueNotice, setVenueNotice] = useState(() => persisted("venueNotice", ""));
  const [suggestionBox, setSuggestionBox] = useState(() => persisted("suggestionBox", []));
  const [subAdmins, setSubAdmins] = useState(() => persisted("subAdmins", DEFAULT_SUBADMINS));
  const [view, setView] = useState(() => initialSession?.view || "login");
  // bookings: key date_time(15min) -> { coachId, start, hours, type }  (type: 'solo' | 'duo')
  const [bookings, setBookings] = useState(() => persisted("bookings", {}));
  const [purchaseLog, setPurchaseLog] = useState(() => persisted("purchaseLog", []));
  const [invoiceCounter, setInvoiceCounter] = useState(() => persisted("invoiceCounter", 1));
  const [studentPurchaseLog, setStudentPurchaseLog] = useState(() => persisted("studentPurchaseLog", []));
  const [ledgerFilter, setLedgerFilter] = useState("all"); // coachId or "all"
  const [ledgerMonth, setLedgerMonth] = useState("all"); // "all" or "YYYY-MM"
  const [editDateRec, setEditDateRec] = useState(null); // {id, date}
  const [weekOffset, setWeekOffset] = useState(0);
  const [calScale, setCalScale] = useState(() => loadCalScale());
  const [myBookingsView, setMyBookingsView] = useState("list"); // list | calendar
  const updateCalScale = (v) => { setCalScale(v); saveCalScale(v); };
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
    if (d.whatsappNumber !== undefined) setWhatsappNumber(d.whatsappNumber);
    if (d.venueNotice !== undefined) setVenueNotice(d.venueNotice);
    if (d.suggestionBox !== undefined) setSuggestionBox(d.suggestionBox);
    if (d.invoiceCounter !== undefined) setInvoiceCounter(d.invoiceCounter);
    if (d.studentPurchaseLog !== undefined) setStudentPurchaseLog(d.studentPurchaseLog);
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
        const seed = { coaches, adminPassword, whatsappNumber, venueNotice, suggestionBox, invoiceCounter, subAdmins, bookings, purchaseLog, studentPurchaseLog, charterLog, assistCancelLog, cancelLog };
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
    const bundle = { coaches, adminPassword, whatsappNumber, venueNotice, suggestionBox, invoiceCounter, subAdmins, bookings, purchaseLog, studentPurchaseLog, charterLog, assistCancelLog, cancelLog };
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
  }, [coaches, adminPassword, whatsappNumber, venueNotice, suggestionBox, invoiceCounter, subAdmins, bookings, purchaseLog, studentPurchaseLog, charterLog, assistCancelLog, cancelLog]);

  const [cancelModal, setCancelModal] = useState(null);
  const [signModal, setSignModal] = useState(null); // {date,start,coachId,type,studentName}
  const [adminCancelModal, setAdminCancelModal] = useState(null); // {date,start,coachId,type}
  const [delLedgerModal, setDelLedgerModal] = useState(null); // ledger record to delete
  const [toast, setToast] = useState(null);
  const [loginForm, setLoginForm] = useState({ id: "", password: "" });
  const [pwForm, setPwForm] = useState({ old: "", new1: "", new2: "" });
  const [editCoach, setEditCoach] = useState(null);
  const [addCreditModal, setAddCreditModal] = useState(null);
  const [adminTab, setAdminTab] = useState(() => initialSession?.adminTab || "overview");
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
  const [suggestionText, setSuggestionText] = useState("");
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
      if (loginForm.password === adminPassword) { setCurrentUser({ id: 0, name: "管理員", role: "admin" }); setView("admin"); saveSession({ role: "admin" }); }
      else showToast("管理員密碼錯誤", "error");
      return;
    }
    const sub = subAdmins.find((s) => (s.username || "").toLowerCase() === uid && s.password === loginForm.password);
    if (sub) {
      setCurrentUser({ ...sub, role: "subadmin" });
      setView("admin");
      const firstAllowed = ["overview", "schedule", "coaches", "ledger", "records", "settings"].find((k) => sub.permissions?.[k]);
      setAdminTab(firstAllowed || "settings");
      saveSession({ role: "subadmin", id: sub.id });
      return;
    }
    const coach = coaches.find((c) => (c.username || "").toLowerCase() === uid && c.password === loginForm.password);
    if (coach) { setCurrentUser({ ...coach, role: "coach" }); setView("calendar"); saveSession({ role: "coach", id: coach.id }); }
    else showToast("帳號或密碼錯誤", "error");
  };
  const logout = () => { setCurrentUser(null); setView("login"); setLoginForm({ id: "", password: "" }); clearSession(); };

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
    const repeatWeeks = Math.max(1, bookModal.repeatWeeks || 1);
    const selected = Array.isArray(bookModal.students) ? bookModal.students : [];
    const extra = (bookModal.studentOther || "").trim();
    const studentList = [...selected, ...(extra ? [extra] : [])].filter(Boolean).slice(0, 4);
    const studentCharges = {};
    studentList.forEach((n) => { const s = myRoster.find((x) => x.name === n); studentCharges[n] = s ? (s.rate || 0) : 0; });
    const price = sessionType === "duo" ? duoPrice(hours) : liveUser.rate * hours;
    const rentalCost = liveUser.rate * hours; // 租場費用：用「落單嗰刻」嘅租金snapshot，日後改租金唔會影響舊紀錄

    let usedRemaining = remaining;
    const newBookingsBySlot = {}; // `${date}_${slot}` -> entry to append
    let okCount = 0, skippedDates = [];
    for (let w = 0; w < repeatWeeks; w++) {
      const wDate = w === 0 ? date : addDaysToDate(date, w * 7);
      if (creditCost > usedRemaining) { skippedDates.push(`${wDate}（堂數不足）`); continue; }
      const err = canPlace(wDate, time, hours);
      if (err) { skippedDates.push(`${wDate}（${err}）`); continue; }
      const entry = { coachId: currentUser.id, start: time, hours, type: sessionType, price, rentalCost, students: studentList, studentCharges, createdAt: new Date().toISOString().slice(0, 16).replace("T", " ") };
      slotsFor(time, hours).forEach((s) => {
        const key = `${wDate}_${s}`;
        newBookingsBySlot[key] = [...(newBookingsBySlot[key] || []), entry];
      });
      usedRemaining -= creditCost;
      okCount++;
    }

    if (okCount === 0) { showToast(skippedDates[0] || "預約失敗", "error"); return; }

    setBookings((prev) => {
      const u = { ...prev };
      Object.entries(newBookingsBySlot).forEach(([key, arr]) => { u[key] = [...(u[key] || []), ...arr]; });
      return u;
    });
    setCoaches((prev) => prev.map((c) => c.id === currentUser.id ? { ...c, used: c.used + creditCost * okCount } : c));
    if (repeatWeeks > 1) {
      showToast(skippedDates.length === 0 ? `已成功預約 ${okCount} 週` : `已預約 ${okCount} 週，跳過 ${skippedDates.length} 週：${skippedDates.join("、")}`);
    } else {
      showToast("預約成功！");
    }
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

  // 撳「攞 QR Code」自動開 WhatsApp，傳訊息去 admin 設定嗰個號碼，連埋呢個 booking 嘅時段
  const openWhatsAppQR = (date, start, hours, coachName) => {
    if (!whatsappNumber) { showToast("管理員未設定 WhatsApp 號碼，請聯絡管理員", "error"); return; }
    const msg = encodeURIComponent(`你好，我想攞場地 QR Code 入場。\n預約時段：\n${coachName}\n${date}\n${start}–${addMinutes(start, hours * 60)}`);
    window.open(`https://wa.me/${whatsappNumber}?text=${msg}`, "_blank");
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
  const normStudent = (s) => (typeof s === "string" ? { name: s, rate: 0, credits: 0, used: 0, phone: "" } : { rate: 0, credits: 0, used: 0, phone: "", ...s });
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
    setStudentPurchaseLog((prev) => [{ id: "sp" + Date.now() + "-" + Math.random().toString(36).slice(2), coachId: currentUser.id, studentName: name, date: new Date().toISOString().slice(0, 10), qty, rate: s.rate || 0, amount: (s.rate || 0) * qty }, ...prev]);
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

  // 匿名改善建議：完全唔存任何身份資訊（冇 coachId、冇帳號），淨係文字＋粗略日期
  // 日曆同步：每位教練一個獨一無二嘅 token，用嚎驗證 Edge Function 嗰個訂閱連結
  const genToken = () => (crypto.randomUUID ? crypto.randomUUID() : "t" + Date.now() + Math.random().toString(36).slice(2));
  const ensureCalendarToken = () => {
    if (liveUser.calendarToken) return liveUser.calendarToken;
    const t = genToken();
    setCoaches((prev) => prev.map((c) => c.id === currentUser.id ? { ...c, calendarToken: t } : c));
    return t;
  };
  const regenerateCalendarToken = () => {
    const t = genToken();
    setCoaches((prev) => prev.map((c) => c.id === currentUser.id ? { ...c, calendarToken: t } : c));
    showToast("已重新生成連結，舊連結會失效");
  };
  const calendarFeedUrl = liveUser?.calendarToken ? `${SUPABASE_URL}/functions/v1/coach-calendar?token=${liveUser.calendarToken}` : "";

  const submitSuggestion = (text) => {
    const t = text.trim();
    if (!t) return;
    setSuggestionBox((prev) => [{ id: "sg" + Date.now() + "-" + Math.random().toString(36).slice(2), text: t, date: formatDate(new Date()), read: false }, ...prev]);
    showToast("已匿名提交，多謝你嘅意見 🙏");
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
  const fmtMoney = (n) => n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  // 生成購買堂數 Invoice（PDF），跟住 Gymily 提供嗰張範本排版；個性化資料（Bill to / Contact Person）留空，
  // 其他全部跟住教練同呢筆購買記錄自動帶入；公司印章自動貼上。
  const generateInvoicePDF = async (record) => {
    try {
      const teal = rgb(0.067, 0.443, 0.478);
      const lightBlue = rgb(0.85, 0.91, 0.96);
      const grey = rgb(0.6, 0.6, 0.6);
      const black = rgb(0.1, 0.1, 0.1);

      const pdfDoc = await PDFDocument.create();
      const page = pdfDoc.addPage([595.28, 841.89]); // A4
      const { width, height } = page.getSize();
      const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
      const bold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
      const italic = await pdfDoc.embedFont(StandardFonts.HelveticaOblique);

      const marginX = 42;
      const rightX = width - marginX;

      // ---- 頂部 teal header ----
      const headerH = 105;
      page.drawRectangle({ x: 0, y: height - headerH, width, height: headerH, color: teal });
      page.drawText("Gymily Studio Limited", { x: rightX - bold.widthOfTextAtSize("Gymily Studio Limited", 13), y: height - 35, size: 13, font: bold, color: rgb(1, 1, 1) });
      const addrLines = ["709,29-35,Sha Tsui Road ,Technology Plaza,Tsuen", "Wan, N.T"];
      addrLines.forEach((line, i) => {
        page.drawText(line, { x: rightX - font.widthOfTextAtSize(line, 9), y: height - 50 - i * 12, size: 9, font, color: rgb(1, 1, 1) });
      });
      page.drawText("Invoice", { x: marginX, y: height - headerH + 22, size: 26, font: bold, color: rgb(1, 1, 1) });

      // ---- Bill to / Contact Person（個性化資料，留空） ----
      let y = height - headerH - 35;
      page.drawText("Bill to :", { x: marginX, y, size: 9, font: bold, color: black });
      page.drawText("Contact Person:", { x: marginX + 180, y, size: 9, font: bold, color: black });

      // ---- 右側：根據教練／購買記錄自動帶入 ----
      const invoiceNo = `GM${String(new Date().getFullYear()).slice(2)}${String(invoiceCounter).padStart(4, "0")}`;
      const infoRows = [
        ["Client number :", `Coach-${record.coachName}`],
        ["Invoice number :", invoiceNo],
        ["Invoice date :", record.date],
        ["Payment terms :", "Lesson"],
      ];
      infoRows.forEach(([label, val], i) => {
        const ly = y - i * 16;
        const labelW = bold.widthOfTextAtSize(label, 9);
        page.drawText(label, { x: rightX - labelW - font.widthOfTextAtSize(val, 9) - 8, y: ly, size: 9, font: bold, color: black });
        page.drawText(val, { x: rightX - font.widthOfTextAtSize(val, 9), y: ly, size: 9, font, color: black });
      });

      // ---- 分隔線 + 表頭 ----
      y -= 95;
      page.drawLine({ start: { x: marginX, y }, end: { x: rightX, y }, thickness: 1.4, color: grey });
      y -= 22;
      const colItemNo = marginX, colDesc = marginX + 60, colQty = marginX + 300, colUnit = marginX + 360, colAmt = marginX + 440;
      [["Item no.", colItemNo], ["DESCRIPTION", colDesc], ["Quantity", colQty], ["Unit Price", colUnit], ["AMOUNT", colAmt]].forEach(([txt, x]) => {
        page.drawText(txt, { x, y, size: 9, font: bold, color: black });
      });
      page.drawText("(HKD)", { x: colUnit, y: y - 12, size: 8, font: bold, color: black });
      page.drawText("(HKD)", { x: colAmt, y: y - 12, size: 8, font: bold, color: black });

      // ---- 資料行 ----
      y -= 55;
      page.drawText("Purchase Lesson from coach", { x: colDesc, y, size: 9.5, font, color: black });
      page.drawText(String(record.qty), { x: colQty, y, size: 9.5, font, color: black });
      page.drawText(`$${fmtMoney(record.rate)}`, { x: colUnit, y, size: 9.5, font, color: black });
      page.drawText(`$${fmtMoney(record.amount)}`, { x: colAmt, y, size: 9.5, font, color: black });

      // ---- 總額 ----
      y -= 70;
      page.drawLine({ start: { x: marginX, y: y + 25 }, end: { x: rightX, y: y + 25 }, thickness: 1, color: grey });
      const totalLabel = "Total Amount =";
      page.drawText(totalLabel, { x: colUnit - 10, y, size: 11, font: bold, color: black });
      page.drawText(`$${fmtMoney(record.amount)}`, { x: colAmt, y, size: 11, font: bold, color: black });

      // ---- Amount Paid（已收現金，全數）----
      y -= 35;
      page.drawLine({ start: { x: marginX, y: y + 25 }, end: { x: rightX, y: y + 25 }, thickness: 1, color: grey });
      page.drawText("Amount Paid", { x: colUnit - 10, y, size: 9.5, font: bold, color: black });
      page.drawText(`$${fmtMoney(record.amount)}`, { x: colAmt, y, size: 9.5, font, color: black });

      // ---- Balance Due（已全數收齊，$0）----
      y -= 35;
      page.drawRectangle({ x: marginX, y: y - 8, width: rightX - marginX, height: 30, color: lightBlue });
      page.drawText("Balance Due", { x: colUnit - 10, y, size: 14, font: bold, color: black });
      page.drawText("$0.00", { x: colAmt, y, size: 16, font: bold, color: black });

      // ---- 簽名格 + 公司印章（照用）----
      y -= 70;
      page.drawRectangle({ x: marginX, y, width: 180, height: 55, borderColor: grey, borderWidth: 1 });
      const stampBytes = await fetch(STAMP_PNG).then((r) => r.arrayBuffer());
      const stampImg = await pdfDoc.embedPng(stampBytes);
      const stampW = 90, stampH = stampImg.height * (stampW / stampImg.width);
      page.drawImage(stampImg, { x: marginX + 240, y: y - 5, width: stampW, height: stampH });
      page.drawLine({ start: { x: marginX + 230, y: y - 12 }, end: { x: marginX + 230 + stampW + 20, y: y - 12 }, thickness: 1, color: black });

      // ---- 底部 teal footer ----
      page.drawRectangle({ x: 0, y: 0, width, height: 26, color: teal });
      const footerTxt = "Thank you for your business";
      page.drawText(footerTxt, { x: (width - italic.widthOfTextAtSize(footerTxt, 10)) / 2, y: 9, size: 10, font: italic, color: rgb(1, 1, 1) });

      const pdfBytes = await pdfDoc.save();
      const blob = new Blob([pdfBytes], { type: "application/pdf" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = `${invoiceNo}_${record.coachName}.pdf`;
      document.body.appendChild(a); a.click();
      setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1000);
      setInvoiceCounter((n) => n + 1);
      showToast(`已生成發票 ${invoiceNo}`);
    } catch (e) {
      console.error(e);
      showToast("生成發票失敗，請再試", "error");
    }
  };

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
  myBookings.sort((a, b) => `${b.date}${b.start}`.localeCompare(`${a.date}${a.start}`));

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
          allBookings.push({ date, start: v.start, hours: v.hours, type: v.type, charterType: v.charterType, price: v.price || 0, coachName: v.coachName || "", coach: v.type === "charter" ? null : getCoach(v.coachId), coachId: v.coachId, createdAt: v.createdAt || null, students: v.students || [] });
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

    // 指定月份已上堂數／已售堂數（同「本月總收入」用返同一個時間範圍，等成行 KPI 對得上數）
    const monthUsed = allBookings.filter((b) => b.type !== "charter" && monthKey(b.date) === viewMonth).reduce((s, b) => s + b.hours, 0);
    const monthSold = purchaseLog.filter((r) => monthKey(r.date) === viewMonth).reduce((s, r) => s + r.qty, 0);

    // 各教練總付款（買堂 + 初始）
    const coachPaid = {};
    coaches.forEach((c) => { coachPaid[c.id] = purchaseLog.filter((r) => r.coachId === c.id).reduce((a, r) => a + r.amount, 0) + initialCreditsOf(c) * c.rate; });

    const isSubAdmin = currentUser.role === "subadmin";
    const visibleTabs = [["overview", "📊 總覽"], ["schedule", "📅 課表"], ["coaches", "👥 教練"], ["ledger", "💰 流水帳"], ["records", "📋 記錄"], ["settings", "⚙️ 設定"]]
      .filter(([k]) => !isSubAdmin || currentUser.permissions?.[k]);
    return (
      <div style={S.appBg}>
        <Header title={isSubAdmin ? `副管理員 · ${currentUser.name}` : "管理員"} onLogout={logout} syncState={syncState} />
        {venueNotice && venueNotice.trim() && <div style={S.noticeBanner}>📢 {venueNotice}（教練都見到呢條公告）</div>}
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
              <div style={S.kpiCard}><div style={S.kpiLabel}>{viewMonth === thisMonth ? "本月已上堂數" : "已上堂數"}</div><div style={S.kpiBig}>{monthUsed}</div></div>
              <div style={S.kpiCard}><div style={S.kpiLabel}>{viewMonth === thisMonth ? "本月已售堂數" : "已售堂數"}</div><div style={S.kpiBig}>{monthSold}</div></div>
            </div>
            <p style={S.assistHint}>本月＝{thisMonth}　｜　累計總收入 ${totalRevenue.toLocaleString()}　｜　累計已上堂數 {totalUsed}　｜　累計已售堂數 {totalSold}</p>

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
                                    const lines = buildEntryLines(v, isTrial, c, true);
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
                        <button style={S.linkBtn} onClick={() => generateInvoicePDF(r)}>🧾 開發票</button>
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
                    const { date, start, hours, type, charterType, price, coachName, coach, coachId, students } = b;
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
                            {students && students.length > 0 && <div>學生：{students.join("、")}</div>}
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

            {currentUser.role === "admin" && (
              <>
                <h2 style={{ ...S.sectionTitle, marginTop: 28 }}>場地公告</h2>
                <div style={S.formCard}>
                  <p style={{ ...S.bookingTime, marginBottom: 14, lineHeight: 1.6 }}>例如「本週洗手間維修，請使用更衣室」。所有教練登入會見到呢條提示。留空就唔顯示。</p>
                  <Field label="公告內容"><textarea style={{ ...S.input, minHeight: 70, resize: "vertical" }} value={venueNotice} onChange={(e) => setVenueNotice(e.target.value)} placeholder="留空＝唔顯示" /></Field>
                </div>

                <h2 style={{ ...S.sectionTitle, marginTop: 28 }}>場地 QR Code WhatsApp 號碼</h2>
                <div style={S.formCard}>
                  <p style={{ ...S.bookingTime, marginBottom: 14, lineHeight: 1.6 }}>教練喺「我的預約」撳「攞 QR Code」會自動開 WhatsApp 傳訊息去呢個號碼。請輸入完整國際格式（例如香港：85291234567，唔使 + 號）。</p>
                  <Field label="WhatsApp 號碼"><input style={S.input} placeholder="例如 85291234567" value={whatsappNumber} onChange={(e) => setWhatsappNumber(e.target.value.replace(/[^0-9]/g, ""))} /></Field>
                </div>

                <h2 style={{ ...S.sectionTitle, marginTop: 28 }}>匿名改善建議（只有你睇到）</h2>
                <p style={S.assistHint}>教練透過「意見」分頁匿名提交，系統冇存任何身份資訊，連你都查唔到係邊位教練寫嘅。</p>
                {suggestionBox.length === 0 ? <p style={S.emptyText}>暫無意見</p> : (
                  <div style={S.bookingList}>
                    {suggestionBox.map((sg) => (
                      <div key={sg.id} style={{ ...S.bookingItem, opacity: sg.read ? 0.55 : 1 }}>
                        <div style={{ flex: 1 }}>
                          <div style={S.bookingTime}>{sg.date}</div>
                          <div style={{ ...S.bookingCoach, fontWeight: 400, marginTop: 4, whiteSpace: "pre-wrap" }}>{sg.text}</div>
                        </div>
                        <button style={S.linkBtn} onClick={() => setSuggestionBox((prev) => prev.map((x) => x.id === sg.id ? { ...x, read: !x.read } : x))}>{sg.read ? "標記未閱" : "標記已閱"}</button>
                      </div>
                    ))}
                  </div>
                )}
              </>
            )}

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
                setSubAdmins(DEFAULT_SUBADMINS);
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
      {venueNotice && venueNotice.trim() && <div style={S.noticeBanner}>📢 {venueNotice}</div>}
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
        <button style={view === "other" ? S.tabActive : S.tab} onClick={() => setView("other")}>⚙️ 其他</button>
      </div>

      {view === "calendar" && (
        <div style={S.calContainer}>
          <div style={S.weekNav}>
            <button style={S.navBtn} onClick={() => setWeekOffset((w) => w - 1)}>‹ 上週</button>
            <span style={S.weekLabel}>{formatDate(days[0])} – {formatDate(days[6])}</span>
            <button style={S.navBtn} onClick={() => setWeekOffset((w) => w + 1)}>下週 ›</button>
          </div>
          <p style={S.gridHint}>每格 15 分鐘　｜　同一時段最多 2 名教練　｜　按空格揀 1對1/1對2 同時長</p>
          <div style={S.scaleRow}>
            <span style={S.scaleLabel}>顯示大小</span>
            {[1, 0.85, 0.65].map((v) => (
              <button key={v} style={Math.abs(calScale - v) < 0.001 ? S.scaleBtnActive : S.scaleBtn} onClick={() => updateCalScale(v)}>{Math.round(v * 100)}%</button>
            ))}
            <input style={S.scaleSlider} type="range" min="0.6" max="1" step="0.01" value={calScale} onChange={(e) => updateCalScale(parseFloat(e.target.value))} />
          </div>
          <div style={S.calScroll}>
            <div style={{ zoom: calScale }}>
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
          </div>
          <p style={S.assistHint}>可預約任何時段；24小時內取消需管理員協助。² = 1對2</p>
        </div>
      )}

      {view === "myBookings" && (
        <div style={S.container}>
          <div style={S.flexBetween}>
            <h2 style={S.sectionTitle}>我的預約記錄</h2>
            <div style={S.segRow}>
              <button style={myBookingsView === "list" ? S.segActive : S.seg} onClick={() => setMyBookingsView("list")}>📋 列表</button>
              <button style={myBookingsView === "calendar" ? S.segActive : S.seg} onClick={() => setMyBookingsView("calendar")}>📅 圖像</button>
            </div>
          </div>
          {myBookingsView === "calendar" ? (
            <div style={{ marginTop: 14 }}>
              <div style={S.weekNav}>
                <button style={S.navBtn} onClick={() => setWeekOffset((w) => w - 1)}>‹ 上週</button>
                <span style={S.weekLabel}>{formatDate(days[0])} – {formatDate(days[6])}</span>
                <button style={S.navBtn} onClick={() => setWeekOffset((w) => w + 1)}>下週 ›</button>
              </div>
              <p style={S.gridHint}>自己嘅課堂正常顯示學生名；其他教練嗰格縮細留白，淨係睇到「有人」，等你一眼睇晒成個禮拜邊忙邊閒。撳「列表」可以管理／取消你自己嘅預約</p>
              <div style={S.calScroll}>
                <table style={S.table}>
                  <thead><tr><th style={S.thTime}></th>
                    {days.map((d) => { const today = isTodayDate(d); return <th key={d} style={{ ...S.th, background: today ? "#13302e" : undefined }}><div style={S.dayLabel}>{formatDay(d)}</div><div style={{ ...S.dateLabel, color: today ? "#4ECDC4" : undefined }}>{d.getDate()}</div>{today && <div style={S.todayTag}>今日</div>}</th>; })}
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
                            const mine = here.filter((v) => v.coachId === currentUser.id);
                            const others = here.filter((v) => v.coachId !== currentUser.id);
                            return (
                              <td key={date} style={{ ...S.td, borderTop: isHourStart ? "1px solid #2a2a2a" : "1px solid #161616" }}>
                                {here.length === 0 ? <div style={S.slotDisabled} /> : (
                                  <div style={S.slotMulti}>
                                    {mine.map((v, idx) => {
                                      const label = (v.students && v.students.length > 0) ? v.students.join("、") : (v.type === "duo" ? "1對2" : "1對1");
                                      const span = Math.round(v.hours * 4);
                                      const relRow = slotIndex(time) - slotIndex(v.start);
                                      const showLabel = relRow === 0;
                                      const showBottomTime = relRow === span - 1;
                                      return (
                                        <div key={"m" + idx} style={{ ...S.slotChip, background: liveUser.color + "33", borderLeft: `3px solid ${liveUser.color}` }}>
                                          {showLabel ? <span style={S.slotNameFull}>{label}</span> : showBottomTime ? <span style={S.slotBottomTime}>{addMinutes(v.start, v.hours * 60)}</span> : null}
                                        </div>
                                      );
                                    })}
                                    {others.map((v, idx) => <div key={"o" + idx} style={S.occupiedBar} />)}
                                  </div>
                                )}
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
          ) : myBookings.length === 0 ? <p style={S.emptyText}>你還未有預約</p> : (
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
                      <button style={S.qrBtn} onClick={() => openWhatsAppQR(date, start, hours, liveUser.name)}>📲 攞 QR Code</button>
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

          <h2 style={{ ...S.sectionTitle, marginTop: 28 }}>同步落自己嘅日曆</h2>
          <div style={S.formCard}>
            {!cloudEnabled ? (
              <p style={{ ...S.bookingTime, lineHeight: 1.6 }}>呢個功能需要先開啟雲端同步（而家呢部裝置用嘅係本機儲存）。請聯絡管理員設定 Supabase 雲端同步。</p>
            ) : (
              <>
                <p style={{ ...S.bookingTime, marginBottom: 10, lineHeight: 1.6 }}>生成一個專屬連結，加入 Google Calendar／Apple Calendar「訂閱日曆」，之後你嘅 booking 會自動定期更新，唔使再手動匯出。</p>
                {calendarFeedUrl ? (
                  <>
                    <input style={{ ...S.input, fontSize: 11 }} readOnly value={calendarFeedUrl} onFocus={(e) => e.target.select()} />
                    <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
                      <button style={{ ...S.creditBtn, flex: 1 }} onClick={async () => { try { await navigator.clipboard.writeText(calendarFeedUrl); showToast("已複製連結"); } catch (e) { showToast("複製失敗，請長按手動複製", "error"); } }}>📋 複製連結</button>
                      <button style={{ ...S.smallBtn, flex: 1 }} onClick={regenerateCalendarToken}>🔄 重新生成</button>
                    </div>
                    <p style={{ ...S.assistHint, marginTop: 8 }}>連結等於密碼，請唔好分享畀其他人。重新生成會令舊連結失效。</p>
                  </>
                ) : (
                  <button style={S.loginBtn} onClick={ensureCalendarToken}>生成同步連結</button>
                )}
              </>
            )}
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
                      <Field label="電話（聯絡用，留底備用）"><input style={S.input} value={s.phone || ""} placeholder="例如 85291234567"
                        onChange={(e) => updateStudentField(s.name, "phone", e.target.value.replace(/[^0-9]/g, ""))} /></Field>
                      <div style={S.bookingTime}>已開 {s.credits || 0} 堂　已用 {s.used || 0} 堂</div>
                      <Field label="剩餘堂數">
                        <input style={{ ...S.input, borderColor: low ? "#5a2020" : undefined, color: low ? "#FF8FA3" : "#4ECDC4", fontWeight: 700 }}
                          type="number" value={remain}
                          onChange={(e) => setStudentRemain(s.name, parseInt(e.target.value) || 0)} />
                      </Field>
                      <button style={{ ...S.creditBtn, marginTop: 4, marginBottom: 14 }} onClick={() => setAddStudentCreditModal({ name: s.name, qty: 1 })}>+ 幫佢開堂數</button>

                      <div style={{ ...S.assistHint, marginBottom: 4 }}>購堂紀錄</div>
                      {(() => {
                        const buyLog = studentPurchaseLog.filter((r) => r.coachId === currentUser.id && r.studentName === s.name);
                        return buyLog.length === 0 ? <p style={{ ...S.emptyText, padding: "8px 0" }}>暫無購堂紀錄</p> : (
                          <div style={{ marginBottom: 14 }}>
                            {buyLog.map((r) => (
                              <div key={r.id} style={S.purchaseRow}>
                                <div style={S.bookingTime}>{r.date}　+{r.qty} 堂　@${r.rate}/堂</div>
                                <div style={{ color: "#6BCB77", fontWeight: 700 }}>${r.amount.toLocaleString()}</div>
                              </div>
                            ))}
                          </div>
                        );
                      })()}

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

      {view === "other" && (
        <div style={S.container}>
          <h2 style={S.sectionTitle}>匿名改善建議</h2>
          <div style={S.noticeBanner}>
            🔒 呢個意見箱<strong>完全匿名</strong>。
          </div>
          <div style={{ ...S.formCard, marginTop: 14 }}>
            <Field label="你嘅意見／建議">
              <textarea style={{ ...S.input, minHeight: 100, resize: "vertical" }} value={suggestionText}
                onChange={(e) => setSuggestionText(e.target.value)} placeholder="想提啲咩改善建議？" />
            </Field>
            <button style={S.loginBtn} onClick={() => { submitSuggestion(suggestionText); setSuggestionText(""); }}>匿名提交</button>
          </div>

          <h2 style={{ ...S.sectionTitle, marginTop: 28 }}>修改我的密碼</h2>
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
            <label style={{ ...S.label, marginTop: 14 }}>每週重複（同一星期幾、同一時間）</label>
            <div style={S.segRow}>
              {[1, 4, 8, 12].map((w) => (
                <button key={w} style={(bookModal.repeatWeeks || 1) === w ? S.segActive : S.seg} onClick={() => setBookModal({ ...bookModal, repeatWeeks: w })}>{w === 1 ? "唔重複" : `${w}週`}</button>
              ))}
            </div>
            {(bookModal.repeatWeeks || 1) > 1 && <p style={S.assistHint}>會一次過幫你book未來 {bookModal.repeatWeeks} 個星期嘅同一個時段；如果某一週已經被佔用或堂數不足，會自動跳過嗰一週，唔影響其他週。</p>}
            <div style={S.priceBox}>
              <div style={S.priceRow}><span>時段</span><span>{bookModal.time} – {addMinutes(bookModal.time, bookModal.hours * 60)}</span></div>
              <div style={S.priceRow}><span>扣堂數</span><span>{bookModal.hours} 堂{(bookModal.repeatWeeks || 1) > 1 ? `（每週，最多扣 ${bookModal.hours * bookModal.repeatWeeks} 堂）` : ""}</span></div>
              <div style={{ ...S.priceRow, color: "#4ECDC4", fontWeight: 700, fontSize: 16 }}><span>{isDuo ? "1對2 收費" : "1對1 收費"}</span><span>${price}{(bookModal.repeatWeeks || 1) > 1 ? "／週" : ""}</span></div>
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

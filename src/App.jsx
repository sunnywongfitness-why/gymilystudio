import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { cloudEnabled, cloudLoad, cloudSave, cloudSubscribe, SUPABASE_URL } from "./supabase.js";
import { PDFDocument, rgb, StandardFonts } from "pdf-lib";
import {
  ADMIN_TAB_KEYS, TIME_SLOTS, LS_KEY,
} from "./constants.js";
import {
  BRAND_NAME, LOGO, STAMP_PNG, DEFAULT_COACHES, DEFAULT_SUBADMINS,
  MAX_CONCURRENT, DUO_BASE, DUO_HALF_HOUR_ADD, CHARTER_PRICE,
  ASSIST_CANCEL_LIMIT, LOW_CREDIT_THRESHOLD, CLOSED_DAYS,
  DEFAULT_ADMIN_PASSWORD, COMPANY_LEGAL_NAME, COMPANY_ADDRESS_LINES, INVOICE_THEME_RGB, INVOICE_PREFIX,
  PASS_HOURLY_RATE, PERSONAL_PASS_HOURS, PERSONAL_PASS_MONTHS, FLEXIBLE_PASS_HOURS, FLEXIBLE_PASS_MONTHS, SHARED_PASS_HOURS, SHARED_PASS_MONTHS,
  onboardingFeeSheetText, onboardingVenueRulesText, onboardingPaymentInfoText, onboardingWelcomeText, onboardingRentalGuideText, onboardingTermsText, retroactiveBookingReminderText,
} from "./brand.js";
import {
  persisted, loadSession, saveSession, clearSession, loadCalScale, saveCalScale,
  stableStringify, initialSession, duoPrice, isWholeVenue, rentalShort, rentalFull,
  isClosedDay, getDaysOfWeek, formatDate, isTodayDate, formatDay, monthKey,
  hoursUntil, addMinutes, slotsFor, slotIndex, buildEntryLines, addDaysToDate, addMonthsToDate, coachColorFromId, actorLabel,
} from "./helpers.js";
import { S } from "./styles.js";
import { EditCoachModal, Field, SignaturePad, Header, Toast } from "./components.jsx";

export default function App() {
  const [coaches, setCoaches] = useState(() => persisted("coaches", DEFAULT_COACHES));
  const [currentUser, setCurrentUser] = useState(() => initialSession?.user || null);
  const [adminPassword, setAdminPassword] = useState(() => persisted("adminPassword", DEFAULT_ADMIN_PASSWORD));
  const [whatsappNumber, setWhatsappNumber] = useState(() => persisted("whatsappNumber", ""));
  const [venueNotice, setVenueNotice] = useState(() => persisted("venueNotice", ""));
  const [paymentQR, setPaymentQR] = useState(() => persisted("paymentQR", "")); // 收款 QR code（base64 圖），admin可隨時上傳/更新
  const [suggestionBox, setSuggestionBox] = useState(() => persisted("suggestionBox", []));
  const [adminCalendarToken, setAdminCalendarToken] = useState(() => persisted("adminCalendarToken", ""));
  const [signatureStore, setSignatureStore] = useState(() => persisted("signatureStore", {})); // 簽名圖獨立存一份，唔跟住 booking 喺每個15分鐘格重複
  const [filmingNotices, setFilmingNotices] = useState(() => persisted("filmingNotices", [])); // 拍片被頂走嘅通知
  const [retroBookingNotices, setRetroBookingNotices] = useState(() => persisted("retroBookingNotices", [])); // 第6項：Admin通知教練「已上堂但未book返」，教練首頁banner顯示，可撳入去補book
  const [passUsageLog, setPassUsageLog] = useState(() => persisted("passUsageLog", [])); // {id, coachId, date, hours, passType, sessionType} 教練自己book堂扣Pass時數嘅記錄，用嚟計「個人證」呢類受限類型仲剩幾多
  const [sharedPasses, setSharedPasses] = useState(() => persisted("sharedPasses", [])); // {id, totalHours, usedHours, purchaseDate, expiryDate, coachIds:[id1,id2], usageByCoach:{[id]:hours}} 共享訓練通行證
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
  const [weekViewMode, setWeekViewMode] = useState("fixed"); // fixed（星期一至日） | rolling（以今日做第一日）— 第3項：兩個模式並存，一鍵切換
  const [calScale, setCalScale] = useState(() => loadCalScale());
  const [myBookingsView, setMyBookingsView] = useState("list"); // list | calendar
  const [myBookingsSortMode, setMyBookingsSortMode] = useState("newest"); // newest | closest（距今日最近排最頂）
  const [studentDrafts, setStudentDrafts] = useState({}); // 學生「每堂收費」／「剩餘堂數」輸入緊嘅暫存字串，等撳delete可以留空唔會即刻變返0，key: `${name}_${field}`
  const updateCalScale = (v) => { setCalScale(v); saveCalScale(v); };
  const [bookModal, setBookModal] = useState(null);   // { date, time }
  const [charterModal, setCharterModal] = useState(null); // admin charter { date, time, hours }
  const [slotChoiceModal, setSlotChoiceModal] = useState(null); // { date, time } -> 揀「包場/小組」定「代教練book堂」
  const [adminCoachBookModal, setAdminCoachBookModal] = useState(null); // admin 代教練 book 堂 { date, time, coachId, sessionType, hours, students }
  const [copyInfoModal, setCopyInfoModal] = useState(null); // { text } 代book成功之後嘅複製文字
  const [retroReminderModal, setRetroReminderModal] = useState(null); // 第6項：Admin發起「補book提醒」揀日期時間 { coachId, coachName, date, start, hours }
  const [retroBookModal, setRetroBookModal] = useState(null); // 第6項：教練撳banner後嘅補book表格 { noticeId, date, start, hours, sessionType, students }
  const [quickBook, setQuickBook] = useState({ date: "", start: "19:00", hours: 1, sessionType: "solo", students: [], studentOther: "" }); // 快速Book表格（首頁表格式輸入，取代/補充grid點格仔），淨係教練自己book用
  const [charterLog, setCharterLog] = useState(() => persisted("charterLog", [])); // {date, bookDate, start, hours, amount}
  const [assistCancelLog, setAssistCancelLog] = useState(() => persisted("assistCancelLog", [])); // {coachId, month, date, start}
  const [cancelLog, setCancelLog] = useState(() => persisted("cancelLog", [])); // {date, start, hours, type, charterType, coachId, coachName, price, cancelledBy, cancelledAt}
  const [drinkProducts, setDrinkProducts] = useState(() => persisted("drinkProducts", [])); // {id, name, price} 飲品產品清單，admin喺設定維護
  const [drinkSalesLog, setDrinkSalesLog] = useState(() => persisted("drinkSalesLog", [])); // {id, coachId, coachName, items:[{productId,name,price,qty}], amount, date, time} 飲品銷售記錄，掛落教練account，唔追蹤買家（學生）身份
  const [drinkOrderOpen, setDrinkOrderOpen] = useState(false); // 教練「其他」分頁入面，飲品訂購區塊係咪展開
  const [drinkCart, setDrinkCart] = useState({}); // { [productId]: qty } 揀緊嘅支數，未確認
  const [drinkQrModal, setDrinkQrModal] = useState(null); // { items, amount } 撳「下一步」之後顯示收款QR畀學生睇
  const [newDrinkForm, setNewDrinkForm] = useState({ name: "", price: "" }); // admin新增產品用嘅暫存輸入
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
    if (d.paymentQR !== undefined) setPaymentQR(d.paymentQR);
    if (d.suggestionBox !== undefined) setSuggestionBox(d.suggestionBox);
    if (d.adminCalendarToken !== undefined) setAdminCalendarToken(d.adminCalendarToken);
    if (d.signatureStore !== undefined) setSignatureStore(d.signatureStore);
    if (d.filmingNotices !== undefined) setFilmingNotices(d.filmingNotices);
    if (d.retroBookingNotices !== undefined) setRetroBookingNotices(d.retroBookingNotices);
    if (d.passUsageLog !== undefined) setPassUsageLog(d.passUsageLog);
    if (d.sharedPasses !== undefined) setSharedPasses(d.sharedPasses);
    if (d.invoiceCounter !== undefined) setInvoiceCounter(d.invoiceCounter);
    if (d.studentPurchaseLog !== undefined) setStudentPurchaseLog(d.studentPurchaseLog);
    if (d.subAdmins !== undefined) setSubAdmins(d.subAdmins);
    if (d.bookings !== undefined) setBookings(d.bookings);
    if (d.purchaseLog !== undefined) setPurchaseLog(d.purchaseLog);
    if (d.charterLog !== undefined) setCharterLog(d.charterLog);
    if (d.assistCancelLog !== undefined) setAssistCancelLog(d.assistCancelLog);
    if (d.cancelLog !== undefined) setCancelLog(d.cancelLog);
    if (d.drinkProducts !== undefined) setDrinkProducts(d.drinkProducts);
    if (d.drinkSalesLog !== undefined) setDrinkSalesLog(d.drinkSalesLog);
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
        const seed = { coaches, adminPassword, whatsappNumber, venueNotice, paymentQR, suggestionBox, adminCalendarToken, signatureStore, filmingNotices, retroBookingNotices, passUsageLog, sharedPasses, invoiceCounter, subAdmins, bookings, purchaseLog, studentPurchaseLog, charterLog, assistCancelLog, cancelLog, drinkProducts, drinkSalesLog };
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
    const bundle = { coaches, adminPassword, whatsappNumber, venueNotice, paymentQR, suggestionBox, adminCalendarToken, signatureStore, filmingNotices, retroBookingNotices, passUsageLog, sharedPasses, invoiceCounter, subAdmins, bookings, purchaseLog, studentPurchaseLog, charterLog, assistCancelLog, cancelLog, drinkProducts, drinkSalesLog };
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
  }, [coaches, adminPassword, whatsappNumber, venueNotice, paymentQR, suggestionBox, adminCalendarToken, signatureStore, filmingNotices, retroBookingNotices, passUsageLog, sharedPasses, invoiceCounter, subAdmins, bookings, purchaseLog, studentPurchaseLog, charterLog, assistCancelLog, cancelLog, drinkProducts, drinkSalesLog]);

  const [cancelModal, setCancelModal] = useState(null);
  const [signModal, setSignModal] = useState(null); // {date,start,coachId,type,studentName}
  const [adminCancelModal, setAdminCancelModal] = useState(null); // {date,start,coachId,type}
  const [delLedgerModal, setDelLedgerModal] = useState(null); // ledger record to delete
  const [toast, setToast] = useState(null);
  const [loginForm, setLoginForm] = useState({ id: "", password: "" });
  const [pwForm, setPwForm] = useState({ old: "", new1: "", new2: "" });
  const [editCoach, setEditCoach] = useState(null);
  const [addCreditModal, setAddCreditModal] = useState(null);
  const [sharedPassModal, setSharedPassModal] = useState(null); // {coachIdA, coachIdB, date}（第9項：共享訓練通行證）
  const [sharedTopUpModal, setSharedTopUpModal] = useState(null); // {sharedId, qty}（教練/admin都可以幫共享Pass加值）
  const [adminTab, setAdminTab] = useState(() => initialSession?.adminTab || "overview");
  const [recordsView, setRecordsView] = useState("bookings"); // bookings | cancelled
  const [recCoach, setRecCoach] = useState("all");
  const [recType, setRecType] = useState("all"); // all|solo|duo|private|group|trial
  const [recRange, setRecRange] = useState("upcoming"); // upcoming|past|month|all
  const [recMonth, setRecMonth] = useState(() => monthKey(formatDate(new Date())));
  const [recExpanded, setRecExpanded] = useState(null);
  const [coachSort, setCoachSort] = useState("remain"); // remain|paid|name
  const [expandedCoachId, setExpandedCoachId] = useState(null);
  const [coachDrillView, setCoachDrillView] = useState("purchase"); // purchase | sessions（第6項：教練每月上堂詳情）
  const [viewMonth, setViewMonth] = useState(() => monthKey(formatDate(new Date())));
  const [monthsExpanded, setMonthsExpanded] = useState(false);
  const [newStudentName, setNewStudentName] = useState("");
  const [suggestionText, setSuggestionText] = useState("");
  const [addStudentCreditModal, setAddStudentCreditModal] = useState(null); // {name, qty}
  const [sigReportModal, setSigReportModal] = useState(null); // {studentName, month}（第2項：學生簽名月度報表）
  const [studentLogOpen, setStudentLogOpen] = useState(null);
  const [rosterSortMode, setRosterSortMode] = useState("custom"); // custom | used | remain | name
  const [resetModal, setResetModal] = useState(false);
  const [delCoachModal, setDelCoachModal] = useState(null); // coach pending deletion
  const [showPasswords, setShowPasswords] = useState(false);

  const days = getDaysOfWeek(weekOffset * 7, weekViewMode);
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
    if (coach) { setCurrentUser({ ...coach, role: "coach" }); setView("home"); saveSession({ role: "coach", id: coach.id }); }
    else showToast("帳號或密碼錯誤", "error");
  };
  const logout = () => { setCurrentUser(null); setView("login"); setLoginForm({ id: "", password: "" }); clearSession(); };

  const isCoach = currentUser?.role === "coach";
  const liveUser = isCoach ? getCoach(currentUser.id) : currentUser;
  const remaining = isCoach && liveUser ? liveUser.credits - liveUser.used : 0;
  const soldOut = isCoach && remaining <= 0;

  // bookings[key] is an ARRAY of entries; each entry occupies "seats" (包場/小組=2, 其他=1)
  // 拍片（charterType==="filming"）對冇「允許拍片」權限嘅教練嚎講，唔佔任何位——即係佈哋睇落同空格一樣，照樣book到
  const cellArr = (date, slot) => bookings[`${date}_${slot}`] || [];
  const canSeeFilming = (viewerCoachId) => {
    if (viewerCoachId == null) return true; // null = admin/系統角度，永遠見到全部
    const viewer = getCoach(viewerCoachId);
    return !!(viewer && viewer.allowFilming === true);
  };
  const seats = (entry, viewerCoachId) => {
    if (entry.type === "charter" && entry.charterType === "filming" && !canSeeFilming(viewerCoachId)) return 0;
    return isWholeVenue(entry) ? MAX_CONCURRENT : 1;
  };
  const occupancy = (date, slot, viewerCoachId) => cellArr(date, slot).reduce((n, e) => n + seats(e, viewerCoachId), 0);
  // 畫面顯示用：對冇拍片權限嘅教練，過濾走（唔屬於自己嘅）拍片 entry，等佈睇落係空格
  const visibleCellArr = (date, slot, viewerCoachId) => cellArr(date, slot).filter((e) => {
    if (!(e.type === "charter" && e.charterType === "filming")) return true;
    if (viewerCoachId != null && e.coachId === viewerCoachId) return true; // 自己落嘅拍片，自己一定見到
    return canSeeFilming(viewerCoachId);
  });

  // can we place a booking of `hours` at date/time? need = 需要幾多個位
  // coachId（如果提供）：額外檢查嗰位教練自己係咪已經有重疊嘅 booking（防止同一個教練幫自己撞期）
  // viewerCoachId（如果提供）：用嗰位教練嘅角度判斷拍片睇唔睇到（null = admin 角度，永遠見到）
  const canPlace = (date, time, hours, need = 1, coachId = null, viewerCoachId = undefined) => {
    const vid = viewerCoachId === undefined ? coachId : viewerCoachId;
    const slots = slotsFor(time, hours);
    for (const s of slots) {
      const [hh] = s.split(":").map(Number);
      if (hh >= 22) return "超出營業時間";
      if (occupancy(date, s, vid) + need > MAX_CONCURRENT)
        return need >= MAX_CONCURRENT ? "呢個時段唔夠空（包場／小組需全場）" : "呢個時段已滿（最多2名）";
      if (coachId != null && cellArr(date, s).some((e) => e.coachId === coachId))
        return "呢個時段同你自己另一個預約重疊";
    }
    return null;
  };

  // 自動頂走拍片：如果新 booking 撞到「對嗰位 viewer 嚎講睇唔到」嘅拍片安排，搵返晒嗰啲拍片安排嚎準備取消
  // 用 Map 去重（一個拍片 booking 跨幾個15分鐘格，唔好重複算）
  const findOverriddenFilming = (date, slots, viewerCoachId) => {
    if (canSeeFilming(viewerCoachId)) return []; // 見到拍片嘅人本身會被 canPlace 擋住，唔會行到呢一步
    const found = new Map();
    slots.forEach((s) => {
      cellArr(date, s).forEach((e) => {
        if (e.type === "charter" && e.charterType === "filming") {
          found.set(`${e.coachId}_${e.start}`, { coachId: e.coachId, start: e.start, hours: e.hours, date });
        }
      });
    });
    return Array.from(found.values());
  };

  const openBook = (date, time) => {
    if (isClosedDay(date)) return showToast("星期四、五休息，不開放預約", "error");
    if (soldOut) return showToast("你已用晒購買時數，請聯絡管理員增購", "error");
    const allowSolo = liveUser.allowSolo !== false;
    const allowDuo = liveUser.allowDuo !== false;
    if (!allowSolo && !allowDuo) return showToast("你冇任何可用嘅預約類型，請聯絡管理員設定", "error");
    setBookModal({ date, time, sessionType: allowSolo ? "solo" : "duo", hours: 1, students: [], studentOther: "" });
  };

  // 教練自己落拍片（要有「允許拍片」權限）：佔全場、$0、唔扣時數，對冇權限嘅教練當空格
  const confirmFilmingBooking = () => {
    const { date, time, hours } = bookModal;
    if (liveUser.allowFilming !== true) { showToast("你冇拍片權限", "error"); return; }
    const err = canPlace(date, time, hours, MAX_CONCURRENT, currentUser.id, currentUser.id);
    if (err) { showToast(err, "error"); return; }
    const entry = { coachId: currentUser.id, start: time, hours, type: "charter", charterType: "filming", price: 0, coachName: liveUser.name, students: [], createdAt: new Date().toISOString().slice(0, 16).replace("T", " ") };
    setBookings((prev) => {
      const u = { ...prev };
      slotsFor(time, hours).forEach((s) => { u[`${date}_${s}`] = [...(u[`${date}_${s}`] || []), entry]; });
      return u;
    });
    showToast("已落拍片安排（其他教練見唔到，但有人book中會自動取消）");
    setBookModal(null);
  };

  // src 可傳入覆蓋用嘅booking資料（快速Book表格用），唔傳就用返 bookModal（原本grid點格仔流程，行為完全冇改）
  const confirmBook = (src) => {
    const m = src || bookModal;
    const { date, time, sessionType, hours } = m;
    if (sessionType === "filming") return confirmFilmingBooking();
    if (sessionType === "solo" && liveUser.allowSolo === false) { showToast("你冇一對一預約權限", "error"); return; }
    if (sessionType === "duo" && liveUser.allowDuo === false) { showToast("你冇一對二預約權限", "error"); return; }
    const passCost = sessionType === "duo" ? hours + 0.5 : hours; // 第9項 Training Pass：1對2 喺原定時長之上，額外多扣0.5小時
    const price = passCost * 100; // Training Pass：買咗Pass之後一律 $100/小時計算（唔理solo/duo）
    const rentalCost = price; // Pass制下，收費即係租場成本，冇再獨立計
    const repeatWeeks = Math.max(1, m.repeatWeeks || 1);
    const selected = Array.isArray(m.students) ? m.students : [];
    const extra = (m.studentOther || "").trim();
    const studentList = [...selected, ...(extra ? [extra] : [])].filter(Boolean).slice(0, 4);
    const studentCharges = {};
    studentList.forEach((n) => { const s = myRoster.find((x) => x.name === n); studentCharges[n] = s ? (s.rate || 0) : 0; });

    // 本地追蹤仲有幾多 Pass 時數可用：跨 repeat 週次要遞減，唔可以靠實時 state（loop 入面 state 未更新）
    let personalLeft = personalRemaining(currentUser.id);
    let flexibleLeft = flexibleRemaining(currentUser.id);
    const sharedLeftMap = {};
    sharedPassesOf(currentUser.id).forEach((sp) => { sharedLeftMap[sp.id] = sharedRemaining(sp); });

    const allDeductions = []; // 逐週分配結果，成功晒先一次過 commit
    const newBookingsBySlot = {}; // `${date}_${slot}` -> entry to append
    const filmingToCancel = []; // 因為呢次 book 堂而被自動取消嘅拍片安排
    let okCount = 0, skippedDates = [];
    for (let w = 0; w < repeatWeeks; w++) {
      const wDate = w === 0 ? date : addDaysToDate(date, w * 7);
      // 分池：solo 優先扣「個人證」（solo限定，用晒佢先，唔好嘥），唔夠先用「彈性／舊制」，再唔夠試共享 Pass
      let dedu = null;
      if (sessionType === "solo" && personalLeft >= passCost) dedu = { pool: "personal", amount: passCost };
      else if (flexibleLeft >= passCost) dedu = { pool: "flexible", amount: passCost };
      else {
        const sharedId = Object.keys(sharedLeftMap).find((id) => sharedLeftMap[id] >= passCost);
        if (sharedId) dedu = { pool: "shared", sharedId, amount: passCost };
      }
      if (!dedu) { skippedDates.push(`${wDate}（Pass 時數不足）`); continue; }
      const err = canPlace(wDate, time, hours, 1, currentUser.id);
      if (err) { skippedDates.push(`${wDate}（${err}）`); continue; }
      const wSlots = slotsFor(time, hours);
      findOverriddenFilming(wDate, wSlots, currentUser.id).forEach((f) => filmingToCancel.push(f));
      const logId = "pu" + Date.now() + "-" + w + "-" + Math.random().toString(36).slice(2);
      dedu.id = logId;
      const entry = {
        coachId: currentUser.id, start: time, hours, type: sessionType, price, rentalCost, students: studentList, studentCharges,
        passPool: dedu.pool, passCost: dedu.amount, passLogId: dedu.pool !== "shared" ? logId : null, sharedPassId: dedu.pool === "shared" ? dedu.sharedId : null,
        createdAt: new Date().toISOString().slice(0, 16).replace("T", " "),
      };
      wSlots.forEach((s) => {
        const key = `${wDate}_${s}`;
        newBookingsBySlot[key] = [...(newBookingsBySlot[key] || []), entry];
      });
      if (dedu.pool === "personal") personalLeft -= dedu.amount;
      else if (dedu.pool === "flexible") flexibleLeft -= dedu.amount;
      else sharedLeftMap[dedu.sharedId] -= dedu.amount;
      allDeductions.push(dedu);
      okCount++;
    }

    if (okCount === 0) { showToast(skippedDates[0] || "預約失敗", "error"); return; }

    setBookings((prev) => {
      const u = { ...prev };
      // 先取消被頂走嘅拍片（移走嗰啲安排嘅全部15分鐘格副本）
      filmingToCancel.forEach((f) => {
        slotsFor(f.start, f.hours).forEach((s) => {
          const key = `${f.date}_${s}`;
          u[key] = (u[key] || []).filter((e) => !(e.type === "charter" && e.charterType === "filming" && e.coachId === f.coachId && e.start === f.start));
        });
      });
      Object.entries(newBookingsBySlot).forEach(([key, arr]) => { u[key] = [...(u[key] || []), ...arr]; });
      return u;
    });
    if (filmingToCancel.length > 0) {
      setFilmingNotices((prev) => [...filmingToCancel.map((f) => ({ id: "fn" + Date.now() + "-" + Math.random().toString(36).slice(2), coachId: f.coachId, date: f.date, start: f.start, hours: f.hours, read: false })), ...prev]);
    }
    commitPassDeduction(currentUser.id, sessionType, allDeductions);
    if (repeatWeeks > 1) {
      showToast(skippedDates.length === 0 ? `已成功預約 ${okCount} 週` : `已預約 ${okCount} 週，跳過 ${skippedDates.length} 週：${skippedDates.join("、")}`);
    } else {
      showToast("預約成功！");
    }
    setBookModal(null);
  };

  // 第6項：Admin發起「補book提醒」——揀返教練漏book嘅日期時間，寫入 retroBookingNotices（教練首頁banner會顯示），
  // 有電話就直接開WhatsApp，冇電話就fallback去copyInfoModal俾admin自己copy
  const sendRetroReminder = () => {
    const { coachId, coachName, date, start, hours } = retroReminderModal;
    if (!date || !start || !hours) { showToast("請填晒日期／時間／時長", "error"); return; }
    const end = addMinutes(start, Number(hours) * 60);
    const notice = { id: "rb" + Date.now() + "-" + Math.random().toString(36).slice(2), coachId, date, start, hours: Number(hours), read: false, createdAt: new Date().toISOString() };
    setRetroBookingNotices((prev) => [notice, ...prev]);
    const text = retroactiveBookingReminderText(coachName, date, start, end);
    const coach = getCoach(coachId);
    if (coach?.phone) {
      window.open(`https://wa.me/${coach.phone}?text=${encodeURIComponent(text)}`, "_blank");
    } else {
      setCopyInfoModal({ title: "提醒教練補book堂", text });
    }
    showToast("已發送補book提醒");
    setRetroReminderModal(null);
  };

  // 第6項：教練撳首頁banner後嘅補book——跟confirmBook同一套Pass分配邏輯（個人證優先→彈性→共享），但淨係單一時段、唔重複
  // 過去時段都要做完整衝突檢查（如果嗰時段已經俾第二個教練book咗，要擋住畀admin人手處理）
  const confirmRetroBooking = () => {
    const { noticeId, date, start, hours, sessionType } = retroBookModal;
    if (sessionType === "solo" && liveUser.allowSolo === false) { showToast("你冇一對一預約權限", "error"); return; }
    if (sessionType === "duo" && liveUser.allowDuo === false) { showToast("你冇一對二預約權限", "error"); return; }
    const err = canPlace(date, start, hours, 1, currentUser.id);
    if (err) { showToast(`呢個時段${err}，請聯絡Admin處理`, "error"); return; }
    const passCost = sessionType === "duo" ? hours + 0.5 : hours;
    const price = passCost * 100;
    const rentalCost = price;
    const selected = Array.isArray(retroBookModal.students) ? retroBookModal.students : [];
    const extra = (retroBookModal.studentOther || "").trim();
    const studentList = [...selected, ...(extra ? [extra] : [])].filter(Boolean).slice(0, 4);
    const studentCharges = {};
    studentList.forEach((n) => { const s = myRoster.find((x) => x.name === n); studentCharges[n] = s ? (s.rate || 0) : 0; });

    const personalLeft = personalRemaining(currentUser.id);
    const flexibleLeft = flexibleRemaining(currentUser.id);
    let dedu = null;
    if (sessionType === "solo" && personalLeft >= passCost) dedu = { pool: "personal", amount: passCost };
    else if (flexibleLeft >= passCost) dedu = { pool: "flexible", amount: passCost };
    else {
      const sp = sharedPassesOf(currentUser.id).find((s) => sharedRemaining(s) >= passCost);
      if (sp) dedu = { pool: "shared", sharedId: sp.id, amount: passCost };
    }
    if (!dedu) { showToast("Pass時數不足，請聯絡Admin增購", "error"); return; }
    const logId = "pu" + Date.now() + "-" + Math.random().toString(36).slice(2);
    dedu.id = logId;
    const entry = {
      coachId: currentUser.id, start, hours, type: sessionType, price, rentalCost, students: studentList, studentCharges,
      passPool: dedu.pool, passCost: dedu.amount, passLogId: dedu.pool !== "shared" ? logId : null, sharedPassId: dedu.pool === "shared" ? dedu.sharedId : null,
      isRetroactive: true, // 第6項：標記呢個係事後補記錄，方便Admin流水帳分辨
      createdAt: new Date().toISOString().slice(0, 16).replace("T", " "),
    };
    setBookings((prev) => {
      const u = { ...prev };
      slotsFor(start, hours).forEach((s) => { u[`${date}_${s}`] = [...(u[`${date}_${s}`] || []), entry]; });
      return u;
    });
    commitPassDeduction(currentUser.id, sessionType, [dedu]);
    if (noticeId) setRetroBookingNotices((prev) => prev.map((n) => n.id === noticeId ? { ...n, read: true } : n));
    showToast("已補返記錄！");
    setRetroBookModal(null);
  };

  // ---- 飲品訂購（admin管理產品，教練自己揀支數、顯示收款QR畀學生睇，唔追蹤學生身份，記錄掛落教練account）----
  const addDrinkProduct = () => {
    const name = newDrinkForm.name.trim();
    const price = Number(newDrinkForm.price);
    if (!name) { showToast("請輸入產品名", "error"); return; }
    if (!price || price <= 0) { showToast("請輸入有效價錢", "error"); return; }
    setDrinkProducts((prev) => [...prev, { id: "dp" + Date.now() + "-" + Math.random().toString(36).slice(2), name, price }]);
    setNewDrinkForm({ name: "", price: "" });
  };
  const updateDrinkProduct = (id, field, value) => {
    setDrinkProducts((prev) => prev.map((p) => p.id === id ? { ...p, [field]: value } : p));
  };
  const removeDrinkProduct = (id) => {
    setDrinkProducts((prev) => prev.filter((p) => p.id !== id));
  };
  const drinkCartTotal = () => drinkProducts.reduce((sum, p) => sum + (Number(drinkCart[p.id]) || 0) * (Number(p.price) || 0), 0);
  const drinkCartCount = () => Object.values(drinkCart).reduce((sum, q) => sum + (Number(q) || 0), 0);
  const openDrinkCheckout = () => {
    const items = drinkProducts
      .filter((p) => Number(drinkCart[p.id]) > 0)
      .map((p) => ({ productId: p.id, name: p.name, price: p.price, qty: Number(drinkCart[p.id]) }));
    if (items.length === 0) { showToast("請先揀飲品數量", "error"); return; }
    const amount = items.reduce((sum, it) => sum + it.price * it.qty, 0);
    setDrinkQrModal({ items, amount });
  };
  const confirmDrinkSale = () => {
    if (!drinkQrModal) return;
    const sale = {
      id: "ds" + Date.now() + "-" + Math.random().toString(36).slice(2),
      coachId: currentUser.id,
      coachName: currentUser.name,
      items: drinkQrModal.items,
      amount: drinkQrModal.amount,
      date: formatDate(new Date()),
      time: new Date().toTimeString().slice(0, 5),
    };
    setDrinkSalesLog((prev) => [sale, ...prev]);
    setDrinkCart({});
    setDrinkQrModal(null);
    setDrinkOrderOpen(false);
    showToast("已記錄，請提提學生完成轉數");
  };

  // ADMIN: place a rental (包場/小組=全場2位, 試堂=1位), price editable
  const confirmCharter = () => {
    const { date, time, charterType, price, coachName } = charterModal;
    const hours = Number(charterModal.hours) || 0;
    if (hours <= 0) { showToast("請輸入有效時長", "error"); return; }
    if (isClosedDay(date)) { showToast("休息日", "error"); return; }
    const need = charterType === "trial" ? 1 : MAX_CONCURRENT;
    const amt = ["trial", "clean"].includes(charterType) ? 0 : (parseInt(price) || 0);
    const repeatWeeks = charterType === "clean" ? Math.max(1, charterModal.repeatWeeks || 1) : 1;

    if (repeatWeeks === 1) {
      const err = canPlace(date, time, hours, need);
      if (err) { showToast(err, "error"); return; }
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
      return;
    }

    // 清潔可重複book固定時間（第5項）：跟返 bookModal 每週重複嘅 skip-on-conflict 模式，某一週撞咗就跳過，唔影響其他週
    const newBookingsBySlot = {};
    const logsToAdd = [];
    let okCount = 0, skippedDates = [];
    for (let w = 0; w < repeatWeeks; w++) {
      const wDate = w === 0 ? date : addDaysToDate(date, w * 7);
      const err = canPlace(wDate, time, hours, need);
      if (err) { skippedDates.push(`${wDate}（${err}）`); continue; }
      const wSlots = slotsFor(time, hours);
      const entry = { coachId: 0, start: time, hours, type: "charter", charterType, price: amt, coachName: coachName || "", createdAt: new Date().toISOString().slice(0, 16).replace("T", " ") };
      wSlots.forEach((s) => {
        const key = `${wDate}_${s}`;
        newBookingsBySlot[key] = [...(newBookingsBySlot[key] || []), entry];
      });
      logsToAdd.push({ date: new Date().toISOString().slice(0, 16).replace("T", " "), bookDate: wDate, start: time, hours, charterType, amount: amt, coachName: coachName || "" });
      okCount++;
    }
    if (okCount === 0) { showToast(skippedDates[0] || "預約失敗", "error"); return; }
    setBookings((prev) => {
      const u = { ...prev };
      Object.entries(newBookingsBySlot).forEach(([key, arr]) => { u[key] = [...(u[key] || []), ...arr]; });
      return u;
    });
    setCharterLog((prev) => [...logsToAdd, ...prev]);
    showToast(skippedDates.length === 0 ? `已成功預約 ${okCount} 週清潔` : `已預約 ${okCount} 週清潔，跳過 ${skippedDates.length} 週：${skippedDates.join("、")}`);
    setCharterModal(null);
  };

  // Admin 代教練 book 堂：同教練自己 book 嘅邏輯一樣（扣嗰位教練嘅時數），完成後生成一段文字畀 admin 自己複製去send
  const confirmAdminCoachBooking = () => {
    const { date, time, coachId, sessionType, hours, students } = adminCoachBookModal;
    const coach = getCoach(coachId);
    if (!coach) { showToast("請揀教練", "error"); return; }
    const creditCost = sessionType === "duo" ? hours + 0.5 : hours; // 同教練自己book（confirmBook）睇齊：duo 都要多扣0.5小時，等個共用池嘅「1單位=$100」估值喺兩條path都啱數（唔改收費本身，$150依然係$150）
    const remain = coach.credits - coach.used;
    if (creditCost > remain) { showToast(`${coach.name} 剩餘時數不足`, "error"); return; }
    const err = canPlace(date, time, hours, 1, coachId);
    if (err) { showToast(err, "error"); return; }
    const price = sessionType === "duo" ? duoPrice(hours) : coach.rate * hours;
    const rentalCost = coach.rate * hours;
    const studentList = (students || []).filter(Boolean).slice(0, 4);
    const coachRoster = getStudentRoster(coachId);
    const studentCharges = {};
    studentList.forEach((n) => { const s = coachRoster.find((x) => x.name === n); studentCharges[n] = s ? (s.rate || 0) : 0; });
    const slots = slotsFor(time, hours);
    const filmingToCancel = findOverriddenFilming(date, slots, coachId);
    const entry = { coachId, start: time, hours, type: sessionType, price, rentalCost, students: studentList, studentCharges, createdAt: new Date().toISOString().slice(0, 16).replace("T", " ") };
    setBookings((prev) => {
      const u = { ...prev };
      filmingToCancel.forEach((f) => {
        slotsFor(f.start, f.hours).forEach((s) => {
          const key = `${f.date}_${s}`;
          u[key] = (u[key] || []).filter((e) => !(e.type === "charter" && e.charterType === "filming" && e.coachId === f.coachId && e.start === f.start));
        });
      });
      slots.forEach((s) => { u[`${date}_${s}`] = [...(u[`${date}_${s}`] || []), entry]; });
      return u;
    });
    if (filmingToCancel.length > 0) {
      setFilmingNotices((prev) => [...filmingToCancel.map((f) => ({ id: "fn" + Date.now() + "-" + Math.random().toString(36).slice(2), coachId: f.coachId, date: f.date, start: f.start, hours: f.hours, read: false })), ...prev]);
    }
    setCoaches((prev) => prev.map((c) => c.id === coachId ? { ...c, used: c.used + creditCost } : c));
    const summary = `你好 ${coach.name}，管理員已幫你預約：\n日期：${date}\n時間：${time}–${addMinutes(time, hours * 60)}\n類型：${sessionType === "duo" ? "1對2" : "1對1"}${studentList.length ? `\n學生：${studentList.join("、")}` : ""}`;
    setAdminCoachBookModal(null);
    setCopyInfoModal({ text: summary });
    showToast("已代教練預約");
  };

  const openCancel = (date, start, coachId, type, charterType) => {
    if (type === "charter" && charterType === "filming") { setCancelModal({ date, start, coachId, type }); return; } // 拍片唔涉及堂數，自己隨時可以取消
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
    // 第12項：記低實際操作者身份——subadmin 記返實際姓名，唔再淨係得籠統嘅"admin"
    const actorTag = byAdmin
      ? (currentUser.role === "subadmin" ? `subadmin:${currentUser.name}` : "admin")
      : "coach";
    setCancelLog((prev) => [{
      date, start, hours: meta.hours, type: meta.type, charterType: meta.charterType || null,
      coachId: meta.type === "charter" ? null : coachId,
      coachName: meta.type === "charter" ? (meta.coachName || "") : (getCoach(coachId)?.name || ""),
      price: meta.price || 0,
      cancelledBy: actorTag,
      cancelledAt: new Date().toISOString().slice(0, 16).replace("T", " "),
    }, ...prev]);
    if (meta.type !== "charter") {
      if (meta.passPool === "shared" && meta.sharedPassId) {
        // 共享 Pass 退款：完全唔掂 coach.credits/used，淨係退返嗰張共享 Pass 自己嘅額度
        setSharedPasses((prev) => prev.map((sp) => sp.id === meta.sharedPassId
          ? { ...sp, usedHours: Math.max(0, (sp.usedHours || 0) - (meta.passCost || 0)), usageByCoach: { ...(sp.usageByCoach || {}), [coachId]: Math.max(0, ((sp.usageByCoach || {})[coachId] || 0) - (meta.passCost || 0)) } }
          : sp));
      } else if (meta.passPool && meta.passLogId) {
        // 個人證／彈性池退款：移走返個對應嘅 passUsageLog 記錄，同退返 coach.used
        setPassUsageLog((prev) => prev.filter((x) => x.id !== meta.passLogId));
        setCoaches((prev) => prev.map((c) => c.id === coachId ? { ...c, used: Math.max(0, c.used - (meta.passCost || meta.hours)) } : c));
      } else {
        // 冇 passPool 標記：舊制／Admin代教練book嘅單次收費堂，退返原本嘅時數（duo 都要退多0.5，同扣減嗰刻對稱）
        const refundAmount = meta.type === "duo" ? meta.hours + 0.5 : meta.hours;
        setCoaches((prev) => prev.map((c) => c.id === coachId ? { ...c, used: Math.max(0, c.used - refundAmount) } : c));
      }
      // 由管理員協助、而且係該教練設定嘅通知時數內取消，計入本月額度
      const win = getCoach(coachId)?.cancelWindowHours ?? 24;
      if (byAdmin && hoursUntil(date, start) < win) {
        setAssistCancelLog((prev) => [{ coachId, month: monthKey(formatDate(new Date())), date, start, by: actorTag }, ...prev]);
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
    const signedAt = new Date().toISOString().slice(0, 16).replace("T", " ");
    const sigKey = `${coachId}_${date}_${start}_${studentName}`;
    // 簽名圖（重嘅 base64 PNG）獨立存一份，唔再跟住 booking 喺每個15分鐘格重複存——大幅減少資料庫用量
    setSignatureStore((prev) => ({ ...prev, [sigKey]: { dataUrl, signedAt } }));
    setBookings((prev) => {
      const u = { ...prev };
      slots.forEach((s) => {
        const arr = (u[`${date}_${s}`] || []).map((e) => {
          if (e.coachId === coachId && e.start === start && e.type === type) {
            return { ...e, signatures: { ...(e.signatures || {}), [studentName]: signedAt } };
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
  const normStudent = (s) => {
    const obj = typeof s === "string" ? { name: s, rate: 0, credits: 0, used: 0 } : { rate: 0, credits: 0, used: 0, ...s };
    delete obj.phone; // 私隱考慮：唔再保留學生電話，亦主動清走舊有已存嘅電話資料
    return obj;
  };
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

  // 自由調位：上／落移一個位（直接改返存住嗰個次序，自訂排序模式先會跟住呢個次序顯示）
  const moveStudent = (name, dir) => {
    const idx = myRoster.findIndex((s) => s.name === name);
    const j = idx + dir;
    if (idx === -1 || j < 0 || j >= myRoster.length) return;
    const arr = [...myRoster];
    [arr[idx], arr[j]] = [arr[j], arr[idx]];
    setCoaches((prev) => prev.map((c) => c.id === currentUser.id ? { ...c, studentRoster: arr } : c));
  };

  // 幫學生開堂數（教練自己用，類似 admin 幫教練「+ 堂」嗰個概念）
  const addStudentCredits = (name, qty, date, expiryDate) => {
    const s = myRoster.find((x) => x.name === name);
    if (!s) return;
    updateStudentField(name, "credits", (s.credits || 0) + qty);
    setStudentPurchaseLog((prev) => [{ id: "sp" + Date.now() + "-" + Math.random().toString(36).slice(2), coachId: currentUser.id, studentName: name, date: date || new Date().toISOString().slice(0, 10), expiryDate: expiryDate || "", qty, rate: s.rate || 0, amount: (s.rate || 0) * qty }, ...prev]);
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

  // Admin 自己嘅日曆同步（全場所有教練嘅 booking，連教練名都顯示）
  const ensureAdminCalendarToken = () => {
    if (adminCalendarToken) return adminCalendarToken;
    const t = genToken();
    setAdminCalendarToken(t);
    return t;
  };
  const regenerateAdminCalendarToken = () => {
    const t = genToken();
    setAdminCalendarToken(t);
    showToast("已重新生成連結，舊連結會失效");
  };
  const adminCalendarFeedUrl = adminCalendarToken ? `${SUPABASE_URL}/functions/v1/coach-calendar?adminToken=${adminCalendarToken}` : "";

  // 上傳收款 QR：縮到最大邊長 512px、輸出 JPEG（quality 0.8），控制 base64 檔案大細，減輕 jsonb 資料庫負擔
  const handleQRUpload = (file) => {
    if (!file) return;
    if (!file.type.startsWith("image/")) { showToast("請揀圖片檔案", "error"); return; }
    const reader = new FileReader();
    reader.onload = (ev) => {
      const img = new Image();
      img.onload = () => {
        const MAX = 512;
        let { width, height } = img;
        if (width > height && width > MAX) { height = Math.round(height * MAX / width); width = MAX; }
        else if (height > MAX) { width = Math.round(width * MAX / height); height = MAX; }
        const canvas = document.createElement("canvas");
        canvas.width = width; canvas.height = height;
        const ctx = canvas.getContext("2d");
        ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, width, height); // 白底，避免透明 QR 掃唔到
        ctx.drawImage(img, 0, 0, width, height);
        setPaymentQR(canvas.toDataURL("image/jpeg", 0.8));
        showToast("已更新收款 QR Code");
      };
      img.onerror = () => showToast("圖片讀取失敗", "error");
      img.src = ev.target.result;
    };
    reader.onerror = () => showToast("檔案讀取失敗", "error");
    reader.readAsDataURL(file);
  };

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

  const addCredits = (coachId, qty, date, expiryDate, passType = null) => {
    const coach = getCoach(coachId);
    const rate = passType ? PASS_HOURLY_RATE : coach.rate;
    const amount = qty * rate;
    const actorTag = currentUser.role === "subadmin" ? `subadmin:${currentUser.name}` : "admin";
    setCoaches((prev) => prev.map((c) => c.id === coachId ? { ...c, credits: c.credits + qty } : c));
    setPurchaseLog((prev) => [{ id: "p" + Date.now() + "-" + Math.random().toString(36).slice(2), date: date || new Date().toISOString().slice(0, 10), coachId, coachName: coach.name, qty, amount, rate, expiryDate: expiryDate || null, addedBy: actorTag, passType: passType || null }, ...prev]);
    showToast(`已為 ${coach.name} 增加 ${qty} 小時${passType ? `（${passType === "personal" ? "個人證" : "彈性證"}）` : ""}（$${amount}）${expiryDate ? `，失效日：${expiryDate}` : ""}`);
  };

  // 開一張共享訓練通行證：硬性2位教練，30小時／12個月，兩位教練其中一位或admin都可以之後幫佢加值（見addSharedPassHours）
  const createSharedPass = (coachIdA, coachIdB, date) => {
    if (coachIdA === coachIdB) { showToast("要揀兩位唔同嘅教練", "error"); return; }
    const id = "sp" + Date.now() + "-" + Math.random().toString(36).slice(2);
    const purchaseDate = date || formatDate(new Date());
    setSharedPasses((prev) => [{
      id, totalHours: SHARED_PASS_HOURS, usedHours: 0, purchaseDate,
      expiryDate: addMonthsToDate(purchaseDate, SHARED_PASS_MONTHS),
      coachIds: [coachIdA, coachIdB], usageByCoach: { [coachIdA]: 0, [coachIdB]: 0 },
    }, ...prev]);
    showToast(`已開共享 Pass（${getCoach(coachIdA)?.name} ＋ ${getCoach(coachIdB)?.name}，${SHARED_PASS_HOURS}小時）`);
  };
  // 幫一張共享 Pass 加時數：兩位當事教練其中一位，或者 admin，都可以做
  const addSharedPassHours = (sharedId, qty) => {
    setSharedPasses((prev) => prev.map((sp) => sp.id === sharedId ? { ...sp, totalHours: (sp.totalHours || 0) + qty } : sp));
    showToast(`已為共享 Pass 增加 ${qty} 小時`);
  };

  // FIFO：將某教練嘅 used 時數，依購買時間順序分配到每筆購買記錄，計出每筆嘅「已用／剩餘」
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

  // 邊位教練有「未用完、已過期或快過期」嘅時數（用嚎提醒管理員／教練）
  const EXPIRY_WARN_DAYS = 14;
  const expiringBatchesOf = (coachId) => {
    const today = formatDate(new Date());
    return purchaseFifoStatus(coachId).filter((b) => b.remaining > 0 && b.expiryDate &&
      (new Date(`${b.expiryDate}T00:00:00`) - new Date(`${today}T00:00:00`)) / 86400000 <= EXPIRY_WARN_DAYS);
  };
  // 已經過期但仲有剩餘時數嘅 Pass 批次（過期唔鎖，淨係提醒）
  const expiredPassBatchesOf = (coachId) => {
    const today = formatDate(new Date());
    return purchaseFifoStatus(coachId).filter((b) => b.remaining > 0 && b.passType && b.expiryDate && b.expiryDate < today);
  };

  // ===== 第9項：Training Pass 制度 =====
  // 呢位教練「個人訓練通行證」（solo限定）總共買咗幾多小時
  const personalPurchased = (coachId) => purchaseLog.filter((r) => r.coachId === coachId && r.passType === "personal").reduce((s, r) => s + r.qty, 0);
  // 呢位教練喺 passUsageLog 入面，用咗幾多「個人證」時數（只有教練自己 book 堂先會計入呢個 log）
  const personalConsumed = (coachId) => passUsageLog.filter((r) => r.coachId === coachId && r.passType === "personal").reduce((s, r) => s + r.hours, 0);
  const personalRemaining = (coachId) => Math.max(0, personalPurchased(coachId) - personalConsumed(coachId));
  // 「彈性／舊制」呢個唔限類型嘅池：總剩餘（credits-used）減去「個人證」嗰截，即係solo/duo/包場都用得嘅部分
  const flexibleRemaining = (coachId) => {
    const c = getCoach(coachId);
    if (!c) return 0;
    const total = (c.credits || 0) - (c.used || 0);
    return Math.max(0, total - personalRemaining(coachId));
  };
  // 呢位教練有份嘅共享 Pass（可能有0張、1張，硬性上限2位教練/張）
  const sharedPassesOf = (coachId) => sharedPasses.filter((sp) => (sp.coachIds || []).includes(coachId));
  const sharedRemaining = (sp) => Math.max(0, (sp.totalHours || 0) - (sp.usedHours || 0));

  // 教練自己 book solo/duo 堂：計算收費同扣邊個池，$100/小時計，1對2 額外多扣0.5小時
  // 回傳 { ok, price, deductions:[{pool,amount}], error }，pool: "personal" | "flexible" | "shared"
  const allocatePassHours = (coachId, sessionType, hours) => {
    const need = sessionType === "duo" ? hours + 0.5 : hours;
    const price = need * 100;
    if (sessionType === "solo") {
      const pRemain = personalRemaining(coachId);
      if (pRemain >= need) return { ok: true, price, need, deductions: [{ pool: "personal", amount: need }] };
    }
    const fRemain = flexibleRemaining(coachId);
    if (fRemain >= need) return { ok: true, price, need, deductions: [{ pool: "flexible", amount: need }] };
    // 自己個人／彈性池都唔夠：試吓有冇共享 Pass 夠用（最後手段，優先用自己嘅時數）
    const shared = sharedPassesOf(coachId).find((sp) => sharedRemaining(sp) >= need);
    if (shared) return { ok: true, price, need, deductions: [{ pool: "shared", sharedId: shared.id, amount: need }] };
    return { ok: false, error: `Pass 時數不足（需要 ${need} 小時），請聯絡 admin 增購` };
  };

  // 實際執行 Pass 扣鐘：personal/flexible 池會記落 passUsageLog + 扣 coach.used（同舊制credits/used共用同一組數字）；
  // shared 池完全獨立，唔會掂 coach.credits/used，淨係扣返嗰張共享 Pass 自己嘅 usedHours/usageByCoach
  const commitPassDeduction = (coachId, sessionType, deductions) => {
    const nonShared = deductions.filter((d) => d.pool !== "shared").reduce((s, d) => s + d.amount, 0);
    if (nonShared > 0) {
      setCoaches((prev) => prev.map((c) => c.id === coachId ? { ...c, used: c.used + nonShared } : c));
      setPassUsageLog((prev) => [
        ...deductions.filter((d) => d.pool !== "shared").map((d) => ({
          id: d.id || ("pu" + Date.now() + "-" + Math.random().toString(36).slice(2)),
          coachId, date: formatDate(new Date()), hours: d.amount, passType: d.pool, sessionType,
        })),
        ...prev,
      ]);
    }
    deductions.filter((d) => d.pool === "shared").forEach((d) => {
      setSharedPasses((prev) => prev.map((sp) => sp.id === d.sharedId
        ? { ...sp, usedHours: (sp.usedHours || 0) + d.amount, usageByCoach: { ...(sp.usageByCoach || {}), [coachId]: ((sp.usageByCoach || {})[coachId] || 0) + d.amount } }
        : sp));
    });
  };

  // sanitize sheet names (Excel: <=31 chars, no : \ / ? * [ ])
  const sheetName = (s) => (s || "").replace(/[:\\/?*[\]]/g, " ").slice(0, 28).trim() || "Sheet";
  const fmtMoney = (n) => n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  // 生成購買堂數 Invoice（PDF），跟住公司提供嗰張範本排版；個性化資料（Bill to / Contact Person）留空，
  // 其他全部跟住教練同呢筆購買記錄自動帶入；公司印章自動貼上。
  const generateInvoicePDF = async (record) => {
    try {
      const teal = rgb(...INVOICE_THEME_RGB);
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
      page.drawText(COMPANY_LEGAL_NAME, { x: rightX - bold.widthOfTextAtSize(COMPANY_LEGAL_NAME, 13), y: height - 35, size: 13, font: bold, color: rgb(1, 1, 1) });
      const addrLines = COMPANY_ADDRESS_LINES;
      addrLines.forEach((line, i) => {
        page.drawText(line, { x: rightX - font.widthOfTextAtSize(line, 9), y: height - 50 - i * 12, size: 9, font, color: rgb(1, 1, 1) });
      });
      page.drawText("Invoice", { x: marginX, y: height - headerH + 22, size: 26, font: bold, color: rgb(1, 1, 1) });

      // ---- Bill to / Contact Person（個性化資料，留空） ----
      let y = height - headerH - 35;
      page.drawText("Bill to :", { x: marginX, y, size: 9, font: bold, color: black });
      page.drawText("Contact Person:", { x: marginX + 180, y, size: 9, font: bold, color: black });

      // ---- 右側：根據教練／購買記錄自動帶入 ----
      const invoiceNo = `${INVOICE_PREFIX}${String(new Date().getFullYear()).slice(2)}${String(invoiceCounter).padStart(4, "0")}`;
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

  // 第2項：學生簽名月度報表——教練自己攞返某個學生某個月嘅上堂記錄連簽名圖，冇簽到嘅堂都會列出並標記「未簽」
  const generateSignatureReportPDF = async (studentName, month) => {
    try {
      const rows = myBookings
        .filter((b) => (b.students || []).includes(studentName) && monthKey(b.date) === month)
        .sort((a, b) => `${a.date}${a.start}`.localeCompare(`${b.date}${b.start}`));
      if (rows.length === 0) { showToast(`${month} 冇 ${studentName} 嘅上堂記錄`, "error"); return; }

      const teal = rgb(...INVOICE_THEME_RGB);
      const black = rgb(0.1, 0.1, 0.1);
      const grey = rgb(0.6, 0.6, 0.6);
      const red = rgb(0.8, 0.2, 0.2);

      const pdfDoc = await PDFDocument.create();
      const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
      const bold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
      const marginX = 42;
      const pageW = 595.28, pageH = 841.89;
      const rightX = pageW - marginX;
      let page, y;

      const newPage = () => {
        page = pdfDoc.addPage([pageW, pageH]);
        page.drawRectangle({ x: 0, y: pageH - 70, width: pageW, height: 70, color: teal });
        page.drawText(`${studentName} 上堂簽到記錄`, { x: marginX, y: pageH - 40, size: 16, font: bold, color: rgb(1, 1, 1) });
        page.drawText(`${month}　教練：${liveUser.name}`, { x: marginX, y: pageH - 58, size: 10, font, color: rgb(1, 1, 1) });
        y = pageH - 95;
      };
      newPage();

      for (const b of rows) {
        const rowH = 60;
        if (y - rowH < 60) newPage();
        page.drawText(`${b.date}（${formatDay(new Date(`${b.date}T00:00:00`))}）  ${b.start}–${addMinutes(b.start, b.hours * 60)}（${b.hours}小時）  ${b.type === "duo" ? "1對2" : "1對1"}`, { x: marginX, y, size: 10, font, color: black });
        const sigKey = `${liveUser.id}_${b.date}_${b.start}_${studentName}`;
        const sig = signatureStore[sigKey];
        if (sig && sig.dataUrl) {
          try {
            const sigBytes = await fetch(sig.dataUrl).then((r) => r.arrayBuffer());
            const sigImg = await pdfDoc.embedPng(sigBytes);
            const sigW = 110, sigH = sigImg.height * (sigW / sigImg.width);
            page.drawRectangle({ x: rightX - sigW - 8, y: y - sigH + 4, width: sigW + 8, height: sigH + 4, borderColor: grey, borderWidth: 0.5 });
            page.drawImage(sigImg, { x: rightX - sigW - 4, y: y - sigH + 6, width: sigW, height: sigH });
            y -= (sigH + 14);
          } catch (e) {
            page.drawText("（簽名圖讀取失敗）", { x: rightX - 110, y: y - 12, size: 9, font, color: red });
            y -= 40;
          }
        } else {
          page.drawText("⚠ 未簽", { x: rightX - 40, y, size: 10, font: bold, color: red });
          y -= 40;
        }
        page.drawLine({ start: { x: marginX, y: y + 6 }, end: { x: rightX, y: y + 6 }, thickness: 0.5, color: grey });
        y -= 8;
      }

      const pdfBytes = await pdfDoc.save();
      const blob = new Blob([pdfBytes], { type: "application/pdf" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = `簽到記錄_${studentName}_${month}.pdf`;
      document.body.appendChild(a); a.click();
      setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1000);
      showToast("已生成簽到月報表");
    } catch (e) {
      console.error(e);
      showToast("生成報表失敗，請再試", "error");
    }
  };


  const exportMyIncomeSheet = () => {
    try {
      const wb = XLSX.utils.book_new();
      const monthRows = myIncomeReport.months.map((m) => ({
        月份: m.month, 計入收入嘅堂數: m.count, 學生收費總額: m.gross, 租場費用: m.rentalCost, 實際收入: m.net,
      }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(monthRows.length ? monthRows : [{ 月份: "", 計入收入嘅堂數: "", 學生收費總額: "", 租場費用: "", 實際收入: "" }]), "每月收入");

      const logRows = [];
      Object.entries(myIncomeReport.studentLog).forEach(([name, log]) => {
        log.forEach((l) => logRows.push({ 學生: name, 日期: l.date, 開始: l.start, 時長小時: l.hours, 類型: l.type === "duo" ? "一對二" : "一對一", 收費: l.charge }));
      });
      logRows.sort((a, b) => `${b.日期}${b.開始}`.localeCompare(`${a.日期}${a.開始}`));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(logRows.length ? logRows : [{ 學生: "", 日期: "", 開始: "", 時長小時: "", 類型: "", 收費: "" }]), "學生上課紀錄");

      const today = new Date().toISOString().slice(0, 10);
      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `我的收入_${liveUser.name}_${today}.xlsx`;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1000);
      showToast("已匯出（可直接用 Google Sheets 或 Excel 打開）");
    } catch (e) {
      showToast("匯出失敗，請再試", "error");
    }
  };

  const exportExcel = () => {
    try {
      const wb = XLSX.utils.book_new();

      // 1) 教練總覽
      const coachRows = coaches.map((c) => ({
        教練: c.name, 帳號: c.username, 已購買時數: c.credits, 已用時數: c.used,
        剩餘時數: c.credits - c.used, 代book每堂租金: c.rate,
        總付款: purchaseLog.filter((r) => r.coachId === c.id).reduce((a, r) => a + r.amount, 0)
          + Math.max(0, c.credits - purchaseLog.filter((r) => r.coachId === c.id).reduce((a, r) => a + r.qty, 0)) * c.rate,
      }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(coachRows), "教練總覽");

      // 2) 流水帳（全部）
      const ledgerRows = purchaseLog.map((r) => ({ 日期: r.date, 教練: r.coachName, 增加時數: r.qty, 每小時: r.rate, 金額: r.amount }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ledgerRows.length ? ledgerRows : [{ 日期: "", 教練: "", 增加時數: "", 每小時: "", 金額: "" }]), "流水帳-全部");

      // 3) 每個教練獨立流水帳
      coaches.forEach((c) => {
        const rows = purchaseLog.filter((r) => r.coachId === c.id).map((r) => ({ 日期: r.date, 增加時數: r.qty, 每小時: r.rate, 金額: r.amount }));
        const total = rows.reduce((a, r) => a + r.金額, 0);
        const body = rows.length ? [...rows, { 日期: "小計", 增加時數: "", 每小時: "", 金額: total }] : [{ 日期: "（無記錄）", 增加時數: "", 每小時: "", 金額: "" }];
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
      const cxRows = cancelLog.map((r) => ({ 原定日期: r.date, 開始: r.start, 時長小時: r.hours, 類型: r.type === "charter" ? rentalFull(r.charterType) : r.type === "duo" ? "一對二" : "一對一", 教練: r.coachName || "", 收費: r.price || 0, 取消方式: actorLabel(r.cancelledBy), 取消時間: r.cancelledAt || "" }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cxRows.length ? cxRows : [{ 原定日期: "", 開始: "", 時長小時: "", 類型: "", 教練: "", 收費: "", 取消方式: "", 取消時間: "" }]), "取消記錄");

      const today = new Date().toISOString().slice(0, 10);
      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${BRAND_NAME}_資料備份_${today}.xlsx`;
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
    const header = ["日期", "教練", "增加時數", "每小時", "金額"];
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
        myBookings.push({ date, start: v.start, hours: v.hours, type: v.type, charterType: v.charterType || null, coachName: v.coachName || "", price: v.price || 0, rentalCost: v.rentalCost ?? (v.price || 0), students: v.students || [], studentCharges: v.studentCharges || {}, signatures: v.signatures || {} });
    });
  });
  myBookings.sort((a, b) => `${b.date}${b.start}`.localeCompare(`${a.date}${a.start}`));
  // 「距今日最近」排序：未來（包括今日）優先，由近到遠；之後先至到過去嘅日子，由最近至最舊
  // 修正前 bug：用 Math.abs() 計距離令過去同未來平等，導致已完成嘅過去日子跑咗上最頂
  const myBookingsSorted = (() => {
    if (myBookingsSortMode !== "closest") return myBookings;
    const todayMs = new Date(`${formatDate(new Date())}T00:00:00`).getTime();
    const dayMs = (d) => new Date(`${d}T00:00:00`).getTime();
    const future = myBookings.filter((b) => dayMs(b.date) >= todayMs).sort((a, b) => dayMs(a.date) - dayMs(b.date));
    const past = myBookings.filter((b) => dayMs(b.date) < todayMs).sort((a, b) => dayMs(b.date) - dayMs(a.date));
    return [...future, ...past];
  })();

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
          <img src={LOGO} alt={BRAND_NAME} style={S.loginLogoImg} />
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
          allBookings.push({ date, start: v.start, hours: v.hours, type: v.type, charterType: v.charterType, price: v.price || 0, coachName: v.coachName || "", coach: v.type === "charter" ? null : getCoach(v.coachId), coachId: v.coachId, createdAt: v.createdAt || null, students: v.students || [], signatures: v.signatures || {} });
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

    // 指定月份已用時數／已購時數（同「本月總收入」用返同一個時間範圍，等成行 KPI 對得上數）
    const monthUsed = allBookings.filter((b) => b.type !== "charter" && monthKey(b.date) === viewMonth).reduce((s, b) => s + b.hours, 0);
    const monthSold = purchaseLog.filter((r) => monthKey(r.date) === viewMonth).reduce((s, r) => s + r.qty, 0);

    // 各教練總付款（買堂 + 初始）
    const coachPaid = {};
    coaches.forEach((c) => { coachPaid[c.id] = purchaseLog.filter((r) => r.coachId === c.id).reduce((a, r) => a + r.amount, 0) + initialCreditsOf(c) * c.rate; });

    const isSubAdmin = currentUser.role === "subadmin";
    const visibleTabs = [["overview", "📊", "總覽"], ["schedule", "📅", "課表"], ["coaches", "👥", "教練"], ["ledger", "💰", "流水帳"], ["records", "📋", "記錄"], ["settings", "⚙️", "設定"]]
      .filter(([k]) => !isSubAdmin || currentUser.permissions?.[k]);
    return (
      <div style={S.appBg}>
        <Header title={isSubAdmin ? `副管理員 · ${currentUser.name}` : "管理員"} onLogout={logout} syncState={syncState} />
        {venueNotice && venueNotice.trim() && <div style={S.noticeBanner}>📢 {venueNotice}（教練都見到呢條公告）</div>}
        <div style={S.tabRow}>
          {visibleTabs.map(([k, icon, label]) => (
            <button key={k} style={adminTab === k ? S.tabActive : S.tab} onClick={() => setAdminTab(k)}><span style={S.tabIcon}>{icon}</span><span>{label}</span></button>
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
              <div style={S.kpiCard}><div style={S.kpiLabel}>{viewMonth === thisMonth ? "本月已用時數" : "已用時數"}</div><div style={S.kpiBig}>{monthUsed}</div></div>
              <div style={S.kpiCard}><div style={S.kpiLabel}>{viewMonth === thisMonth ? "本月已購時數" : "已購時數"}</div><div style={S.kpiBig}>{monthSold}</div></div>
            </div>
            <p style={S.assistHint}>本月＝{thisMonth}　｜　累計總收入 ${totalRevenue.toLocaleString()}　｜　累計已用時數 {totalUsed}　｜　累計已購時數 {totalSold}</p>

            <div style={{ ...S.flexBetween, marginBottom: 0 }}>
              <h2 style={S.sectionTitle}>每月收入</h2>
              {allMonths.length > 6 && (
                <button style={S.linkBtn} onClick={() => setMonthsExpanded((v) => !v)}>{monthsExpanded ? "收埋" : `顯示全部（${allMonths.length}）`}</button>
              )}
            </div>
            <div style={S.bookingList}>
              {allMonths.length === 0 ? <p style={S.emptyText}>暫無收入</p> : (monthsExpanded ? allMonths : allMonths.slice(0, 6)).map((m) => (
                <div key={m} style={S.monthCard}>
                  <div style={S.monthHead}>{m === "初始" ? "初始已售時數" : m}</div>
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
                <option value="remain">剩餘時數（少→多）</option>
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
                    ⚠️ 時數快用完（剩 ≤ {LOW_CREDIT_THRESHOLD}）：{lowList.map((c) => `${c.name}（剩${c.credits - c.used}）`).join("、")}　— 可提早提醒增購
                  </div>
                )}
                {expiringCoaches.length > 0 && (
                  <div style={{ ...S.lowWarnBox, background: "#332a0f", color: "#FFB347" }}>
                    ⏰ 時數即將／已經過期：{expiringCoaches.map((c) => {
                      const batches = expiringBatchesOf(c.id);
                      return `${c.name}（${batches.map((b) => `${b.remaining}小時@${b.expiryDate}`).join("、")}）`;
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
                            <div style={{ ...S.segRow, marginBottom: 10 }}>
                              <button style={coachDrillView === "purchase" ? S.segActive : S.seg} onClick={() => setCoachDrillView("purchase")}>購堂記錄</button>
                              <button style={coachDrillView === "sessions" ? S.segActive : S.seg} onClick={() => setCoachDrillView("sessions")}>上堂記錄</button>
                            </div>
                            {coachDrillView === "purchase" ? (
                              fifo.length === 0 ? <p style={S.emptyText}>暫無購買記錄</p> : fifo.map((b) => {
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
                              })
                            ) : (() => {
                              // 第6項：呢位教練喺「查看月份」（頂部 viewMonth）入面嘅上堂詳情——日期＋時長＋學生名＋有冇簽到
                              const sessions = allBookings
                                .filter((b) => b.coachId === c.id && monthKey(b.date) === viewMonth)
                                .sort((a, b) => `${a.date}${a.start}`.localeCompare(`${b.date}${b.start}`));
                              return sessions.length === 0 ? <p style={S.emptyText}>{viewMonth} 冇上堂記錄</p> : sessions.map((s, i) => {
                                const names = s.students && s.students.length > 0 ? s.students : ["（未填學生名）"];
                                const signedCount = (s.students || []).filter((n) => s.signatures && s.signatures[n]).length;
                                const totalStudents = (s.students || []).length;
                                const signLabel = totalStudents === 0 ? "" : signedCount === totalStudents ? "✅ 已簽" : signedCount === 0 ? "⚠️ 未簽" : `${signedCount}/${totalStudents} 已簽`;
                                return (
                                  <div key={i} style={S.purchaseRow}>
                                    <div>
                                      <div style={S.bookingTime}>{s.date}（{formatDay(new Date(`${s.date}T00:00:00`))}） {s.start}–{addMinutes(s.start, s.hours * 60)}（{s.hours}小時）</div>
                                      <div style={{ fontSize: 12, color: "#aaa", marginTop: 2 }}>{names.join("、")}　{s.type === "duo" ? "1對2" : "1對1"}</div>
                                    </div>
                                    {signLabel && <div style={{ fontSize: 12, color: signedCount === totalStudents ? "#6BCB77" : signedCount === 0 ? "#FF8FA3" : "#FFB347" }}>{signLabel}</div>}
                                  </div>
                                );
                              });
                            })()}
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
              <button style={S.navBtn} onClick={() => setWeekOffset(0)}>今日</button>
              <button style={S.navBtn} onClick={() => setWeekOffset((w) => w + 1)}>下週 ›</button>
              <button style={S.navBtn} onClick={() => { setWeekViewMode((m) => m === "fixed" ? "rolling" : "fixed"); setWeekOffset(0); }} title="切換週視圖模式">🔁 {weekViewMode === "fixed" ? "一至日" : "今日起"}</button>
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
                                    {relRow === 0 && <button style={S.cancelSlotBtn} onClick={() => setAdminCancelModal({ date, start: whole.start, coachId: whole.coachId, type: "charter" })}>✕</button>}
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
                                  {canAdd && <button style={S.slotAdd} onClick={() => setSlotChoiceModal({ date, time })}>+</button>}
                                </div>
                              ) : closed ? <div style={S.slotClosed} />
                                : isPast ? <div style={S.slotPast} />
                                : <button style={S.slotEmpty} onClick={() => setSlotChoiceModal({ date, time })}>+</button>}
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
                    <div style={S.bookingTime}>Pass時數 {c.used}/{c.credits} 小時　代book每堂 ${c.rate}　密碼 {showPasswords ? c.password : "••••"}</div>
                    {(() => {
                      const os = c.onboardingStatus || {};
                      const done = ["payment", "welcome", "rental", "terms"].filter((k) => os[k]).length;
                      return done > 0 && done < 4 ? <div style={{ fontSize: 11, color: "#FFB347" }}>Onboarding {done}/4</div> : done === 4 ? <div style={{ fontSize: 11, color: "#6BCB77" }}>Onboarding 完成 ✓</div> : null;
                    })()}
                  </div>
                  <button style={S.creditBtn} onClick={() => setAddCreditModal({ coachId: c.id, qty: 1, date: formatDate(new Date()), expiryDate: "", passType: "" })}>+ 堂</button>
                  <button style={S.smallBtn} onClick={() => setRetroReminderModal({ coachId: c.id, coachName: c.name, date: formatDate(new Date()), start: "19:00", hours: 1 })}>⚠️ 提醒補book</button>
                  <button style={S.smallBtn} onClick={() => setEditCoach(c)}>編輯</button>
                  <button style={S.delBtn} onClick={() => setDelCoachModal(c)}>刪</button>
                </div>
              ))}
            </div>

            <div style={{ ...S.flexBetween, marginTop: 24 }}>
              <h2 style={{ ...S.sectionTitle, marginBottom: 0 }}>共享訓練通行證</h2>
              <button style={S.addBtn} disabled={coaches.length < 2} onClick={() => setSharedPassModal({ coachIdA: coaches[0]?.id, coachIdB: coaches[1]?.id, date: formatDate(new Date()) })}>+ 開共享 Pass</button>
            </div>
            {coaches.length < 2 && <p style={S.assistHint}>要至少2位教練先可以開共享 Pass。</p>}
            {sharedPasses.length === 0 ? <p style={S.emptyText}>暫無共享 Pass</p> : (
              <div style={S.bookingList}>
                {sharedPasses.map((sp) => {
                  const names = (sp.coachIds || []).map((id) => getCoach(id)?.name || "（已刪除教練）");
                  const remain = sharedRemaining(sp);
                  return (
                    <div key={sp.id} style={S.purchaseRow}>
                      <div>
                        <div style={S.bookingCoach}>{names.join(" ＋ ")}</div>
                        <div style={S.bookingTime}>{sp.purchaseDate}　失效日：{sp.expiryDate || "無限期"}</div>
                        <div style={{ fontSize: 11, color: "#888", marginTop: 2 }}>{(sp.coachIds || []).map((id) => `${getCoach(id)?.name || id}用咗${(sp.usageByCoach || {})[id] || 0}hr`).join("　")}</div>
                      </div>
                      <div style={{ textAlign: "right", fontSize: 12 }}>
                        <div style={{ color: "#6BCB77" }}>已用 {sp.usedHours || 0}</div>
                        <div style={{ color: remain > 0 ? "#4ECDC4" : "#555" }}>剩 {remain} / {sp.totalHours}</div>
                        <button style={S.linkBtn} onClick={() => setSharedTopUpModal({ sharedId: sp.id, qty: 5 })}>+ 加值</button>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}

        {adminTab === "ledger" && (() => {
          let filtered = ledgerFilter === "all" ? purchaseLog : purchaseLog.filter((r) => String(r.coachId) === String(ledgerFilter));
          if (ledgerMonth !== "all") filtered = filtered.filter((r) => monthKey(r.date) === ledgerMonth);
          const filteredTotal = filtered.reduce((s, r) => s + r.amount, 0);
          return (
          <div style={S.container}>
            <h2 style={S.sectionTitle}>購買時數流水帳</h2>
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
                      <div style={S.bookingTime}>{r.date}　@${r.rate}/堂{r.addedBy ? `　由${r.addedBy === "admin" ? "管理員" : r.addedBy.startsWith("subadmin:") ? `副管理員（${r.addedBy.slice(9)}）` : r.addedBy}新增` : ""}</div>
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
                    <option value="clean">清潔</option>
                    <option value="filming">拍片</option>
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
                            {type !== "charter" && <div>扣時數：{hours} 小時</div>}
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
                          <span style={S.cancelledTag}>{actorLabel(r.cancelledBy)}</span>
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
            {currentUser.role === "admin" && (
              <div style={{ display: "flex", gap: 6, overflowX: "auto", padding: "2px 2px 10px", marginBottom: 6, position: "sticky", top: 0, background: "#0f0f0f", zIndex: 30 }}>
                {[["set-password", "密碼"], ["set-notice", "場地公告"], ["set-qr-account", "收款QR"], ["set-drinks", "飲品"], ["set-calendar", "日曆同步"], ["set-whatsapp", "WhatsApp"], ["set-suggestions", "意見箱"], ["set-export", "備份"], ["set-subadmins", "副管理員"], ["set-reset", "重設資料"]].map(([id, label]) => (
                  <button key={id} style={{ ...S.smallBtn, whiteSpace: "nowrap", flexShrink: 0 }}
                    onClick={() => document.getElementById(id)?.scrollIntoView({ behavior: "smooth", block: "start" })}>{label}</button>
                ))}
              </div>
            )}
            <h2 id="set-password" style={S.sectionTitle}>修改{isSubAdmin ? "我的" : "管理員"}密碼</h2>
            <div style={S.formCard}>
              <Field label="舊密碼"><input style={S.input} type="password" value={pwForm.old} onChange={(e) => setPwForm({ ...pwForm, old: e.target.value })} /></Field>
              <Field label="新密碼"><input style={S.input} type="password" value={pwForm.new1} onChange={(e) => setPwForm({ ...pwForm, new1: e.target.value })} /></Field>
              <Field label="確認新密碼"><input style={S.input} type="password" value={pwForm.new2} onChange={(e) => setPwForm({ ...pwForm, new2: e.target.value })} /></Field>
              <button style={S.loginBtn} onClick={changePassword}>更新密碼</button>
            </div>
            <p style={S.assistHint}>※ 教練自己 book 堂：Training Pass 制，$100/小時（1對2 額外多扣0.5小時）。Admin 代教練 book 堂：維持單次收費（1對1 用教練每堂租金；1對2 $150/小時，每加0.5小時 +$50）。</p>

            {currentUser.role === "admin" && (
              <>
                <h2 id="set-notice" style={{ ...S.sectionTitle, marginTop: 28 }}>場地公告</h2>
                <div style={S.formCard}>
                  <p style={{ ...S.bookingTime, marginBottom: 14, lineHeight: 1.6 }}>例如「本週洗手間維修，請使用更衣室」。所有教練登入會見到呢條提示。留空就唔顯示。</p>
                  <Field label="公告內容"><textarea style={{ ...S.input, minHeight: 70, resize: "vertical" }} value={venueNotice} onChange={(e) => setVenueNotice(e.target.value)} placeholder="留空＝唔顯示" /></Field>
                </div>

                <h2 id="set-qr-account" style={{ ...S.sectionTitle, marginTop: 28 }}>收款 QR Code</h2>
                <div style={S.formCard}>
                  <p style={{ ...S.bookingTime, marginBottom: 14, lineHeight: 1.6 }}>上傳收款 QR code（例如轉數快／PayMe）。會喺教練 Onboarding 付款資訊同買 Pass 畫面顯示，方便教練掃碼付款。可隨時更換。</p>
                  {paymentQR ? (
                    <div style={{ textAlign: "center", marginBottom: 12 }}>
                      <img src={paymentQR} alt="收款 QR" style={{ maxWidth: 200, width: "100%", borderRadius: 10, background: "#fff", padding: 8, boxSizing: "border-box" }} />
                    </div>
                  ) : <p style={S.emptyText}>仲未上傳</p>}
                  <label style={{ ...S.loginBtn, display: "block", textAlign: "center", cursor: "pointer" }}>
                    {paymentQR ? "更換 QR Code" : "上傳 QR Code"}
                    <input type="file" accept="image/*" style={{ display: "none" }} onChange={(e) => handleQRUpload(e.target.files?.[0])} />
                  </label>
                  {paymentQR && <button style={{ ...S.smallBtn, width: "100%", marginTop: 8 }} onClick={() => setPaymentQR("")}>移除</button>}
                </div>

                <h2 id="set-drinks" style={{ ...S.sectionTitle, marginTop: 28 }}>飲品產品管理</h2>
                <div style={S.formCard}>
                  <p style={{ ...S.bookingTime, marginBottom: 14, lineHeight: 1.6 }}>教練喺「其他」分頁可以幫學生落單，撳掣顯示上面嘅收款QR。呢度管理有邊啲產品同價錢。</p>
                  {drinkProducts.length === 0 ? (
                    <p style={S.emptyText}>仲未上架任何產品</p>
                  ) : (
                    <div style={{ marginBottom: 14 }}>
                      {drinkProducts.map((p) => (
                        <div key={p.id} style={{ display: "flex", gap: 8, alignItems: "center", padding: "8px 0", borderBottom: "1px solid #222" }}>
                          <input style={{ ...S.input, flex: 2 }} value={p.name} onChange={(e) => updateDrinkProduct(p.id, "name", e.target.value)} />
                          <div style={{ display: "flex", alignItems: "center", flex: 1 }}>
                            <span style={{ marginRight: 4 }}>$</span>
                            <input style={S.input} type="number" min="0" value={p.price} onChange={(e) => updateDrinkProduct(p.id, "price", Number(e.target.value) || 0)} />
                          </div>
                          <button style={S.smallBtn} onClick={() => removeDrinkProduct(p.id)}>刪除</button>
                        </div>
                      ))}
                    </div>
                  )}
                  <label style={S.label}>新增產品</label>
                  <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
                    <input style={{ ...S.input, flex: 2 }} placeholder="產品名（例如：樽裝水）" value={newDrinkForm.name} onChange={(e) => setNewDrinkForm({ ...newDrinkForm, name: e.target.value })} />
                    <input style={{ ...S.input, flex: 1 }} type="number" min="0" placeholder="$" value={newDrinkForm.price} onChange={(e) => setNewDrinkForm({ ...newDrinkForm, price: e.target.value })} />
                  </div>
                  <button style={{ ...S.loginBtn, marginTop: 10 }} onClick={addDrinkProduct}>新增</button>
                </div>

                <h2 id="set-calendar" style={{ ...S.sectionTitle, marginTop: 28 }}>同步落自己嘅日曆</h2>
                <div style={S.formCard}>
                  {!cloudEnabled ? (
                    <p style={{ ...S.bookingTime, lineHeight: 1.6 }}>呢個功能需要先開啟雲端同步。</p>
                  ) : (
                    <>
                      <p style={{ ...S.bookingTime, marginBottom: 10, lineHeight: 1.6 }}>生成連結加入 Google／Apple Calendar,顯示全場所有教練嘅 booking(連教練名)。</p>
                      {adminCalendarFeedUrl ? (
                        <>
                          <input style={{ ...S.input, fontSize: 11 }} readOnly value={adminCalendarFeedUrl} onFocus={(e) => e.target.select()} />
                          <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
                            <button style={{ ...S.creditBtn, flex: 1 }} onClick={async () => { try { await navigator.clipboard.writeText(adminCalendarFeedUrl); showToast("已複製連結"); } catch (e) { showToast("複製失敗，請長按手動複製", "error"); } }}>📋 複製連結</button>
                            <button style={{ ...S.smallBtn, flex: 1 }} onClick={regenerateAdminCalendarToken}>🔄 重新生成</button>
                          </div>
                          <p style={{ ...S.assistHint, marginTop: 8 }}>連結等於密碼，請唔好分享畀其他人。</p>
                        </>
                      ) : (
                        <button style={S.loginBtn} onClick={ensureAdminCalendarToken}>生成同步連結</button>
                      )}
                    </>
                  )}
                </div>

                <h2 id="set-whatsapp" style={{ ...S.sectionTitle, marginTop: 28 }}>場地 QR Code WhatsApp 號碼</h2>
                <div style={S.formCard}>
                  <p style={{ ...S.bookingTime, marginBottom: 14, lineHeight: 1.6 }}>教練喺「我的預約」撳「攞 QR Code」會自動開 WhatsApp 傳訊息去呢個號碼。請輸入完整國際格式（例如香港：85291234567，唔使 + 號）。</p>
                  <Field label="WhatsApp 號碼"><input style={S.input} placeholder="例如 85291234567" value={whatsappNumber} onChange={(e) => setWhatsappNumber(e.target.value.replace(/[^0-9]/g, ""))} /></Field>
                </div>

                <h2 style={{ ...S.sectionTitle, marginTop: 28 }}>收款 QR Code（轉數快／PayMe）</h2>
                <div style={S.formCard}>
                  <p style={{ ...S.bookingTime, marginBottom: 14, lineHeight: 1.6 }}>上傳一張收款 QR Code 圖，將來會顯示喺 Onboarding 付款資訊、教練買 Pass 畫面等地方，方便教練直接掃碼轉數。可隨時重新上傳更換。</p>
                  {paymentQR && <img src={paymentQR} alt="收款 QR Code" style={{ width: 160, height: 160, objectFit: "contain", background: "#fff", borderRadius: 10, marginBottom: 10 }} />}
                  <input type="file" accept="image/*" onChange={(e) => {
                    const file = e.target.files?.[0];
                    if (!file) return;
                    const reader = new FileReader();
                    reader.onload = () => { setPaymentQR(reader.result); showToast("已上傳收款 QR Code"); };
                    reader.onerror = () => showToast("上傳失敗，請再試一次", "error");
                    reader.readAsDataURL(file);
                    e.target.value = "";
                  }} />
                  {paymentQR && <button style={{ ...S.smallBtn, marginTop: 10 }} onClick={() => { setPaymentQR(""); showToast("已移除收款 QR Code"); }}>移除</button>}
                </div>

                <h2 id="set-suggestions" style={{ ...S.sectionTitle, marginTop: 28 }}>匿名改善建議（只有你睇到）</h2>
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

            <h2 id="set-export" style={{ ...S.sectionTitle, marginTop: 28 }}>匯出資料備份</h2>
            <div style={S.formCard}>
              <p style={{ ...S.bookingTime, marginBottom: 14, lineHeight: 1.6 }}>匯出檔案可直接用 Google Sheets 或 Excel 打開，包含教練總覽、全部流水帳、每個教練獨立流水帳、上堂記錄、包場小組記錄、取消記錄。建議定期備份。</p>
              <button style={{ ...S.loginBtn, background: "#6BCB77" }} onClick={exportExcel}>📊 匯出 Google Sheet 備份</button>
              <button style={{ ...S.loginBtn, background: "#2a2a2a", color: "#fff", marginTop: 10 }} onClick={copyLedgerCsv}>📋 複製流水帳 (CSV)</button>
              <p style={{ ...S.assistHint, marginTop: 10 }}>※ 若下載冇反應（手機 app 常見），可改按「複製流水帳」再貼入 Google Sheets / Excel；或喺電腦瀏覽器開啟再匯出。</p>
            </div>

            {currentUser.role === "admin" && (
              <>
                <h2 id="set-subadmins" style={{ ...S.sectionTitle, marginTop: 28 }}>副管理員帳戶</h2>
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
                <h2 id="set-reset" style={{ ...S.sectionTitle, marginTop: 28 }}>重設資料</h2>
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
            onOnboardingSend={(stepKey, coachName, initialPassHours, phone) => {
              let text = "";
              if (stepKey === "fee") text = onboardingFeeSheetText();
              else if (stepKey === "guide") text = onboardingVenueRulesText();
              else if (stepKey === "payment") {
                const hrs = initialPassHours === "" || initialPassHours === null || initialPassHours === undefined ? NaN : Number(initialPassHours);
                if (!hrs || hrs <= 0) { showToast("請先輸入初始 Pass 時數，先會生成付款資訊文字", "error"); return false; }
                text = onboardingPaymentInfoText(hrs);
              }
              else if (stepKey === "welcome") text = onboardingWelcomeText(coachName);
              else if (stepKey === "rental") text = onboardingRentalGuideText();
              else if (stepKey === "terms") text = onboardingTermsText();
              const stepLabels = { fee: "收費表", guide: "場地守則", payment: "付款資訊", welcome: "歡迎訊息", rental: "租場須知", terms: "使用條款" };
              if (phone) {
                window.open(`https://wa.me/${phone}?text=${encodeURIComponent(text)}`, "_blank");
              } else {
                setCopyInfoModal({ title: `Onboarding：${stepLabels[stepKey] || ""}`, text });
              }
              const trackedKeys = ["payment", "welcome", "rental", "terms"];
              if (trackedKeys.includes(stepKey) && editCoach?.id) {
                setCoaches((prev) => prev.map((c) => c.id === editCoach.id ? { ...c, onboardingStatus: { ...(c.onboardingStatus || {}), [stepKey]: true } } : c));
              }
              return true;
            }}
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
                const color = coachColorFromId(newId);
                const initials = clean.name.trim().split(/\s+/).map((w) => w[0]).join("").slice(0, 2).toUpperCase() || "NA";
                setCoaches((prev) => [...prev, { ...clean, id: newId, color, initials, used: 0 }]);
                if (clean.credits > 0) setPurchaseLog((prev) => [{ id: "p" + Date.now() + "-" + Math.random().toString(36).slice(2), date: new Date().toISOString().slice(0, 10), coachId: newId, coachName: clean.name, qty: clean.credits, amount: clean.credits * clean.rate, rate: clean.rate, addedBy: currentUser.role === "subadmin" ? `subadmin:${currentUser.name}` : "admin" }, ...prev]);
                showToast(`已新增教練 ${clean.name}（@${uname}）`);
              }
              setEditCoach(null);
            }} />
        )}

        {addCreditModal && (
          <div style={S.modalOverlay}><div style={S.modal}>
            <h3 style={S.modalTitle}>增加時數</h3>
            <p style={S.modalText}>{getCoach(addCreditModal.coachId)?.name}</p>
            <label style={S.label}>Pass 類型</label>
            <div style={{ ...S.segRow, flexWrap: "wrap" }}>
              <button style={addCreditModal.passType === "personal" ? S.segActive : S.seg}
                onClick={() => setAddCreditModal({ ...addCreditModal, passType: "personal", qty: PERSONAL_PASS_HOURS, date: addCreditModal.date || formatDate(new Date()), expiryDate: addMonthsToDate(addCreditModal.date || formatDate(new Date()), PERSONAL_PASS_MONTHS) })}>
                個人證（{PERSONAL_PASS_HOURS}hr／solo限定）
              </button>
              <button style={addCreditModal.passType === "flexible" ? S.segActive : S.seg}
                onClick={() => setAddCreditModal({ ...addCreditModal, passType: "flexible", qty: FLEXIBLE_PASS_HOURS, date: addCreditModal.date || formatDate(new Date()), expiryDate: addMonthsToDate(addCreditModal.date || formatDate(new Date()), FLEXIBLE_PASS_MONTHS) })}>
                彈性證（{FLEXIBLE_PASS_HOURS}hr／全類型）
              </button>
              <button style={!addCreditModal.passType ? S.segActive : S.seg}
                onClick={() => setAddCreditModal({ ...addCreditModal, passType: "", qty: 1 })}>
                自訂（不限類型）
              </button>
            </div>
            {addCreditModal.passType === "personal" && <p style={S.assistHint}>個人訓練通行證：淨係可以用嚟 book 1對1，唔可以book 1對2。</p>}
            {addCreditModal.passType === "flexible" && <p style={S.assistHint}>彈性訓練通行證：1對1／1對2／包場都用得。</p>}
            <Field label="小時數（可 0.5 為一格）"><input style={S.input} type="number" step="0.5" min="0.5" value={addCreditModal.qty} onChange={(e) => setAddCreditModal({ ...addCreditModal, qty: e.target.value })} /></Field>
            <Field label="增加日期"><input style={S.input} type="date" value={addCreditModal.date || formatDate(new Date())} onChange={(e) => setAddCreditModal({ ...addCreditModal, date: e.target.value })} /></Field>
            <Field label="失效日期（留空＝無限期；過咗期都仍然可以用，只係會提示教練）"><input style={S.input} type="date" value={addCreditModal.expiryDate || ""} onChange={(e) => setAddCreditModal({ ...addCreditModal, expiryDate: e.target.value })} /></Field>
            {addCreditModal.passType ? (
              <p style={S.amountPreview}>Pass 制：買咗之後一律 ${PASS_HOURLY_RATE}/小時計算（book 堂嗰陣先收，呢度唔預收費用）</p>
            ) : (
              <p style={S.amountPreview}>金額：${((getCoach(addCreditModal.coachId)?.rate || 0) * (Number(addCreditModal.qty) || 0)).toLocaleString()}</p>
            )}
            <div style={S.modalBtns}>
              <button style={S.modalCancel} onClick={() => setAddCreditModal(null)}>取消</button>
              <button style={S.modalConfirm} onClick={() => { const qty = Number(addCreditModal.qty) || 0; if (qty <= 0) { showToast("請輸入有效時數", "error"); return; } addCredits(addCreditModal.coachId, qty, addCreditModal.date, addCreditModal.expiryDate, addCreditModal.passType || null); setAddCreditModal(null); }}>確認增加</button>
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
              <button style={charterModal.charterType === "private" ? S.segActive : S.seg} onClick={() => setCharterModal({ ...charterModal, charterType: "private", price: ["trial", "clean"].includes(charterModal.charterType) ? CHARTER_PRICE : charterModal.price })}>私人包場</button>
              <button style={charterModal.charterType === "group" ? S.segActive : S.seg} onClick={() => setCharterModal({ ...charterModal, charterType: "group", price: ["trial", "clean"].includes(charterModal.charterType) ? CHARTER_PRICE : charterModal.price })}>小組訓練</button>
              <button style={charterModal.charterType === "trial" ? S.segActive : S.seg} onClick={() => setCharterModal({ ...charterModal, charterType: "trial", price: 0 })}>試堂</button>
              <button style={charterModal.charterType === "clean" ? S.segActive : S.seg} onClick={() => setCharterModal({ ...charterModal, charterType: "clean", price: 0 })}>🧹 清潔</button>
            </div>
            {charterModal.charterType === "trial"
              ? <p style={{ ...S.assistHint, marginTop: 6 }}>試堂只佔 1 個位，同一時段仲可以有教練 book，唔收費。</p>
              : charterModal.charterType === "clean"
              ? <p style={{ ...S.assistHint, marginTop: 6 }}>封場清潔，獨佔全場（2 位），$0，全部教練都見到呢格係「清潔」，唔可以book。</p>
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
                  onChange={(e) => setCharterModal({ ...charterModal, hours: e.target.value })} />
              </div>
            )}

            <label style={{ ...S.label, marginTop: 14 }}>負責教練</label>
            <select style={{ ...S.select, width: "100%", boxSizing: "border-box" }} value={charterModal.coachName} onChange={(e) => setCharterModal({ ...charterModal, coachName: e.target.value })}>
              <option value="">（未指定）</option>
              {coaches.map((c) => <option key={c.id} value={c.name}>{c.name}</option>)}
            </select>

            {charterModal.charterType === "clean" && (
              <>
                <label style={{ ...S.label, marginTop: 14 }}>每週重複（同一星期幾、同一時間）</label>
                <div style={S.segRow}>
                  {[1, 4, 8, 12].map((w) => (
                    <button key={w} style={(charterModal.repeatWeeks || 1) === w ? S.segActive : S.seg} onClick={() => setCharterModal({ ...charterModal, repeatWeeks: w })}>{w === 1 ? "唔重複" : `${w}週`}</button>
                  ))}
                </div>
                {(charterModal.repeatWeeks || 1) > 1 && <p style={S.assistHint}>會一次過幫你book未來 {charterModal.repeatWeeks} 個星期嘅同一個清潔時段；如果某一週已經被佔用，會自動跳過嗰一週，唔影響其他週。</p>}
              </>
            )}
            {["trial", "clean"].includes(charterModal.charterType) ? (
              <p style={{ ...S.amountPreview, color: "#999", marginTop: 14 }}>{charterModal.charterType === "clean" ? "封場清潔不收費，唔會計入收入。" : "試堂不收費，唔會計入收入。"}</p>
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
              <div style={{ ...S.priceRow, color: "#4ECDC4", fontWeight: 700, fontSize: 16 }}><span>收費</span><span>{["trial", "clean"].includes(charterModal.charterType) ? "免費" : `$${parseInt(charterModal.price) || 0}`}</span></div>
            </div>
            <div style={S.modalBtns}>
              <button style={S.modalCancel} onClick={() => setCharterModal(null)}>返回</button>
              <button style={S.modalConfirm} onClick={confirmCharter}>確認落單</button>
            </div>
          </div></div>
        )}

        {slotChoiceModal && (
          <div style={S.modalOverlay}><div style={S.modal}>
            <h3 style={S.modalTitle}>呢個時段想點落單？</h3>
            <p style={S.modalText}>{slotChoiceModal.date}　{slotChoiceModal.time}</p>
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              <button style={S.loginBtn} onClick={() => { const { date, time } = slotChoiceModal; setSlotChoiceModal(null); setAdminCoachBookModal({ date, time, coachId: coaches[0]?.id || null, sessionType: "solo", hours: 1, students: [] }); }}>👤 代教練 Book 堂</button>
              <button style={{ ...S.loginBtn, background: "#2a2a2a", color: "#fff" }} onClick={() => { const { date, time } = slotChoiceModal; setSlotChoiceModal(null); setCharterModal({ date, time, charterType: "private", hours: 1, price: CHARTER_PRICE, coachName: "" }); }}>🏟️ 包場／小組／試堂</button>
            </div>
            <button style={{ ...S.modalCancel, marginTop: 14, width: "100%" }} onClick={() => setSlotChoiceModal(null)}>取消</button>
          </div></div>
        )}

        {adminCoachBookModal && (() => {
          const selCoach = getCoach(adminCoachBookModal.coachId);
          const roster = selCoach ? getStudentRoster(selCoach.id) : [];
          const isDuo = adminCoachBookModal.sessionType === "duo";
          const price = selCoach ? (isDuo ? duoPrice(adminCoachBookModal.hours) : selCoach.rate * adminCoachBookModal.hours) : 0;
          const creditCostPreview = isDuo ? adminCoachBookModal.hours + 0.5 : adminCoachBookModal.hours;
          const remain = selCoach ? selCoach.credits - selCoach.used : 0;
          return (
            <div style={S.modalOverlay}><div style={{ ...S.modal, textAlign: "left" }}>
              <h3 style={{ ...S.modalTitle, textAlign: "center" }}>代教練 Book 堂</h3>
              <p style={{ ...S.modalText, textAlign: "center" }}>{adminCoachBookModal.date}　{adminCoachBookModal.time}</p>

              <label style={S.label}>教練</label>
              <select style={{ ...S.select, width: "100%", boxSizing: "border-box" }} value={adminCoachBookModal.coachId || ""}
                onChange={(e) => setAdminCoachBookModal({ ...adminCoachBookModal, coachId: parseInt(e.target.value), students: [] })}>
                {coaches.map((c) => <option key={c.id} value={c.id}>{c.name}（剩 {c.credits - c.used} 小時）</option>)}
              </select>

              <label style={{ ...S.label, marginTop: 14 }}>類型</label>
              <div style={S.segRow}>
                <button style={!isDuo ? S.segActive : S.seg} onClick={() => setAdminCoachBookModal({ ...adminCoachBookModal, sessionType: "solo" })}>1對1</button>
                <button style={isDuo ? S.segActive : S.seg} onClick={() => setAdminCoachBookModal({ ...adminCoachBookModal, sessionType: "duo" })}>1對2</button>
              </div>

              <label style={{ ...S.label, marginTop: 14 }}>時長</label>
              <div style={S.segRow}>
                <button style={adminCoachBookModal.hours === 1 ? S.segActive : S.seg} onClick={() => setAdminCoachBookModal({ ...adminCoachBookModal, hours: 1 })}>1 小時</button>
                <button style={adminCoachBookModal.hours === 1.5 ? S.segActive : S.seg} onClick={() => setAdminCoachBookModal({ ...adminCoachBookModal, hours: 1.5 })}>1.5 小時</button>
              </div>

              {roster.length > 0 && (
                <>
                  <label style={{ ...S.label, marginTop: 14 }}>學生（最多4位）</label>
                  <div style={S.studentChipWrap}>
                    {roster.map(({ name }) => {
                      const sel = adminCoachBookModal.students.includes(name);
                      const atMax = !sel && adminCoachBookModal.students.length >= 4;
                      return (
                        <button key={name} disabled={atMax} style={sel ? S.studentChipActive : atMax ? S.studentChipDisabled : S.studentChip}
                          onClick={() => setAdminCoachBookModal({ ...adminCoachBookModal, students: sel ? adminCoachBookModal.students.filter((n) => n !== name) : [...adminCoachBookModal.students, name] })}>{name}</button>
                      );
                    })}
                  </div>
                </>
              )}

              <div style={S.priceBox}>
                <div style={S.priceRow}><span>時段</span><span>{adminCoachBookModal.time} – {addMinutes(adminCoachBookModal.time, adminCoachBookModal.hours * 60)}</span></div>
                <div style={S.priceRow}><span>扣時數</span><span>{creditCostPreview} 小時{isDuo ? "（1對2額外+0.5）" : ""}（{selCoach?.name} 剩 {remain} 小時）</span></div>
                <div style={{ ...S.priceRow, color: "#4ECDC4", fontWeight: 700, fontSize: 16 }}><span>收費</span><span>${price}</span></div>
              </div>
              <div style={S.modalBtns}>
                <button style={S.modalCancel} onClick={() => setAdminCoachBookModal(null)}>返回</button>
                <button style={S.modalConfirm} onClick={confirmAdminCoachBooking}>確認預約</button>
              </div>
            </div></div>
          );
        })()}

        {copyInfoModal && (
          <div style={S.modalOverlay}><div style={S.modal}>
            <h3 style={S.modalTitle}>{copyInfoModal.title || "預約成功"}</h3>
            <p style={S.modalText}>複製落面文字，自行 send 畀教練（系統冇存教練電話，唔會自動發送）</p>
            <textarea style={{ ...S.input, minHeight: 100, resize: "vertical", textAlign: "left" }} readOnly value={copyInfoModal.text} onFocus={(e) => e.target.select()} />
            <div style={S.modalBtns}>
              <button style={S.modalCancel} onClick={() => setCopyInfoModal(null)}>關閉</button>
              <button style={S.modalConfirm} onClick={async () => { try { await navigator.clipboard.writeText(copyInfoModal.text); showToast("已複製"); } catch (e) { showToast("複製失敗，請長按手動複製", "error"); } }}>📋 複製</button>
            </div>
          </div></div>
        )}

        {retroReminderModal && (
          <div style={S.modalOverlay}><div style={S.modal}>
            <h3 style={S.modalTitle}>提醒 {retroReminderModal.coachName} 補book</h3>
            <p style={S.modalText}>揀返教練實際已上堂但未book返嘅時段，會發送提醒（有電話直接開WhatsApp，冇電話畀你copy），同時喺教練首頁出現banner。</p>
            <Field label="日期"><input style={S.input} type="date" value={retroReminderModal.date} onChange={(e) => setRetroReminderModal({ ...retroReminderModal, date: e.target.value })} /></Field>
            <Field label="開始時間"><input style={S.input} type="time" value={retroReminderModal.start} onChange={(e) => setRetroReminderModal({ ...retroReminderModal, start: e.target.value })} /></Field>
            <Field label="時長（小時）"><input style={S.input} type="number" step="0.5" min="0.5" value={retroReminderModal.hours} onChange={(e) => setRetroReminderModal({ ...retroReminderModal, hours: e.target.value })} /></Field>
            <div style={S.modalBtns}>
              <button style={S.modalCancel} onClick={() => setRetroReminderModal(null)}>取消</button>
              <button style={S.modalConfirm} onClick={sendRetroReminder}>發送提醒</button>
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
            <p style={S.modalText}>{adminCancelModal.date}　{adminCancelModal.start}<br />確定幫呢個時段取消？{adminCancelModal.type !== "charter" ? "（會退回對應時數）" : ""}</p>
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
            <p style={S.modalText}>{delLedgerModal.coachName}　+{delLedgerModal.qty} 小時　${delLedgerModal.amount.toLocaleString()}<br />（{delLedgerModal.date}）<br /><br />確定剷除？教練時數會相應扣減，此動作無法復原。</p>
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
                setCoaches(DEFAULT_COACHES); setAdminPassword(DEFAULT_ADMIN_PASSWORD); setBookings({});
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
      {soldOut && <div style={S.soldOutBanner}>⚠️ 你已用晒購買時數，請聯絡管理員增購後再預約</div>}
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
        <button style={view === "calendar" ? S.tabActive : S.tab} onClick={() => setView("calendar")}><span style={S.tabIcon}>📅</span><span>預約場地</span></button>
        <button style={view === "myBookings" ? S.tabActive : S.tab} onClick={() => setView("myBookings")}><span style={S.tabIcon}>📋</span><span>我的預約</span>{myBookings.length > 0 && <span style={S.badge}>{myBookings.length}</span>}</button>
        <button style={view === "home" ? S.tabActive : S.tab} onClick={() => setView("home")}>
          <span style={{ position: "relative" }}>
            <span style={S.tabIcon}>🏠</span>
            {isCoach && (retroBookingNotices.some((n) => n.coachId === currentUser.id && !n.read) || filmingNotices.some((n) => n.coachId === currentUser.id && !n.read)) && (
              <span style={{ position: "absolute", top: -2, right: -4, width: 8, height: 8, borderRadius: "50%", background: "#FFB347" }} />
            )}
          </span>
          <span>首頁</span>
        </button>
        <button style={view === "income" ? S.tabActive : S.tab} onClick={() => setView("income")}><span style={S.tabIcon}>👥</span><span>學生管理</span></button>
        <button style={view === "other" ? S.tabActive : S.tab} onClick={() => setView("other")}><span style={S.tabIcon}>⚙️</span><span>其他</span></button>
      </div>

      {view === "home" && (
        <div style={S.calContainer}>
          {isCoach && (() => {
            const myRetroNotices = retroBookingNotices.filter((n) => n.coachId === currentUser.id && !n.read);
            const myFilmingNotices = filmingNotices.filter((n) => n.coachId === currentUser.id && !n.read);
            if (myRetroNotices.length === 0 && myFilmingNotices.length === 0) return null;
            return (
              <>
                {myRetroNotices.map((n) => (
                  <div key={n.id} style={S.noticeBanner}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                      <span>⚠️ {n.date} {n.start}–{addMinutes(n.start, n.hours * 60)} 未book返記錄</span>
                      <div style={{ display: "flex", gap: 8 }}>
                        <button style={S.linkBtn} onClick={() => setRetroBookModal({ noticeId: n.id, date: n.date, start: n.start, hours: n.hours, sessionType: "solo", students: [] })}>處理</button>
                        <button style={S.linkBtn} onClick={() => setRetroBookingNotices((prev) => prev.map((x) => x.id === n.id ? { ...x, read: true } : x))}>知道了</button>
                      </div>
                    </div>
                  </div>
                ))}
                {myFilmingNotices.map((n) => (
                  <div key={n.id} style={S.noticeBanner}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
                      <span>📹 {n.date} {n.start} 拍片時段已被佔用，自動取消</span>
                      <button style={S.linkBtn} onClick={() => setFilmingNotices((prev) => prev.map((x) => x.id === n.id ? { ...x, read: true } : x))}>知道了</button>
                    </div>
                  </div>
                ))}
              </>
            );
          })()}

          {isCoach && (() => {
            // 第0.1項：Reminder Card——今日場地使用一眼睇。三態：已預約／未有預約／已取消
            const todayStr = formatDate(new Date());
            const nowMin = (() => { const n = new Date(); return n.getHours() * 60 + n.getMinutes(); })();
            const toMin = (t) => { const [h, m] = t.split(":").map(Number); return h * 60 + m; };
            const todayBookings = myBookings.filter((b) => b.date === todayStr).sort((a, b) => toMin(a.start) - toMin(b.start));
            const todayCancelled = cancelLog.filter((c) => c.coachId === currentUser.id && c.date === todayStr);

            if (todayBookings.length > 0) {
              // 揀「最近一堂」：優先揀仲未完嘅（結束時間 >= 而家），冇就揀最後一堂（全部已完成）
              const upcoming = todayBookings.find((b) => toMin(b.start) + b.hours * 60 >= nowMin);
              const featured = upcoming || todayBookings[todayBookings.length - 1];
              const end = addMinutes(featured.start, featured.hours * 60);
              return (
                <div style={{ ...S.reminderCard, ...S.reminderCardConfirmed }}>
                  <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, color: "#888", marginBottom: 8 }}>
                    <span>今日場地使用</span>
                    {todayBookings.length > 1 && <span>今日共 {todayBookings.length} 堂</span>}
                  </div>
                  <div style={S.reminderStatusRow}>
                    <div style={S.reminderIconOk}>✓</div>
                    <div style={{ fontWeight: 700, color: "#4ECDC4" }}>{todayBookings.length > 1 ? "最近一堂" : "已預約"}</div>
                  </div>
                  <div style={{ marginLeft: 28, marginTop: 4 }}>
                    <div style={{ fontSize: 18, fontWeight: 700 }}>{featured.start} – {end}</div>
                    <div style={{ fontSize: 12, color: "#888" }}>{featured.type === "duo" ? "1對2" : "1對1"}{featured.students?.length ? " · " + featured.students.join("、") : ""}</div>
                  </div>
                  {todayBookings.length > 1 && (
                    <div style={S.reminderChipWrap}>
                      {todayBookings.map((b) => (
                        <span key={`${b.date}_${b.start}`} style={S.reminderChip}>{b.start}{b.start === featured.start ? " ← 最近" : ""}</span>
                      ))}
                    </div>
                  )}
                </div>
              );
            }
            if (todayCancelled.length > 0) {
              const c = todayCancelled[todayCancelled.length - 1];
              return (
                <div style={{ ...S.reminderCard, ...S.reminderCardCancelled }}>
                  <div style={{ fontSize: 11, color: "#888", marginBottom: 8 }}>今日場地使用</div>
                  <div style={S.reminderStatusRow}>
                    <div style={S.reminderIconDanger}>!</div>
                    <div style={{ fontWeight: 700, color: "#e2685a" }}>今日預約已取消</div>
                  </div>
                  <div style={{ marginLeft: 28, marginTop: 4, fontSize: 13, color: "#ddd" }}>原定 {c.start} 嘅場地使用已取消。如需使用，請重新 Booking。</div>
                  <button style={{ ...S.smallBtn, marginLeft: 28, marginTop: 10 }} onClick={() => { setQuickBook((q) => ({ ...q, date: todayStr, start: c.start })); document.getElementById("quickBookCard")?.scrollIntoView({ behavior: "smooth", block: "center" }); }}>重新 Booking</button>
                </div>
              );
            }
            return (
              <div style={{ ...S.reminderCard, ...S.reminderCardEmpty }}>
                <div style={{ fontSize: 11, color: "#888", marginBottom: 8 }}>今日場地使用</div>
                <div style={S.reminderStatusRow}>
                  <div style={S.reminderIconWarn}>!</div>
                  <div style={{ fontWeight: 700, color: "#FFB347" }}>今日尚未有預約</div>
                </div>
                <div style={{ marginLeft: 28, marginTop: 4, fontSize: 13, color: "#ddd" }}>如需使用場地，請先完成 Booking。</div>
                <button style={{ ...S.smallBtn, marginLeft: 28, marginTop: 10 }} onClick={() => { setQuickBook((q) => ({ ...q, date: todayStr })); document.getElementById("quickBookCard")?.scrollIntoView({ behavior: "smooth", block: "center" }); }}>＋ 建立 Booking</button>
              </div>
            );
          })()}

          {isCoach && (() => {
            // 第0.2項：快速Book表格——教練自己book用嘅表格式輸入，取代/補充grid點格仔。用返confirmBook同一套驗證同Pass邏輯，唔開新規則
            const qbDate = quickBook.date || formatDate(new Date());
            const isDuo = quickBook.sessionType === "duo";
            const passCost = isDuo ? Number(quickBook.hours) + 0.5 : Number(quickBook.hours);
            const end = addMinutes(quickBook.start, Number(quickBook.hours) * 60);
            const startOptions = TIME_SLOTS.filter((_, i) => i % 4 === 0); // 表格用整點做選項，減少揀嘢負擔；實際扣鐘同grid一樣以15分鐘為單位計算
            return (
              <div id="quickBookCard" style={S.quickBookCard}>
                <div style={{ fontSize: 11, color: "#888", marginBottom: 10, letterSpacing: 0.5 }}>快速 BOOK 堂</div>
                <div style={S.qbRow}>
                  <div style={S.qbField}>
                    <label style={{ ...S.label, fontSize: 11 }}>日期</label>
                    <input style={S.input} type="date" min={formatDate(new Date())} value={qbDate} onChange={(e) => setQuickBook({ ...quickBook, date: e.target.value })} />
                  </div>
                  <div style={S.qbField}>
                    <label style={{ ...S.label, fontSize: 11 }}>開始時間</label>
                    <select style={S.input} value={quickBook.start} onChange={(e) => setQuickBook({ ...quickBook, start: e.target.value })}>
                      {startOptions.map((t) => {
                        const full = canPlace(qbDate, t, Number(quickBook.hours) || 1, 1, currentUser.id, currentUser.id) !== null;
                        return <option key={t} value={t} disabled={full}>{t}{full ? "（已滿）" : ""}</option>;
                      })}
                    </select>
                  </div>
                </div>
                <div style={S.qbRow}>
                  <div style={S.qbField}>
                    <label style={{ ...S.label, fontSize: 11 }}>時長（最少 1 小時）</label>
                    <select style={S.input} value={quickBook.hours} onChange={(e) => setQuickBook({ ...quickBook, hours: Number(e.target.value) })}>
                      <option value={1}>1 小時</option>
                      <option value={1.5}>1.5 小時</option>
                      <option value={2}>2 小時</option>
                    </select>
                  </div>
                  <div style={S.qbField}>
                    <label style={{ ...S.label, fontSize: 11 }}>類型</label>
                    <div style={S.segRow}>
                      <button style={!isDuo ? S.segActive : S.seg} onClick={() => setQuickBook({ ...quickBook, sessionType: "solo" })}>1對1</button>
                      <button style={isDuo ? S.segActive : S.seg} onClick={() => setQuickBook({ ...quickBook, sessionType: "duo" })}>1對2</button>
                    </div>
                  </div>
                </div>
                {myRoster.length > 0 ? (
                  <div style={S.qbField}>
                    <label style={{ ...S.label, fontSize: 11 }}>學生</label>
                    <select style={S.input} value={(quickBook.students || [])[0] || ""} onChange={(e) => setQuickBook({ ...quickBook, students: e.target.value ? [e.target.value] : [] })}>
                      <option value="">請選擇學生</option>
                      {myRoster.map(({ name }) => <option key={name} value={name}>{name}</option>)}
                    </select>
                  </div>
                ) : (
                  <input style={S.input} value={quickBook.studentOther || ""} placeholder="學生名（打名就得）" onChange={(e) => setQuickBook({ ...quickBook, studentOther: e.target.value })} />
                )}
                <div style={S.qbPreview}>
                  <div style={S.qbPreviewRow}><span>將預約</span><strong style={{ color: "#4ECDC4" }}>{qbDate} {quickBook.start}–{end} · {isDuo ? "1對2" : "1對1"}</strong></div>
                  <div style={S.qbPreviewRow}><span>預計扣減</span><strong style={{ color: "#FFB347" }}>{passCost} 小時</strong></div>
                </div>
                <button style={S.modalConfirm} onClick={() => confirmBook({ date: qbDate, time: quickBook.start, sessionType: quickBook.sessionType, hours: Number(quickBook.hours), students: quickBook.students, studentOther: quickBook.studentOther, repeatWeeks: 1 })}>確認 Book 堂</button>
                <p style={S.assistHint}>同點格仔 book 堂用同一套時段衝突同Pass扣鐘規則，最終以送出時系統驗證為準。</p>
              </div>
            );
          })()}
        </div>
      )}

      {view === "calendar" && (
        <div style={S.calContainer}>
          <div style={S.weekNav}>
            <button style={S.navBtn} onClick={() => setWeekOffset((w) => w - 1)}>‹ 上週</button>
            <span style={S.weekLabel}>{formatDate(days[0])} – {formatDate(days[6])}</span>
            <button style={S.navBtn} onClick={() => setWeekOffset(0)}>今日</button>
            <button style={S.navBtn} onClick={() => setWeekOffset((w) => w + 1)}>下週 ›</button>
            <button style={S.navBtn} onClick={() => { setWeekViewMode((m) => m === "fixed" ? "rolling" : "fixed"); setWeekOffset(0); }} title="切換週視圖模式">🔁 {weekViewMode === "fixed" ? "一至日" : "今日起"}</button>
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
                        const here = visibleCellArr(date, time, currentUser.id);
                        const occ = occupancy(date, time, currentUser.id);
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
          {(() => {
            const myNotices = filmingNotices.filter((n) => n.coachId === currentUser.id && !n.read);
            return myNotices.length > 0 && (
              <div style={S.noticeBanner}>
                {myNotices.map((n) => (
                  <div key={n.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
                    <span>🎬 你 {n.date} {n.start} 嘅拍片安排已經被教練 book 走，請另揀時間</span>
                    <button style={S.linkBtn} onClick={() => setFilmingNotices((prev) => prev.map((x) => x.id === n.id ? { ...x, read: true } : x))}>知道了</button>
                  </div>
                ))}
              </div>
            );
          })()}
          {(() => {
            // 第9項：我的 Pass 資訊panel——個人/彈性/共享 breakdown，共享嗰張可以自己加值
            const pRemain2 = personalRemaining(currentUser.id);
            const fRemain2 = flexibleRemaining(currentUser.id);
            const myShared = sharedPassesOf(currentUser.id);
            const hasAnyPass = pRemain2 > 0 || fRemain2 > 0 || myShared.length > 0 || purchaseLog.some((r) => r.coachId === currentUser.id && r.passType);
            if (!hasAnyPass) return null;
            return (
              <div style={{ ...S.formCard, marginBottom: 14 }}>
                <div style={{ ...S.assistHint, marginBottom: 6 }}>我的 Training Pass</div>
                {pRemain2 > 0 && <div style={S.bookingTime}>個人證（solo限定）剩 {pRemain2} 小時</div>}
                {fRemain2 > 0 && <div style={S.bookingTime}>彈性／舊制剩 {fRemain2} 小時</div>}
                {myShared.map((sp) => {
                  const other = (sp.coachIds || []).find((id) => id !== currentUser.id);
                  return (
                    <div key={sp.id} style={{ ...S.flexBetween, marginTop: 4 }}>
                      <span style={S.bookingTime}>共享 Pass（同 {getCoach(other)?.name || "?"}）剩 {sharedRemaining(sp)} / {sp.totalHours} 小時</span>
                      <button style={S.linkBtn} onClick={() => setSharedTopUpModal({ sharedId: sp.id, qty: 5 })}>+ 加值</button>
                    </div>
                  );
                })}
              </div>
            );
          })()}
          {(() => {
            // 第9-3項：Pass 用晒／過期提醒（純banner，過期唔會鎖，淨係提示）
            const pRemain = personalRemaining(currentUser.id);
            const fRemain = flexibleRemaining(currentUser.id);
            const hasAnyPassPurchase = purchaseLog.some((r) => r.coachId === currentUser.id && r.passType);
            const usedUp = hasAnyPassPurchase && pRemain <= 0 && fRemain <= 0;
            const expiredBatches = expiredPassBatchesOf(currentUser.id);
            if (!usedUp && expiredBatches.length === 0) return null;
            return (
              <div style={{ ...S.noticeBanner, background: "#332a0f" }}>
                {usedUp && <div>⚠️ 你嘅 Training Pass 時數已經用晒，請聯絡 admin 購買新 Pass。</div>}
                {expiredBatches.length > 0 && <div>⏰ 你有 Pass 已經過期（仍然可以用）：{expiredBatches.map((b) => `${b.remaining}小時@${b.expiryDate}`).join("、")}</div>}
              </div>
            );
          })()}
          <div style={S.flexBetween}>
            <h2 style={S.sectionTitle}>我的預約記錄</h2>
            <div style={S.segRow}>
              <button style={myBookingsView === "list" ? S.segActive : S.seg} onClick={() => setMyBookingsView("list")}>📋 列表</button>
              <button style={myBookingsView === "calendar" ? S.segActive : S.seg} onClick={() => setMyBookingsView("calendar")}>📅 圖像</button>
            </div>
          </div>
          {myBookingsView === "list" && (
            <div style={{ ...S.segRow, marginTop: 8 }}>
              <button style={myBookingsSortMode === "newest" ? S.segActive : S.seg} onClick={() => setMyBookingsSortMode("newest")}>新到舊</button>
              <button style={myBookingsSortMode === "closest" ? S.segActive : S.seg} onClick={() => setMyBookingsSortMode("closest")}>距今日最近</button>
            </div>
          )}
          {myBookingsView === "calendar" ? (
            <div style={{ marginTop: 14 }}>
              <div style={S.weekNav}>
                <button style={S.navBtn} onClick={() => setWeekOffset((w) => w - 1)}>‹ 上週</button>
                <span style={S.weekLabel}>{formatDate(days[0])} – {formatDate(days[6])}</span>
                <button style={S.navBtn} onClick={() => setWeekOffset(0)}>今日</button>
                <button style={S.navBtn} onClick={() => setWeekOffset((w) => w + 1)}>下週 ›</button>
                <button style={S.navBtn} onClick={() => { setWeekViewMode((m) => m === "fixed" ? "rolling" : "fixed"); setWeekOffset(0); }} title="切換週視圖模式">🔁 {weekViewMode === "fixed" ? "一至日" : "今日起"}</button>
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
                            const here = visibleCellArr(date, time, currentUser.id);
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
          ) : myBookingsSorted.length === 0 ? <p style={S.emptyText}>你還未有預約</p> : (
            <div style={S.bookingList}>
              {myBookingsSorted.map(({ date, start, hours, type, charterType, coachName, students, signatures }, i) => {
                const hrs = hoursUntil(date, start);
                const isPast = hrs < 0;
                const isFilming = type === "charter" && charterType === "filming";
                const locked = !isFilming && hrs >= 0 && hrs < (liveUser.cancelWindowHours ?? 24);
                return (
                  <div key={i} style={S.bookingItem}>
                    <div style={{ ...S.dot, background: liveUser.color }} />
                    <div style={{ flex: 1 }}>
                      <div style={S.bookingCoach}>{date} <span style={isFilming ? S.filmingTag : type === "duo" ? S.duoTag : S.soloTag}>{isFilming ? "🎬 拍片" : type === "duo" ? "1對2" : "1對1"}</span></div>
                      <div style={S.bookingTime}>{start} – {addMinutes(start, hours * 60)}（{hours}小時）</div>
                      {!isFilming && <button style={S.qrBtn} onClick={() => openWhatsAppQR(date, start, hours, liveUser.name)}>📲 攞 QR Code</button>}
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
                      : <button style={S.cancelBtn} onClick={() => openCancel(date, start, currentUser.id, type, charterType)}>取消</button>}
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
          <div style={S.flexBetween}>
            <h2 style={S.sectionTitle}>我的收入（近3個月）</h2>
            <button style={S.creditBtn} onClick={exportMyIncomeSheet}>📊 匯出 Google Sheet</button>
          </div>
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
          {roster.length > 1 && (
            <div style={S.segRow}>
              <button style={rosterSortMode === "custom" ? S.segActive : S.seg} onClick={() => setRosterSortMode("custom")}>自訂次序</button>
              <button style={rosterSortMode === "used" ? S.segActive : S.seg} onClick={() => setRosterSortMode("used")}>上堂數</button>
              <button style={rosterSortMode === "remain" ? S.segActive : S.seg} onClick={() => setRosterSortMode("remain")}>剩餘堂數</button>
              <button style={rosterSortMode === "name" ? S.segActive : S.seg} onClick={() => setRosterSortMode("name")}>名稱</button>
            </div>
          )}
          {roster.length === 0 && <p style={S.emptyText}>仲未有學生，落面新增啦</p>}
          <div style={S.bookingList}>
            {(() => {
              const sorted = [...roster];
              if (rosterSortMode === "used") sorted.sort((a, b) => (b.used || 0) - (a.used || 0));
              else if (rosterSortMode === "remain") sorted.sort((a, b) => ((b.credits || 0) - (b.used || 0)) - ((a.credits || 0) - (a.used || 0)));
              else if (rosterSortMode === "name") sorted.sort((a, b) => a.name.localeCompare(b.name));
              return sorted.map((s, idx) => {
              const remain = (s.credits || 0) - (s.used || 0);
              const low = remain <= LOW_CREDIT_THRESHOLD;
              const open = studentLogOpen === s.name;
              const log = myIncomeReport.studentLog[s.name] || [];
              return (
                <div key={s.name}>
                  <div style={{ ...S.coachStatRow, cursor: "pointer" }} onClick={() => setStudentLogOpen(open ? null : s.name)}>
                    {rosterSortMode === "custom" && (
                      <div style={{ display: "flex", flexDirection: "column", marginRight: 6 }} onClick={(e) => e.stopPropagation()}>
                        <button style={S.moveBtn} disabled={idx === 0} onClick={() => moveStudent(s.name, -1)}>▲</button>
                        <button style={S.moveBtn} disabled={idx === sorted.length - 1} onClick={() => moveStudent(s.name, 1)}>▼</button>
                      </div>
                    )}
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
                      <Field label="每堂收費 ($)"><input style={S.input} type="number" min="0"
                        value={studentDrafts[`${s.name}_rate`] !== undefined ? studentDrafts[`${s.name}_rate`] : s.rate}
                        onChange={(e) => {
                          const v = e.target.value;
                          setStudentDrafts((prev) => ({ ...prev, [`${s.name}_rate`]: v }));
                          if (v !== "") updateStudentField(s.name, "rate", Number(v) || 0);
                        }}
                        onBlur={() => setStudentDrafts((prev) => { const n = { ...prev }; delete n[`${s.name}_rate`]; return n; })} /></Field>
                      <div style={S.bookingTime}>已開 {s.credits || 0} 堂　已用 {s.used || 0} 堂</div>
                      <Field label="剩餘堂數">
                        <input style={{ ...S.input, borderColor: low ? "#5a2020" : undefined, color: low ? "#FF8FA3" : "#4ECDC4", fontWeight: 700 }}
                          type="number"
                          value={studentDrafts[`${s.name}_remain`] !== undefined ? studentDrafts[`${s.name}_remain`] : remain}
                          onChange={(e) => {
                            const v = e.target.value;
                            setStudentDrafts((prev) => ({ ...prev, [`${s.name}_remain`]: v }));
                            if (v !== "") setStudentRemain(s.name, Number(v) || 0);
                          }}
                          onBlur={() => setStudentDrafts((prev) => { const n = { ...prev }; delete n[`${s.name}_remain`]; return n; })} />
                      </Field>
                      <div style={{ display: "flex", gap: 8, marginTop: 4, marginBottom: 14 }}>
                        <button style={{ ...S.creditBtn, flex: 1 }} onClick={() => setAddStudentCreditModal({ name: s.name, qty: 1, date: formatDate(new Date()), expiryDate: "" })}>+ 幫佢開堂數</button>
                        <button style={{ ...S.smallBtn, flex: 1 }} onClick={() => setSigReportModal({ studentName: s.name, month: monthKey(formatDate(new Date())) })}>📄 簽到月報表</button>
                      </div>

                      <div style={{ ...S.assistHint, marginBottom: 4 }}>購堂紀錄</div>
                      {(() => {
                        const buyLog = studentPurchaseLog.filter((r) => r.coachId === currentUser.id && r.studentName === s.name);
                        return buyLog.length === 0 ? <p style={{ ...S.emptyText, padding: "8px 0" }}>暫無購堂紀錄</p> : (
                          <div style={{ marginBottom: 14 }}>
                            {buyLog.map((r) => {
                              const today = formatDate(new Date());
                              const expSoon = r.expiryDate && (new Date(`${r.expiryDate}T00:00:00`) - new Date(`${today}T00:00:00`)) / 86400000 <= 14;
                              const expPast = r.expiryDate && r.expiryDate < today;
                              return (
                                <div key={r.id} style={S.purchaseRow}>
                                  <div style={S.bookingTime}>
                                    {r.date}　+{r.qty} 堂　@${r.rate}/堂
                                    {r.expiryDate && <span style={{ color: expPast ? "#FF6B6B" : expSoon ? "#FFB347" : "#777" }}>　{expPast ? "已過期" : "失效"}：{r.expiryDate}</span>}
                                  </div>
                                  <div style={{ color: "#6BCB77", fontWeight: 700 }}>${r.amount.toLocaleString()}</div>
                                </div>
                              );
                            })}
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
              });
            })()}
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

          <h2 style={{ ...S.sectionTitle, marginTop: 28 }}>🥤 飲品訂購</h2>
          <div style={S.formCard}>
            <button style={{ ...S.smallBtn, width: "100%", textAlign: "center" }} onClick={() => setDrinkOrderOpen((o) => !o)}>
              {drinkOrderOpen ? "▲ 收埋" : "▼ 打開飲品清單"}{drinkCartCount() > 0 ? `（已揀 ${drinkCartCount()} 支）` : ""}
            </button>
            {drinkOrderOpen && (
              drinkProducts.length === 0 ? (
                <p style={{ ...S.emptyText, marginTop: 12 }}>Admin 仲未上架任何飲品</p>
              ) : (
                <>
                  <div style={{ maxHeight: 320, overflowY: "auto", marginTop: 12 }}>
                    {drinkProducts.map((p) => {
                      const qty = Number(drinkCart[p.id]) || 0;
                      return (
                        <div key={p.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 0", borderBottom: "1px solid #222" }}>
                          <div>
                            <div style={{ fontWeight: 600 }}>{p.name}</div>
                            <div style={S.assistHint}>${p.price} / 支</div>
                          </div>
                          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                            <button style={S.smallBtn} onClick={() => setDrinkCart((c) => ({ ...c, [p.id]: Math.max(0, qty - 1) }))}>－</button>
                            <span style={{ minWidth: 20, textAlign: "center" }}>{qty}</span>
                            <button style={S.smallBtn} onClick={() => setDrinkCart((c) => ({ ...c, [p.id]: qty + 1 }))}>＋</button>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", marginTop: 14, fontWeight: 700 }}>
                    <span>合共</span><span>${drinkCartTotal()}</span>
                  </div>
                  <button style={{ ...S.loginBtn, marginTop: 10 }} onClick={openDrinkCheckout}>下一步：顯示收款 QR</button>
                </>
              )
            )}
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

      {retroBookModal && (
        <div style={S.modalOverlay}><div style={S.modal}>
          <h3 style={S.modalTitle}>補book記錄</h3>
          <p style={{ ...S.modalText, textAlign: "center" }}>{retroBookModal.date}　{retroBookModal.start}–{addMinutes(retroBookModal.start, Number(retroBookModal.hours) * 60)}（{retroBookModal.hours}小時）</p>
          <label style={S.label}>類型</label>
          <div style={S.segRow}>
            <button style={retroBookModal.sessionType === "solo" ? S.segActive : S.seg} onClick={() => setRetroBookModal({ ...retroBookModal, sessionType: "solo" })}>1對1</button>
            <button style={retroBookModal.sessionType === "duo" ? S.segActive : S.seg} onClick={() => setRetroBookModal({ ...retroBookModal, sessionType: "duo" })}>1對2</button>
          </div>
          <label style={{ ...S.label, marginTop: 14 }}>學生（最多4位）</label>
          {myRoster.length === 0 ? (
            <p style={S.assistHint}>你仲未有學生名單，可以喺「上堂情況」分頁新增。</p>
          ) : (
            <div style={S.studentChipWrap}>
              {myRoster.map(({ name }) => {
                const sel = Array.isArray(retroBookModal.students) && retroBookModal.students.includes(name);
                const atMax = !sel && (retroBookModal.students || []).length >= 4;
                return (
                  <button key={name} disabled={atMax} style={sel ? S.studentChipActive : atMax ? S.studentChipDisabled : S.studentChip}
                    onClick={() => {
                      const cur = retroBookModal.students || [];
                      setRetroBookModal({ ...retroBookModal, students: sel ? cur.filter((n) => n !== name) : [...cur, name] });
                    }}>{name}</button>
                );
              })}
            </div>
          )}
          <p style={S.assistHint}>補book都會照樣扣Pass時數，同準時book冇分別；如果嗰個時段已經俾人book咗，會提示你聯絡Admin處理。</p>
          <div style={S.modalBtns}>
            <button style={S.modalCancel} onClick={() => setRetroBookModal(null)}>取消</button>
            <button style={S.modalConfirm} onClick={confirmRetroBooking}>確認補book</button>
          </div>
        </div></div>
      )}

      {drinkQrModal && (
        <div style={S.modalOverlay}><div style={S.modal}>
          <h3 style={S.modalTitle}>飲品付款</h3>
          <div style={{ textAlign: "left", margin: "10px 0" }}>
            {drinkQrModal.items.map((it) => (
              <div key={it.productId} style={{ display: "flex", justifyContent: "space-between", padding: "4px 0" }}>
                <span>{it.name} × {it.qty}</span><span>${it.price * it.qty}</span>
              </div>
            ))}
          </div>
          <p style={{ ...S.modalTitle, fontSize: 22 }}>合共 ${drinkQrModal.amount}</p>
          {paymentQR ? (
            <div style={{ textAlign: "center", margin: "10px 0" }}>
              <img src={paymentQR} alt="收款 QR" style={{ maxWidth: 200, width: "100%", borderRadius: 10, background: "#fff", padding: 8, boxSizing: "border-box" }} />
            </div>
          ) : (
            <p style={S.emptyText}>Admin 仲未上傳收款 QR，請叫學生直接搵你轉數。</p>
          )}
          <p style={S.assistHint}>請畀學生掃碼自行轉數（用學生自己部電話嘅銀行 App），系統唔會自動核實款項有冇入到。</p>
          <div style={S.modalBtns}>
            <button style={S.modalCancel} onClick={() => setDrinkQrModal(null)}>取消</button>
            <button style={S.modalConfirm} onClick={confirmDrinkSale}>已顯示，記錄呢單</button>
          </div>
        </div></div>
      )}

      {bookModal && (() => {
        const isDuo = bookModal.sessionType === "duo";
        const isFilming = bookModal.sessionType === "filming";
        const passCost = isDuo ? bookModal.hours + 0.5 : bookModal.hours;
        const price = passCost * 100;
        const allowSolo = liveUser.allowSolo !== false;
        const allowDuo = liveUser.allowDuo !== false;
        const allowFilming = liveUser.allowFilming === true;
        const pRemain = personalRemaining(currentUser.id);
        const fRemain = flexibleRemaining(currentUser.id);
        const poolLabel = !isFilming ? (
          bookModal.sessionType === "solo" && pRemain >= passCost ? "個人證" :
          fRemain >= passCost ? "彈性／舊制" :
          sharedPassesOf(currentUser.id).some((sp) => sharedRemaining(sp) >= passCost) ? "共享 Pass" : "時數不足"
        ) : "";
        return (
          <div style={S.modalOverlay}><div style={{ ...S.modal, textAlign: "left" }}>
            <h3 style={{ ...S.modalTitle, textAlign: "center" }}>預約場地</h3>
            <p style={{ ...S.modalText, textAlign: "center" }}>{bookModal.date}　{bookModal.time}</p>
            <label style={S.label}>類型</label>
            <div style={S.segRow}>
              <button style={!allowSolo ? S.segDisabled : (!isDuo && !isFilming) ? S.segActive : S.seg} disabled={!allowSolo} onClick={() => allowSolo && setBookModal({ ...bookModal, sessionType: "solo" })}>1對1</button>
              <button style={!allowDuo ? S.segDisabled : isDuo ? S.segActive : S.seg} disabled={!allowDuo} onClick={() => allowDuo && setBookModal({ ...bookModal, sessionType: "duo" })}>1對2</button>
              {allowFilming && <button style={isFilming ? S.segActive : S.seg} onClick={() => setBookModal({ ...bookModal, sessionType: "filming" })}>🎬 拍片</button>}
            </div>
            {isFilming && <p style={S.assistHint}>拍片佔全場、$0、唔扣時數。其他冇拍片權限嘅教練見唔到呢格（當空格）；如果有人 book 中，呢個拍片安排會自動取消，你會收到通知。</p>}
            <label style={{ ...S.label, marginTop: 14 }}>時長</label>
            <div style={S.segRow}>
              <button style={bookModal.hours === 1 ? S.segActive : S.seg} onClick={() => setBookModal({ ...bookModal, hours: 1 })}>1 小時</button>
              <button style={bookModal.hours === 1.5 ? S.segActive : S.seg} onClick={() => setBookModal({ ...bookModal, hours: 1.5 })}>1.5 小時</button>
            </div>
            {!isFilming && (
              <>
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
                {(bookModal.repeatWeeks || 1) > 1 && <p style={S.assistHint}>會一次過幫你book未來 {bookModal.repeatWeeks} 個星期嘅同一個時段；如果某一週已經被佔用或Pass時數不足，會自動跳過嗰一週，唔影響其他週。</p>}
              </>
            )}
            <div style={S.priceBox}>
              <div style={S.priceRow}><span>時段</span><span>{bookModal.time} – {addMinutes(bookModal.time, bookModal.hours * 60)}</span></div>
              {isFilming ? (
                <div style={{ ...S.priceRow, color: "#4ECDC4", fontWeight: 700, fontSize: 16 }}><span>拍片</span><span>$0（唔扣時數）</span></div>
              ) : (
                <>
                  <div style={S.priceRow}><span>扣 Pass 時數</span><span>{passCost} 小時{isDuo ? "（1對2額外+0.5）" : ""}{(bookModal.repeatWeeks || 1) > 1 ? `（每週，最多扣 ${(passCost * bookModal.repeatWeeks).toFixed(1)} 小時）` : ""}</span></div>
                  <div style={S.priceRow}><span>扣邊個池</span><span>{poolLabel}</span></div>
                  <div style={{ ...S.priceRow, color: "#4ECDC4", fontWeight: 700, fontSize: 16 }}><span>{isDuo ? "1對2 收費" : "1對1 收費"}</span><span>${price}{(bookModal.repeatWeeks || 1) > 1 ? "／週" : ""}</span></div>
                </>
              )}
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
          <p style={S.modalText}>{cancelModal.date}　{cancelModal.start}<br />取消後退回對應時數</p>
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
            onChange={(e) => setAddStudentCreditModal({ ...addStudentCreditModal, qty: e.target.value })} /></Field>
          <Field label="增加日期"><input style={S.input} type="date" value={addStudentCreditModal.date || formatDate(new Date())}
            onChange={(e) => setAddStudentCreditModal({ ...addStudentCreditModal, date: e.target.value })} /></Field>
          <Field label="失效日期（可選，留空＝冇限期）"><input style={S.input} type="date" value={addStudentCreditModal.expiryDate || ""}
            onChange={(e) => setAddStudentCreditModal({ ...addStudentCreditModal, expiryDate: e.target.value })} /></Field>
          <div style={S.modalBtns}>
            <button style={S.modalCancel} onClick={() => setAddStudentCreditModal(null)}>取消</button>
            <button style={S.modalConfirm} onClick={() => { const qty = Number(addStudentCreditModal.qty) || 0; if (qty <= 0) { showToast("請輸入有效堂數", "error"); return; } addStudentCredits(addStudentCreditModal.name, qty, addStudentCreditModal.date, addStudentCreditModal.expiryDate); setAddStudentCreditModal(null); }}>確認增加</button>
          </div>
        </div></div>
      )}
      {sharedPassModal && (
        <div style={S.modalOverlay}><div style={S.modal}>
          <h3 style={S.modalTitle}>開共享訓練通行證</h3>
          <p style={S.assistHint}>{SHARED_PASS_HOURS}小時／{SHARED_PASS_MONTHS}個月，硬性限制2位教練，任何一位或 admin 都可以之後幫呢張 Pass 加值。</p>
          <label style={S.label}>教練 A</label>
          <select style={{ ...S.select, width: "100%", boxSizing: "border-box" }} value={sharedPassModal.coachIdA} onChange={(e) => setSharedPassModal({ ...sharedPassModal, coachIdA: Number(e.target.value) })}>
            {coaches.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
          <label style={{ ...S.label, marginTop: 10 }}>教練 B</label>
          <select style={{ ...S.select, width: "100%", boxSizing: "border-box" }} value={sharedPassModal.coachIdB} onChange={(e) => setSharedPassModal({ ...sharedPassModal, coachIdB: Number(e.target.value) })}>
            {coaches.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
          <Field label="開始日期"><input style={S.input} type="date" value={sharedPassModal.date || formatDate(new Date())} onChange={(e) => setSharedPassModal({ ...sharedPassModal, date: e.target.value })} /></Field>
          <div style={S.modalBtns}>
            <button style={S.modalCancel} onClick={() => setSharedPassModal(null)}>取消</button>
            <button style={S.modalConfirm} onClick={() => { createSharedPass(sharedPassModal.coachIdA, sharedPassModal.coachIdB, sharedPassModal.date); setSharedPassModal(null); }}>確認開卡</button>
          </div>
        </div></div>
      )}
      {sharedTopUpModal && (
        <div style={S.modalOverlay}><div style={S.modal}>
          <h3 style={S.modalTitle}>幫共享 Pass 加值</h3>
          <Field label="加幾多小時"><input style={S.input} type="number" step="0.5" min="0.5" value={sharedTopUpModal.qty} onChange={(e) => setSharedTopUpModal({ ...sharedTopUpModal, qty: e.target.value })} /></Field>
          <div style={S.modalBtns}>
            <button style={S.modalCancel} onClick={() => setSharedTopUpModal(null)}>取消</button>
            <button style={S.modalConfirm} onClick={() => { const qty = Number(sharedTopUpModal.qty) || 0; if (qty <= 0) { showToast("請輸入有效小時數", "error"); return; } addSharedPassHours(sharedTopUpModal.sharedId, qty); setSharedTopUpModal(null); }}>確認加值</button>
          </div>
        </div></div>
      )}
      {sigReportModal && (
        <div style={S.modalOverlay}><div style={S.modal}>
          <h3 style={S.modalTitle}>簽到月報表</h3>
          <p style={S.modalText}>{sigReportModal.studentName}</p>
          <Field label="月份"><input style={S.select} type="month" value={sigReportModal.month} onChange={(e) => setSigReportModal({ ...sigReportModal, month: e.target.value })} /></Field>
          <p style={S.assistHint}>會生成一份 PDF，列出呢個月全部上堂記錄（連簽名圖），冇簽到嘅堂會標記「未簽」。</p>
          <div style={S.modalBtns}>
            <button style={S.modalCancel} onClick={() => setSigReportModal(null)}>取消</button>
            <button style={S.modalConfirm} onClick={async () => { await generateSignatureReportPDF(sigReportModal.studentName, sigReportModal.month); setSigReportModal(null); }}>生成 PDF</button>
          </div>
        </div></div>
      )}
      {toast && <Toast toast={toast} />}
    </div>
  );
}

// 獨立嘅 presentational components（唔直接讀寫 App() 嘅 state，淨係收 props）
import { useState, useRef, useEffect } from "react";
import { LOGO, BRAND_NAME } from "./brand.js";
import { S } from "./styles.js";

export function EditCoachModal({ coach, onClose, onSave }) {
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
export function Field({ label, children }) { return <div style={S.inputGroup}><label style={S.label}>{label}</label>{children}</div>; }

// 簽名板：支援滑鼠同觸控（手機）畫簽名，輸出 base64 PNG
export function SignaturePad({ studentName, onSave, onCancel }) {
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
export function Header({ title, onLogout, syncState }) {
  const sync = {
    synced: { t: "☁️ 已同步", c: "#6BCB77" },
    connecting: { t: "⟳ 同步中", c: "#FFB347" },
    error: { t: "⚠️ 同步失敗", c: "#FF6B6B" },
    local: { t: "📱 本機", c: "#777" },
  }[syncState] || null;
  return <div style={S.header}><img src={LOGO} alt={BRAND_NAME} style={S.headerLogoImg} /><span style={S.headerUser}>{title}</span>{sync && <span style={{ ...S.syncBadge, color: sync.c }}>{sync.t}</span>}<button style={S.logoutBtn} onClick={onLogout}>登出</button></div>;
}
export function Toast({ toast }) { return <div style={{ ...S.toast, background: toast.type === "error" ? "#FF6B6B" : "#4ECDC4" }}>{toast.msg}</div>; }

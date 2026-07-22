import { useState, useEffect, useMemo, useRef, useCallback } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import { supabase } from "./lib/supabaseClient";

const PAY_METHODS = ["Cash", "Card", "Bank transfer", "Online"];
const CARD_PROVIDERS = ["Com Bank", "DFCC", "NTB", "Global"];
const CURRENCIES = ["LKR", "USD", "EUR", "GBP", "AUD"];
const CUR_SYMBOL = { LKR: "Rs", USD: "$", EUR: "€", GBP: "£", AUD: "A$" };
const RATE_CURRENCIES = ["USD", "EUR", "GBP"]; // exchange rates supported for now
const ROOM_STATUSES = [
  { id: "vacant", label: "Vacant" },
  { id: "checkin", label: "Check-in" },
  { id: "staying", label: "Staying" },
  { id: "checkout", label: "Check-out" },
];
const STATUS_STYLE = {
  vacant: "bg-white border-slate-200 text-slate-500",
  checkin: "bg-sky-50 border-sky-300 text-sky-800",
  staying: "bg-teal-50 border-teal-300 text-teal-800",
  checkout: "bg-amber-50 border-amber-300 text-amber-800",
};
const BCOM_CAT = "Room – Booking.com";
const BCOM_RATE = 0.18;
const CARD_COMMISSION_RATE = 0.03;
const ONLINE_COMMISSION_RATE = 0.02;

const fmt = (n) => "Rs " + Math.round(n || 0).toLocaleString("en-LK");
const fmtCur = (n, currency) => {
  const cur = currency || "LKR";
  const sym = CUR_SYMBOL[cur] || cur;
  if (cur === "LKR") return sym + " " + Math.round(n || 0).toLocaleString("en-LK");
  return sym + " " + (Number(n) || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};
const todayStr = () => new Date().toISOString().slice(0, 10);
const isRoomCat = (c) => c && c.startsWith("Room");

// ---- Supabase <-> app-shape mapping ----
const rowToTxn = (r) => ({
  id: r.id,
  type: r.type,
  amount: Number(r.amount),
  currency: r.currency || "LKR",
  category: r.category,
  room: r.room || undefined,
  guestEvent: r.guest_event || undefined,
  guestName: r.guest_name || undefined,
  method: r.method,
  date: r.txn_date,
  note: r.note || "",
  status: r.status,
  enteredBy: r.entered_by,
  edited: r.edited,
});
const rowToRoom = (r) => ({ no: r.no, status: r.status, guest: r.guest || "", out: r.checkout_date || "" });

function Wave() {
  return (
    <svg viewBox="0 0 400 16" className="w-full block" preserveAspectRatio="none" style={{ height: 12 }} aria-hidden="true">
      <path d="M0 8 Q 25 0, 50 8 T 100 8 T 150 8 T 200 8 T 250 8 T 300 8 T 350 8 T 400 8 V16 H0 Z" fill="#0d5c5c" opacity="0.12" />
    </svg>
  );
}

function StatCard({ label, value, tone }) {
  const color = tone === "up" ? "text-emerald-700" : tone === "down" ? "text-rose-700" : "text-slate-800";
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-3">
      <div className="text-xs text-slate-500 mb-1">{label}</div>
      <div className={"text-lg font-semibold tabular-nums " + color}>{value}</div>
    </div>
  );
}

const GUEST_EVENT_STYLE = {
  checkin: "bg-sky-50 text-sky-700 border-sky-200",
  staying: "bg-teal-50 text-teal-700 border-teal-200",
  checkout: "bg-amber-50 text-amber-700 border-amber-200",
};
const GUEST_EVENT_LABEL = { checkin: "Check-in", staying: "Staying", checkout: "Check-out" };

function TxnRow({ t, onDelete, onMarkPaid, onEdit, history }) {
  const income = t.type === "income";
  const pending = t.status === "pending";
  const [showHistory, setShowHistory] = useState(false);
  return (
    <div className="py-2.5 border-b border-slate-100 last:border-0">
      <div className="flex items-center gap-3">
        <div className={"w-8 h-8 rounded-full flex items-center justify-center shrink-0 text-sm font-semibold " + (income ? "bg-emerald-50 text-emerald-700" : "bg-rose-50 text-rose-700")}>
          {income ? "+" : "−"}
        </div>
        <div className="flex-1 min-w-0">
          <div className="text-sm font-medium text-slate-800 truncate flex items-center gap-2">
            <span className="truncate">{t.category}{t.room ? " · Rm " + t.room : ""}</span>
            {t.guestEvent && (
              <span className={"text-[10px] font-semibold uppercase tracking-wide border rounded px-1.5 py-0.5 shrink-0 " + GUEST_EVENT_STYLE[t.guestEvent]}>
                {GUEST_EVENT_LABEL[t.guestEvent]}
              </span>
            )}
            {t.currency && t.currency !== "LKR" && <span className="text-[10px] font-semibold uppercase tracking-wide bg-indigo-50 text-indigo-700 border border-indigo-200 rounded px-1.5 py-0.5 shrink-0">{t.currency}</span>}
            {pending && <span className="text-[10px] font-semibold uppercase tracking-wide bg-amber-50 text-amber-700 border border-amber-200 rounded px-1.5 py-0.5 shrink-0">Pending</span>}
            {t.edited && (
              <button onClick={() => setShowHistory(!showHistory)} className="text-[10px] font-semibold uppercase tracking-wide bg-purple-50 text-purple-700 border border-purple-200 rounded px-1.5 py-0.5 shrink-0">Edited</button>
            )}
          </div>
          <div className="text-xs text-slate-500 truncate">{t.guestName ? t.guestName + " · " : ""}{t.note ? t.note + " · " : ""}{t.method} · {t.date}{t.enteredBy ? " · by " + t.enteredBy : ""}</div>
        </div>
        <div className={"text-sm font-semibold tabular-nums " + (income ? "text-emerald-700" : "text-rose-700")}>
          {income ? "+" : "−"}{fmtCur(t.amount, t.currency).replace(/^\D+\s?/, "")}
        </div>
        {pending && onMarkPaid && (
          <button onClick={() => onMarkPaid(t.id)} className="text-[10px] font-semibold text-teal-700 border border-teal-200 rounded px-1.5 py-1 shrink-0">Mark paid</button>
        )}
        {onEdit && <button onClick={() => onEdit(t.id)} className="text-slate-400 hover:text-teal-700 text-sm px-1" aria-label="Edit entry">✎</button>}
        {onDelete && <button onClick={() => onDelete(t.id)} className="text-slate-300 hover:text-rose-500 text-xs px-1" aria-label="Delete entry">✕</button>}
      </div>
      {showHistory && history && history.length > 0 && (
        <div className="ml-11 mt-1.5 bg-purple-50 border border-purple-200 rounded-lg p-2">
          <p className="text-[10px] font-semibold text-purple-700 uppercase tracking-wide mb-1">Change history</p>
          {history.map((h) => (
            <p key={h.id} className="text-[11px] text-purple-800 mb-0.5 last:mb-0">
              {new Date(h.changed_at).toLocaleString("en-LK", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" })} by {h.changed_by}: {h.summary}
            </p>
          ))}
        </div>
      )}
    </div>
  );
}

function CategoryPicker({ value, onChange, options, onAddNew }) {
  const [open, setOpen] = useState(false);
  const [q, setQ] = useState("");
  const ref = useRef(null);
  useEffect(() => {
    const handler = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);
  const filtered = options.filter((o) => o.toLowerCase().includes(q.trim().toLowerCase()));
  const exact = options.some((o) => o.toLowerCase() === q.trim().toLowerCase());
  const pick = (o) => { onChange(o); setOpen(false); setQ(""); };
  return (
    <div className="relative mb-4" ref={ref}>
      <button onClick={() => setOpen(!open)} className="w-full border border-slate-200 rounded-lg px-3 py-2.5 text-sm bg-white text-left flex items-center justify-between">
        <span className="truncate">{value || "Select category"}</span>
        <span className="text-slate-400 ml-2" aria-hidden="true">▾</span>
      </button>
      {open && (
        <div className="absolute z-20 left-0 right-0 bg-white border border-slate-200 rounded-lg mt-1 shadow-lg overflow-hidden">
          <input autoFocus value={q} onChange={(e) => setQ(e.target.value)} placeholder="Type to search…"
            className="w-full px-3 py-2.5 text-sm border-b border-slate-100 focus:outline-none" />
          <div className="max-h-52 overflow-y-auto">
            {filtered.map((o) => (
              <button key={o} onClick={() => pick(o)}
                className={"w-full text-left px-3 py-2 text-sm border-b border-slate-50 last:border-0 " + (o === value ? "bg-teal-50 text-teal-800 font-medium" : "text-slate-700 hover:bg-slate-50")}>
                {o}
              </button>
            ))}
            {filtered.length === 0 && !q.trim() && <p className="px-3 py-3 text-xs text-slate-400">No categories</p>}
            {q.trim() && !exact && (
              <button onClick={() => { onAddNew(q.trim()); pick(q.trim()); }} className="w-full text-left px-3 py-2.5 text-sm text-teal-700 font-medium bg-teal-50/50">
                + Add "{q.trim()}" as new category
              </button>
            )}
            {q.trim() && exact && filtered.length === 0 && (
              <button onClick={() => pick(q.trim())} className="w-full text-left px-3 py-2 text-sm text-slate-700">{q.trim()}</button>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

function AddForm({ onSave, onCancel, rooms, prefill, editTxn, incomeCats, expenseCats, onAddCategory }) {
  const roomNos = rooms.map((r) => r.no);
  const [type, setType] = useState(editTxn?.type || "income");
  const [amount, setAmount] = useState(editTxn ? String(editTxn.amount) : "");
  const [currency, setCurrency] = useState(editTxn?.currency || "LKR");
  const [category, setCategory] = useState(editTxn?.category || prefill?.category || incomeCats[0] || "");
  const [room, setRoom] = useState(editTxn?.room || prefill?.room || "");
  const [roomStatus, setRoomStatus] = useState(() => {
    const rno = editTxn?.room || prefill?.room;
    if (!rno) return "";
    const cur = rooms.find((x) => x.no === rno);
    return cur && cur.status !== "vacant" ? cur.status : "checkin";
  });
  const [guestName, setGuestName] = useState(() => {
    if (editTxn?.guestName) return editTxn.guestName;
    const rno = prefill?.room;
    if (!rno) return "";
    const cur = rooms.find((x) => x.no === rno);
    return cur?.guest || "";
  });
  const [method, setMethod] = useState(() => (editTxn && CARD_PROVIDERS.includes(editTxn.method) ? "Card" : editTxn?.method || "Cash"));
  const [cardProvider, setCardProvider] = useState(() => (editTxn && CARD_PROVIDERS.includes(editTxn.method) ? editTxn.method : CARD_PROVIDERS[0]));
  const [date, setDate] = useState(editTxn?.date || todayStr());
  const [note, setNote] = useState(editTxn?.note || "");
  const [status, setStatus] = useState(editTxn?.status || "paid");
  const [saving, setSaving] = useState(false);
  const cats = type === "income" ? incomeCats : expenseCats;

  const pickRoom = (r) => {
    if (room === r) { setRoom(""); setRoomStatus(""); setGuestName(""); return; }
    setRoom(r);
    const cur = rooms.find((x) => x.no === r);
    setRoomStatus(cur && cur.status !== "vacant" ? cur.status : "checkin");
    setGuestName(cur && cur.status !== "vacant" ? cur.guest || "" : "");
  };
  const switchType = (t) => { setType(t); setCategory(t === "income" ? incomeCats[0] : expenseCats[0]); setRoom(""); setRoomStatus(""); setGuestName(""); };
  const valid = Number(amount) > 0 && category;
  const finalMethod = method === "Card" ? cardProvider : method;

  const submit = async () => {
    setSaving(true);
    await onSave({ id: editTxn?.id, type, amount: Number(amount), currency, category, room: room || undefined, roomStatus: room ? roomStatus : undefined, guestName: room ? guestName : undefined, method: finalMethod, date, note, status });
    setSaving(false);
  };

  return (
    <div className="px-4 pb-6">
      <h2 className="text-lg font-semibold text-slate-800 mt-4 mb-3">{editTxn ? "Edit entry" : "New entry"}</h2>
      {editTxn && (
        <p className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 mb-4">
          Editing a saved entry is recorded and visible to everyone as "Edited," including what changed.
        </p>
      )}
      <div className="flex rounded-lg overflow-hidden border border-slate-200 mb-4">
        <button onClick={() => switchType("income")} className={"flex-1 py-2.5 text-sm font-medium " + (type === "income" ? "bg-emerald-600 text-white" : "bg-white text-slate-600")}>Income</button>
        <button onClick={() => switchType("expense")} className={"flex-1 py-2.5 text-sm font-medium " + (type === "expense" ? "bg-rose-600 text-white" : "bg-white text-slate-600")}>Expense</button>
      </div>

      <label className="block text-xs text-slate-500 mb-1">Currency</label>
      <div className="flex gap-2 mb-2 flex-wrap">
        {CURRENCIES.map((c) => (
          <button key={c} onClick={() => setCurrency(c)} className={"px-3 py-1.5 rounded-full text-xs font-medium border " + (currency === c ? "bg-teal-700 text-white border-teal-700" : "bg-white text-slate-600 border-slate-200")}>{c}</button>
        ))}
      </div>

      <label className="block text-xs text-slate-500 mb-1">Amount ({currency})</label>
      <input type="number" inputMode="decimal" value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="0"
        className="w-full text-2xl font-semibold border border-slate-200 rounded-lg px-3 py-2.5 mb-4 focus:outline-none focus:ring-2 focus:ring-teal-600 tabular-nums" />

      <label className="block text-xs text-slate-500 mb-1">Category — type to search or add new</label>
      <CategoryPicker value={category} onChange={setCategory} options={cats} onAddNew={(name) => onAddCategory(type, name)} />

      {type === "income" && isRoomCat(category) && (
        <>
          <label className="block text-xs text-slate-500 mb-1">Room number</label>
          <div className="flex gap-2 mb-4 flex-wrap">
            {roomNos.map((r) => (
              <button key={r} onClick={() => pickRoom(r)} className={"w-12 py-1.5 rounded-lg text-xs font-semibold border " + (room === r ? "bg-teal-700 text-white border-teal-700" : "bg-white text-slate-600 border-slate-200")}>{r}</button>
            ))}
          </div>
          {room && (
            <>
              <label className="block text-xs text-slate-500 mb-1">Guest status — updates the rooms board</label>
              <div className="flex gap-2 mb-4">
                <button onClick={() => setRoomStatus("checkin")} className={"flex-1 py-2 rounded-lg text-xs font-medium border " + (roomStatus === "checkin" ? "bg-sky-600 text-white border-sky-600" : "bg-white text-slate-600 border-slate-200")}>Check-in</button>
                <button onClick={() => setRoomStatus("staying")} className={"flex-1 py-2 rounded-lg text-xs font-medium border " + (roomStatus === "staying" ? "bg-teal-600 text-white border-teal-600" : "bg-white text-slate-600 border-slate-200")}>Staying</button>
                <button onClick={() => setRoomStatus("checkout")} className={"flex-1 py-2 rounded-lg text-xs font-medium border " + (roomStatus === "checkout" ? "bg-amber-500 text-white border-amber-500" : "bg-white text-slate-600 border-slate-200")}>Check-out</button>
              </div>
              <label className="block text-xs text-slate-500 mb-1">Guest name — tags this payment so same-day turnovers stay distinct</label>
              <input value={guestName} onChange={(e) => setGuestName(e.target.value)} placeholder="Guest name"
                className="w-full border border-slate-200 rounded-lg px-3 py-2.5 mb-4 text-sm" />
            </>
          )}
        </>
      )}

      <label className="block text-xs text-slate-500 mb-1">Payment method</label>
      <div className="flex gap-2 mb-3 flex-wrap">
        {PAY_METHODS.map((m) => (
          <button key={m} onClick={() => setMethod(m)} className={"px-3 py-1.5 rounded-full text-xs font-medium border " + (method === m ? "bg-teal-700 text-white border-teal-700" : "bg-white text-slate-600 border-slate-200")}>{m}</button>
        ))}
      </div>
      {method === "Card" && (
        <>
          <label className="block text-xs text-slate-500 mb-1">Card provider</label>
          <select value={cardProvider} onChange={(e) => setCardProvider(e.target.value)}
            className="w-full border border-slate-200 rounded-lg px-3 py-2.5 mb-4 text-sm bg-white">
            {CARD_PROVIDERS.map((p) => <option key={p} value={p}>{p}</option>)}
          </select>
        </>
      )}

      <label className="block text-xs text-slate-500 mb-1">Payment status</label>
      <div className="flex gap-2 mb-4">
        <button onClick={() => setStatus("paid")} className={"px-3 py-1.5 rounded-full text-xs font-medium border " + (status === "paid" ? "bg-teal-700 text-white border-teal-700" : "bg-white text-slate-600 border-slate-200")}>{type === "income" ? "Received" : "Paid"}</button>
        <button onClick={() => setStatus("pending")} className={"px-3 py-1.5 rounded-full text-xs font-medium border " + (status === "pending" ? "bg-amber-500 text-white border-amber-500" : "bg-white text-slate-600 border-slate-200")}>Pending</button>
      </div>
      {status === "pending" && (
        <p className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 mb-4">
          {type === "income" ? "This will show as a receivable until you mark it paid." : "This will show as a payable until you mark it paid."}
        </p>
      )}

      <label className="block text-xs text-slate-500 mb-1">Date</label>
      <input type="date" value={date} onChange={(e) => setDate(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2.5 mb-4 text-sm bg-white" />

      <label className="block text-xs text-slate-500 mb-1">Note (optional)</label>
      <input value={note} onChange={(e) => setNote(e.target.value)} placeholder="2 nights, incl. breakfast"
        className="w-full border border-slate-200 rounded-lg px-3 py-2.5 mb-6 text-sm" />

      <div className="flex gap-3">
        <button onClick={onCancel} className="flex-1 py-3 rounded-lg border border-slate-200 text-sm font-medium text-slate-600">Cancel</button>
        <button disabled={!valid || saving} onClick={submit} className={"flex-1 py-3 rounded-lg text-sm font-semibold text-white " + (valid && !saving ? "bg-teal-700" : "bg-slate-300")}>
          {saving ? "Saving…" : editTxn ? "Save changes" : "Save entry"}
        </button>
      </div>
    </div>
  );
}

function RoomEditor({ room, onSave, onAddIncome, onClose }) {
  const [status, setStatus] = useState(room.status);
  const [guest, setGuest] = useState(room.guest || "");
  const [out, setOut] = useState(room.out || "");
  useEffect(() => { setStatus(room.status); setGuest(room.guest || ""); setOut(room.out || ""); }, [room.no]);
  const save = () => onSave({ ...room, status, guest: status === "vacant" ? "" : guest, out: status === "vacant" ? "" : out });
  return (
    <div className="bg-white rounded-xl border border-teal-300 p-4 mt-3">
      <div className="flex items-center justify-between mb-3">
        <h2 className="text-sm font-semibold text-slate-800">Room {room.no}</h2>
        <button onClick={onClose} className="text-slate-400 text-xs" aria-label="Close editor">✕</button>
      </div>
      <label className="block text-xs text-slate-500 mb-1">Status</label>
      <div className="flex gap-2 mb-3 flex-wrap">
        {ROOM_STATUSES.map((s) => (
          <button key={s.id} onClick={() => setStatus(s.id)} className={"px-3 py-1.5 rounded-full text-xs font-medium border " + (status === s.id ? "bg-teal-700 text-white border-teal-700" : "bg-white text-slate-600 border-slate-200")}>{s.label}</button>
        ))}
      </div>
      {status !== "vacant" && (
        <>
          <label className="block text-xs text-slate-500 mb-1">Guest name</label>
          <input value={guest} onChange={(e) => setGuest(e.target.value)} placeholder="Guest name" className="w-full border border-slate-200 rounded-lg px-3 py-2 mb-3 text-sm" />
          <label className="block text-xs text-slate-500 mb-1">Check-out date</label>
          <input type="date" value={out} onChange={(e) => setOut(e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 mb-3 text-sm bg-white" />
        </>
      )}
      <div className="flex gap-2">
        <button onClick={save} className="flex-1 py-2.5 rounded-lg bg-teal-700 text-white text-sm font-semibold">Save</button>
        <button onClick={() => onAddIncome(room.no)} className="flex-1 py-2.5 rounded-lg border border-teal-300 text-teal-700 text-sm font-semibold">Add room income</button>
      </div>
    </div>
  );
}

function DrillDown({ label, txns }) {
  return (
    <div className="mt-2 mb-3 bg-slate-50 border border-slate-200 rounded-lg p-3">
      <p className="text-[11px] font-semibold text-slate-500 uppercase tracking-wide mb-2">{label} — {txns.length} entr{txns.length === 1 ? "y" : "ies"}</p>
      {txns.length === 0 ? (
        <p className="text-xs text-slate-400">No entries found</p>
      ) : txns.map((t) => (
        <div key={t.id} className="flex justify-between items-start text-xs py-1.5 border-b border-slate-100 last:border-0 gap-2">
          <div className="min-w-0">
            <div className="text-slate-700 truncate">{t.note || t.category}{t.room ? " · Rm " + t.room : ""}</div>
            <div className="text-slate-400">{t.date} · {t.method}</div>
          </div>
          <div className="tabular-nums font-medium text-slate-800 shrink-0">{fmt(t.amount)}</div>
        </div>
      ))}
    </div>
  );
}

const REQUIRED_COLS = ["type", "amount", "currency", "category", "room", "guest_event", "guest_name", "method", "txn_date", "note", "status", "entered_by"];

function normalizeDate(raw) {
  if (!raw) return null;
  const s = raw.trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s; // already correct
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // DD/MM/YYYY (Excel default after re-save)
  if (m) {
    const [, d, mo, y] = m;
    return `${y}-${mo.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }
  m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/); // DD-MM-YYYY
  if (m) {
    const [, d, mo, y] = m;
    return `${y}-${mo.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }
  return null;
}

function ImportCSV({ incomeCats, expenseCats, onDone }) {
  const [raw, setRaw] = useState("");
  const [parsed, setParsed] = useState(null); // { rows, errors, minDate, maxDate, byCat, unknownCats }
  const [confirmed, setConfirmed] = useState(false);
  const [isExcelFile, setIsExcelFile] = useState(false);
  const [importing, setImporting] = useState(false);
  const [result, setResult] = useState(null);
  const fileRef = useRef(null);
  const [recentBatches, setRecentBatches] = useState([]);
  const [undoing, setUndoing] = useState(null);

  const loadRecentBatches = useCallback(async () => {
    const { data, error } = await supabase
      .from("transactions")
      .select("import_batch, txn_date, amount")
      .not("import_batch", "is", null)
      .order("created_at", { ascending: false })
      .limit(1000);
    if (error || !data) return;
    const groups = {};
    for (const row of data) {
      if (!groups[row.import_batch]) groups[row.import_batch] = { batch: row.import_batch, count: 0, minDate: row.txn_date, maxDate: row.txn_date };
      const g = groups[row.import_batch];
      g.count++;
      if (row.txn_date < g.minDate) g.minDate = row.txn_date;
      if (row.txn_date > g.maxDate) g.maxDate = row.txn_date;
    }
    setRecentBatches(Object.values(groups).sort((a, b) => (a.batch < b.batch ? 1 : -1)).slice(0, 10));
  }, []);

  useEffect(() => { loadRecentBatches(); }, [loadRecentBatches]);


  const analyzeRows = (data, fields) => {
    setResult(null);
    const cleanFields = (fields || []).map((f) => (f || "").trim());
    const cleanedData = data.map((row) => {
      const out = {};
      for (const [k, v] of Object.entries(row)) out[(k || "").trim()] = v;
      return out;
    });
    const missing = REQUIRED_COLS.filter((c) => !cleanFields.includes(c));
    if (missing.length) {
      const found = cleanFields.length ? cleanFields.join(", ") : "(none — the file may be empty, or the header row isn't the first row)";
      setParsed({ rows: [], errors: [`Missing required column(s): ${missing.join(", ")}`, `Columns found in your file: ${found}`], minDate: null, maxDate: null, byCat: [], unknownCats: [] });
      return;
    }
    const rows = cleanedData.filter((r) => r.type && r.amount && r.txn_date);
    const errors = [];
    const allCats = new Set([...incomeCats, ...expenseCats]);
    const unknownCats = new Set();
    let minDate = null, maxDate = null;
    const catTotals = {};
    rows.forEach((r, i) => {
      if (!["income", "expense"].includes(r.type)) errors.push(`Row ${i + 1}: type must be 'income' or 'expense', got '${r.type}'`);
      if (isNaN(Number(r.amount)) || Number(r.amount) <= 0) errors.push(`Row ${i + 1}: invalid amount '${r.amount}'`);
      const normalized = normalizeDate(r.txn_date);
      if (!normalized) errors.push(`Row ${i + 1}: couldn't understand date '${r.txn_date}' (expected YYYY-MM-DD, DD/MM/YYYY, or DD-MM-YYYY)`);
      else r.txn_date = normalized;
      if (!allCats.has(r.category)) unknownCats.add(r.category);
      if (r.txn_date && (!minDate || r.txn_date < minDate)) minDate = r.txn_date;
      if (r.txn_date && (!maxDate || r.txn_date > maxDate)) maxDate = r.txn_date;
      const key = r.type + "|" + r.category;
      catTotals[key] = (catTotals[key] || 0) + (Number(r.amount) || 0);
    });
    const byCat = Object.entries(catTotals).map(([k, v]) => { const [type, cat] = k.split("|"); return { type, cat, total: v }; }).sort((a, b) => b.total - a.total);
    setParsed({ rows, errors: errors.slice(0, 20), errorCount: errors.length, minDate, maxDate, byCat, unknownCats: [...unknownCats] });
    setConfirmed(false);
  };

  const analyze = (text) => {
    const res = Papa.parse(text.trim(), { header: true, skipEmptyLines: true });
    analyzeRows(res.data, res.meta.fields);
  };

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const isExcel = /\.(xlsx|xls)$/i.test(file.name);
    setIsExcelFile(isExcel);
    if (isExcel) {
      setRaw(`[Excel file: ${file.name} — preview generated automatically below]`);
      const reader = new FileReader();
      reader.onload = (ev) => {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, { raw: false, defval: "" });
        const fields = data.length ? Object.keys(data[0]) : [];
        analyzeRows(data, fields);
      };
      reader.readAsArrayBuffer(file);
    } else {
      const reader = new FileReader();
      reader.onload = (ev) => { setRaw(ev.target.result); analyze(ev.target.result); };
      reader.readAsText(file);
    }
  };

  const doImport = async () => {
    if (!parsed || parsed.rows.length === 0) return;
    setImporting(true);
    const batchId = "import_" + Date.now() + "_" + Math.random().toString(36).slice(2, 8);
    const batchSize = 200;
    let inserted = 0;
    const errors = [];
    for (let i = 0; i < parsed.rows.length; i += batchSize) {
      const batch = parsed.rows.slice(i, i + batchSize).map((r) => ({
        type: r.type,
        amount: Number(r.amount),
        currency: r.currency || "LKR",
        category: r.category,
        room: r.room || null,
        guest_event: r.guest_event || null,
        guest_name: r.guest_name || null,
        method: r.method || "Cash",
        txn_date: r.txn_date,
        note: r.note || null,
        status: r.status || "paid",
        entered_by: r.entered_by || "admin",
        import_batch: batchId,
      }));
      const { error } = await supabase.from("transactions").insert(batch);
      if (error) errors.push(error.message);
      else inserted += batch.length;
    }
    setImporting(false);
    setResult({ inserted, errors, batchId: errors.length === 0 ? batchId : null, minDate: parsed.minDate, maxDate: parsed.maxDate });
    if (errors.length === 0) {
      setRaw(""); setParsed(null);
      onDone();
      loadRecentBatches();
    }
  };

  const undoBatch = async (batchId) => {
    setUndoing(batchId);
    const { error } = await supabase.from("transactions").delete().eq("import_batch", batchId);
    setUndoing(null);
    if (error) { alert("Error undoing import: " + error.message); return; }
    if (result && result.batchId === batchId) setResult(null);
    onDone();
    loadRecentBatches();
  };

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
      <h2 className="text-sm font-semibold text-slate-800 mb-1">Import transactions from CSV or Excel</h2>
      <p className="text-xs text-slate-500 mb-3">
        Upload a .csv or .xlsx file with columns: type, amount, currency, category, room, guest_event, guest_name, method, txn_date, note, status, entered_by.
        You'll see exactly what date range and totals are about to be added before anything is saved.
      </p>
      <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" onChange={handleFile} className="block w-full text-xs mb-3" />
      <p className="text-[11px] text-slate-400 mb-3">Or paste CSV text directly (not available for Excel):</p>
      <textarea value={raw} onChange={(e) => { setRaw(e.target.value); setIsExcelFile(false); }} rows={4} disabled={isExcelFile}
        placeholder="type,amount,currency,category,room,guest_event,guest_name,method,txn_date,note,status,entered_by&#10;expense,1500,LKR,Kitchen – General,,,,Cash,2026-07-20,Egg house,paid,admin"
        className={"w-full border border-slate-200 rounded-lg px-3 py-2 mb-3 text-xs font-mono " + (isExcelFile ? "bg-slate-50 text-slate-400" : "")} />
      {!isExcelFile && (
        <button onClick={() => analyze(raw)} disabled={!raw.trim()} className={"w-full py-2.5 rounded-lg text-sm font-semibold text-white mb-3 " + (raw.trim() ? "bg-teal-700" : "bg-slate-300")}>
          Preview import
        </button>
      )}

      {parsed && parsed.errors.length > 0 && (
        <div className="bg-rose-50 border border-rose-200 rounded-lg p-3 mb-3">
          <p className="text-xs font-semibold text-rose-700 mb-1">
            {parsed.rows.length === 0 ? "Can't import — fix this first:" : `${parsed.errorCount} row(s) have problems (showing first ${parsed.errors.length}):`}
          </p>
          {parsed.errors.map((e, i) => <p key={i} className="text-[11px] text-rose-600">{e}</p>)}
        </div>
      )}

      {parsed && parsed.rows.length > 0 && (
        <div className="border border-teal-300 rounded-lg p-3 mb-3 bg-teal-50/40">
          <p className="text-sm font-semibold text-slate-800 mb-2">{parsed.rows.length} transactions ready to import</p>
          <div className="bg-teal-700 text-white rounded-lg px-3 py-3 mb-3 text-center">
            <p className="text-[10px] uppercase tracking-wide opacity-80 mb-0.5">This import covers</p>
            <p className="text-base font-bold">
              {parsed.minDate === parsed.maxDate ? parsed.minDate : `${parsed.minDate}  →  ${parsed.maxDate}`}
            </p>
          </div>
          {parsed.unknownCats.length > 0 && (
            <p className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded px-2 py-1.5 mb-2">
              These categories don't exist yet and will be saved as-is (add them in Settings → Categories if you want them selectable for new entries): {parsed.unknownCats.join(", ")}
            </p>
          )}
          <div className="bg-white rounded-lg p-2 mb-3 max-h-40 overflow-y-auto">
            {parsed.byCat.map((c) => (
              <div key={c.type + c.cat} className="flex justify-between text-xs py-0.5">
                <span className={c.type === "income" ? "text-emerald-700" : "text-rose-700"}>{c.cat}</span>
                <span className="tabular-nums font-medium">{fmt(c.total)}</span>
              </div>
            ))}
          </div>
          <label className="flex items-start gap-2 mb-3 cursor-pointer">
            <input type="checkbox" checked={confirmed} onChange={(e) => setConfirmed(e.target.checked)} className="mt-0.5" />
            <span className="text-xs text-slate-700">
              Yes, <span className="font-semibold">{parsed.minDate === parsed.maxDate ? parsed.minDate : `${parsed.minDate} to ${parsed.maxDate}`}</span> is the correct date range for this import.
            </span>
          </label>
          <button onClick={doImport} disabled={!confirmed || importing}
            className={"w-full py-3 rounded-lg text-sm font-semibold text-white " + (confirmed && !importing ? "bg-teal-700" : "bg-slate-300")}>
            {importing ? "Importing…" : `Confirm and import ${parsed.rows.length} rows`}
          </button>
        </div>
      )}

      {result && (
        <div className={"rounded-lg p-3 mb-3 " + (result.errors.length ? "bg-rose-50 border border-rose-200" : "bg-emerald-50 border border-emerald-200")}>
          <p className={"text-sm font-semibold " + (result.errors.length ? "text-rose-700" : "text-emerald-700")}>
            {result.inserted} row(s) imported{result.errors.length ? `, ${result.errors.length} error(s)` : " successfully"}.
          </p>
          {result.errors.map((e, i) => <p key={i} className="text-xs text-rose-600 mt-1">{e}</p>)}
          {result.batchId && (
            <button onClick={() => undoBatch(result.batchId)} disabled={undoing === result.batchId}
              className="mt-2 text-xs font-semibold text-rose-700 border border-rose-300 rounded-lg px-3 py-1.5 bg-white">
              {undoing === result.batchId ? "Undoing…" : `Undo this import (${result.minDate} to ${result.maxDate})`}
            </button>
          )}
        </div>
      )}

      {recentBatches.length > 0 && (
        <div className="border-t border-slate-100 pt-3 mt-1">
          <p className="text-xs font-semibold text-slate-500 uppercase tracking-wide mb-2">Recent imports</p>
          {recentBatches.map((b) => (
            <div key={b.batch} className="flex items-center justify-between text-xs py-1.5 border-b border-slate-50 last:border-0">
              <span className="text-slate-600">{b.minDate === b.maxDate ? b.minDate : `${b.minDate} → ${b.maxDate}`} <span className="text-slate-400">({b.count} rows)</span></span>
              <button onClick={() => undoBatch(b.batch)} disabled={undoing === b.batch}
                className="text-[11px] font-semibold text-rose-600 border border-rose-200 rounded px-2 py-1">
                {undoing === b.batch ? "Undoing…" : "Undo"}
              </button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function BudgetEditor({ budgets, onSave, expenseCats }) {
  const [draft, setDraft] = useState(budgets);
  const [show, setShow] = useState(false);
  useEffect(() => setDraft(budgets), [budgets]);
  const set = (cat, v) => setDraft({ ...draft, [cat]: v === "" ? "" : Number(v) });
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
      <div className="flex items-center justify-between">
        <h2 className="text-sm font-semibold text-slate-800">Monthly budgets</h2>
        <button onClick={() => setShow(!show)} className="text-xs text-teal-700 font-medium">{show ? "Hide" : "Edit"}</button>
      </div>
      <p className="text-xs text-slate-500 mt-1">{show ? "Set a monthly limit per expense category. Leave blank for no limit." : Object.keys(budgets).length + " budget(s) set"}</p>
      {show && (
        <>
          <div className="mt-3">
            {expenseCats.map((c) => (
              <div key={c} className="flex items-center gap-3 mb-2">
                <span className="text-xs text-slate-600 flex-1">{c}</span>
                <input type="number" inputMode="numeric" value={draft[c] ?? ""} onChange={(e) => set(c, e.target.value)} placeholder="No limit"
                  className="w-32 border border-slate-200 rounded-lg px-2 py-1.5 text-sm text-right tabular-nums" />
              </div>
            ))}
          </div>
          <button onClick={() => onSave(Object.fromEntries(Object.entries(draft).filter(([, v]) => v !== "" && v > 0)))}
            className="mt-2 w-full py-2.5 rounded-lg bg-teal-700 text-white text-sm font-semibold">Save budgets</button>
        </>
      )}
    </div>
  );
}

export default function App() {
  const [view, setView] = useState("dashboard");
  const [txns, setTxns] = useState([]);
  const [histories, setHistories] = useState({});
  const [rooms, setRooms] = useState([]);
  const [budgets, setBudgets] = useState({});
  const [incomeCats, setIncomeCats] = useState([]);
  const [expenseCats, setExpenseCats] = useState([]);
  const [adminPin, setAdminPin] = useState("0105");
  const [rates, setRates] = useState({}); // { USD: 300, EUR: 320, GBP: 380 }
  const [rateDraft, setRateDraft] = useState({});
  const [loaded, setLoaded] = useState(false);
  const [loadError, setLoadError] = useState("");
  const [filter, setFilter] = useState("all");
  const [savedMsg, setSavedMsg] = useState("");
  const [selectedRoom, setSelectedRoom] = useState(null);
  const [addPrefill, setAddPrefill] = useState(null);
  const [editingTxn, setEditingTxn] = useState(null);
  const [roomListDraft, setRoomListDraft] = useState("");
  const [reportTab, setReportTab] = useState("daily");
  const [reportDate, setReportDate] = useState(todayStr());
  const [expandedCat, setExpandedCat] = useState(null); // { scope: 'day'|'month'|'year', type, category }
  const [role, setRole] = useState("staff");
  const [pinInput, setPinInput] = useState("");
  const [showUnlock, setShowUnlock] = useState(false);
  const [pinError, setPinError] = useState(false);
  const [newPin, setNewPin] = useState("");

  const flash = (m) => { setSavedMsg(m); setTimeout(() => setSavedMsg(""), 2000); };

  const loadAll = useCallback(async () => {
    setLoadError("");
    const [txnsRes, histRes, roomsRes, catsRes, budgetsRes, settingsRes] = await Promise.all([
      supabase.from("transactions").select("*").order("txn_date", { ascending: false }).order("created_at", { ascending: false }),
      supabase.from("transaction_history").select("*").order("changed_at", { ascending: true }),
      supabase.from("rooms").select("*").order("no", { ascending: true }),
      supabase.from("categories").select("*").order("id", { ascending: true }),
      supabase.from("budgets").select("*"),
      supabase.from("settings").select("*"),
    ]);
    const firstError = [txnsRes, histRes, roomsRes, catsRes, budgetsRes, settingsRes].find((r) => r.error);
    if (firstError) { setLoadError(firstError.error.message); setLoaded(true); return; }

    setTxns(txnsRes.data.map(rowToTxn));
    const hg = {};
    for (const h of histRes.data) (hg[h.transaction_id] ||= []).push(h);
    setHistories(hg);
    setRooms(roomsRes.data.map(rowToRoom));
    setRoomListDraft(roomsRes.data.map((r) => r.no).join(", "));
    setIncomeCats(catsRes.data.filter((c) => c.type === "income").map((c) => c.name));
    setExpenseCats(catsRes.data.filter((c) => c.type === "expense").map((c) => c.name));
    setBudgets(Object.fromEntries(budgetsRes.data.map((b) => [b.category, Number(b.monthly_limit)])));
    const pin = settingsRes.data.find((s) => s.key === "admin_pin");
    if (pin) setAdminPin(pin.value);
    const loadedRates = {};
    for (const cur of RATE_CURRENCIES) {
      const row = settingsRes.data.find((s) => s.key === "fx_" + cur.toLowerCase());
      if (row) loadedRates[cur] = Number(row.value);
    }
    setRates(loadedRates);
    setRateDraft(loadedRates);
    setLoaded(true);
  }, []);

  useEffect(() => { loadAll(); }, [loadAll]);

  // ---- mutations ----
  const tryUnlock = () => {
    if (pinInput === adminPin) { setRole("admin"); setShowUnlock(false); setPinInput(""); setPinError(false); }
    else { setPinError(true); setPinInput(""); }
  };
  const lockToStaff = () => { setRole("staff"); setView("dashboard"); setSelectedRoom(null); };
  const savePin = async () => {
    const p = newPin.trim();
    if (!/^\d{4,6}$/.test(p)) { flash("PIN must be 4–6 digits"); return; }
    const { error } = await supabase.from("settings").upsert({ key: "admin_pin", value: p });
    if (error) { flash("Error: " + error.message); return; }
    setAdminPin(p); setNewPin(""); flash("Admin PIN updated");
  };
  const saveRates = async () => {
    const rows = RATE_CURRENCIES.filter((c) => rateDraft[c] > 0).map((c) => ({ key: "fx_" + c.toLowerCase(), value: String(rateDraft[c]) }));
    if (rows.length) {
      const { error } = await supabase.from("settings").upsert(rows);
      if (error) { flash("Error: " + error.message); return; }
    }
    const next = {};
    for (const c of RATE_CURRENCIES) if (rateDraft[c] > 0) next[c] = Number(rateDraft[c]);
    setRates(next);
    flash("Exchange rates updated");
  };

  const FIELD_LABELS = { amount: "amount", currency: "currency", category: "category", room: "room", method: "method", date: "date", note: "note", status: "payment status" };
  const describeChanges = (before, after) => {
    const parts = [];
    for (const k of Object.keys(FIELD_LABELS)) {
      const b = before[k] ?? ""; const a = after[k] ?? "";
      if (String(b) !== String(a)) {
        const fv = (v) => (k === "amount" ? fmt(Number(v) || 0) : String(v) || "—");
        parts.push(FIELD_LABELS[k] + " " + fv(b) + " → " + fv(a));
      }
    }
    return parts.length ? parts.join("; ") : "no field changes";
  };

  const addTxn = async (t) => {
    const { roomStatus, guestName, id, ...fields } = t;
    const dbFields = { type: fields.type, amount: fields.amount, currency: fields.currency || "LKR", category: fields.category, room: fields.room || null, guest_event: fields.room ? roomStatus : null, guest_name: fields.room ? guestName || null : null, method: fields.method, txn_date: fields.date, note: fields.note || null, status: fields.status };

    if (id) {
      const before = txns.find((x) => x.id === id);
      const summary = describeChanges(before, fields);
      const { error: upErr } = await supabase.from("transactions").update({ ...dbFields, edited: true }).eq("id", id);
      if (upErr) { flash("Error saving: " + upErr.message); return; }
      const { error: histErr } = await supabase.from("transaction_history").insert({ transaction_id: id, changed_by: role, summary });
      if (histErr) { flash("Error logging history: " + histErr.message); }
    } else {
      const { error } = await supabase.from("transactions").insert({ ...dbFields, entered_by: role });
      if (error) { flash("Error saving: " + error.message); return; }
    }

    if (fields.room && roomStatus) {
      const cur = rooms.find((r) => r.no === fields.room);
      await supabase.from("rooms").update({
        status: roomStatus,
        guest: roomStatus === "vacant" ? "" : (guestName || cur?.guest || ""),
        checkout_date: roomStatus === "vacant" ? null : cur?.out || null,
      }).eq("no", fields.room);
    }

    await loadAll();
    setAddPrefill(null); setEditingTxn(null);
    setView(role === "admin" ? "dashboard" : "transactions");
  };

  const deleteTxn = async (id) => {
    const { error } = await supabase.from("transactions").delete().eq("id", id);
    if (error) { flash("Error deleting: " + error.message); return; }
    await loadAll();
  };
  const markPaid = async (id) => {
    const { error } = await supabase.from("transactions").update({ status: "paid" }).eq("id", id);
    if (error) { flash("Error: " + error.message); return; }
    await loadAll();
  };
  const saveRoom = async (updated) => {
    const { error } = await supabase.from("rooms").update({
      status: updated.status, guest: updated.guest || "", checkout_date: updated.out || null,
    }).eq("no", updated.no);
    if (error) { flash("Error: " + error.message); return; }
    await loadAll();
    setSelectedRoom(null);
  };
  const addIncomeForRoom = (no) => { setAddPrefill({ category: "Room – Walk-in", room: no }); setSelectedRoom(null); setView("add"); };
  const saveRoomList = async () => {
    const nos = roomListDraft.split(",").map((s) => s.trim()).filter(Boolean);
    if (nos.length === 0) return;
    const existing = new Set(rooms.map((r) => r.no));
    const toAdd = nos.filter((n) => !existing.has(n));
    const toRemove = [...existing].filter((n) => !nos.includes(n));
    if (toAdd.length) await supabase.from("rooms").insert(toAdd.map((no) => ({ no })));
    if (toRemove.length) await supabase.from("rooms").delete().in("no", toRemove);
    await loadAll();
    flash("Room list saved");
  };
  const addCategory = async (type, name) => {
    const list = type === "income" ? incomeCats : expenseCats;
    if (list.some((c) => c.toLowerCase() === name.toLowerCase())) return;
    const { error } = await supabase.from("categories").insert({ type, name });
    if (error) { flash("Error: " + error.message); return; }
    await loadAll();
  };
  const removeCategory = async (type, name) => {
    const { error } = await supabase.from("categories").delete().match({ type, name });
    if (error) { flash("Error: " + error.message); return; }
    await loadAll();
  };
  const saveBudgets = async (b) => {
    await supabase.from("budgets").delete().neq("category", "___none___");
    const rows = Object.entries(b).map(([category, monthly_limit]) => ({ category, monthly_limit }));
    if (rows.length) await supabase.from("budgets").insert(rows);
    await loadAll();
    flash("Budgets saved");
  };

  // ---- derived data ----
  const today = todayStr();
  const month = today.slice(0, 7);

  const year = today.slice(0, 4);

  const stats = useMemo(() => {
    let ti = 0, te = 0, mi = 0, me = 0, yi = 0, ye = 0, recv = 0, pay = 0;
    let bcomToday = 0, bcomMonth = 0, bcomYear = 0, bcomAll = 0;
    let cardComToday = 0, cardComMonth = 0, cardComYear = 0, cardComAll = 0;
    let onlineComToday = 0, onlineComMonth = 0, onlineComYear = 0, onlineComAll = 0;
    const fx = {}; // { USD: { today:{in,out}, month:{in,out}, year:{in,out}, todayCom, monthCom, yearCom, allCom, catInc:{cat:amt} }, ... }
    for (const t of txns) {
      const v = Number(t.amount) || 0;
      const cur = t.currency || "LKR";
      const paid = t.status !== "pending";
      if (!paid) { if (cur === "LKR") { t.type === "income" ? (recv += v) : (pay += v); } continue; }

      if (cur !== "LKR") {
        if (!fx[cur]) fx[cur] = { today: { in: 0, out: 0 }, month: { in: 0, out: 0 }, year: { in: 0, out: 0 }, todayCom: 0, monthCom: 0, yearCom: 0, allCom: 0, catInc: {} };
        const fxCom = t.type !== "income" ? 0 : CARD_PROVIDERS.includes(t.method) ? v * CARD_COMMISSION_RATE : t.method === "Online" ? v * ONLINE_COMMISSION_RATE : 0;
        if (t.date === today) { t.type === "income" ? (fx[cur].today.in += v) : (fx[cur].today.out += v); fx[cur].todayCom += fxCom; }
        if (t.date && t.date.startsWith(month)) {
          t.type === "income" ? (fx[cur].month.in += v) : (fx[cur].month.out += v);
          fx[cur].monthCom += fxCom;
          if (t.type === "income") fx[cur].catInc[t.category] = (fx[cur].catInc[t.category] || 0) + v;
        }
        if (t.date && t.date.startsWith(year)) { t.type === "income" ? (fx[cur].year.in += v) : (fx[cur].year.out += v); fx[cur].yearCom += fxCom; }
        fx[cur].allCom += fxCom;
        continue;
      }

      const bcom = t.type === "income" && t.category === BCOM_CAT ? v * BCOM_RATE : 0;
      const cardCom = t.type === "income" && CARD_PROVIDERS.includes(t.method) ? v * CARD_COMMISSION_RATE : 0;
      const onlineCom = t.type === "income" && t.method === "Online" ? v * ONLINE_COMMISSION_RATE : 0;
      bcomAll += bcom; cardComAll += cardCom; onlineComAll += onlineCom;
      if (t.date === today) { t.type === "income" ? (ti += v) : (te += v); bcomToday += bcom; cardComToday += cardCom; onlineComToday += onlineCom; }
      if (t.date && t.date.startsWith(month)) { t.type === "income" ? (mi += v) : (me += v); bcomMonth += bcom; cardComMonth += cardCom; onlineComMonth += onlineCom; }
      if (t.date && t.date.startsWith(year)) { t.type === "income" ? (yi += v) : (ye += v); bcomYear += bcom; cardComYear += cardCom; onlineComYear += onlineCom; }
    }
    return { ti, te, mi, me, yi, ye, recv, pay, bcomToday, bcomMonth, bcomYear, bcomAll, cardComToday, cardComMonth, cardComYear, cardComAll, onlineComToday, onlineComMonth, onlineComYear, onlineComAll, fx };
  }, [txns]);

  // rate-converted foreign totals: { today: X, month: X, year: X } in LKR equivalent, only for currencies with a set rate
  const fxConverted = useMemo(() => {
    const out = { today: 0, month: 0, year: 0, missingCur: [] };
    for (const [cur, v] of Object.entries(stats.fx)) {
      const rate = rates[cur];
      if (!rate) { if ((v.today.in || v.month.in || v.year.in)) out.missingCur.push(cur); continue; }
      out.today += (v.today.in - v.today.out - v.todayCom) * rate;
      out.month += (v.month.in - v.month.out - v.monthCom) * rate;
      out.year += (v.year.in - v.year.out - v.yearCom) * rate;
    }
    return out;
  }, [stats.fx, rates]);

  const roomStats = useMemo(() => {
    const occupied = rooms.filter((r) => r.status === "staying" || r.status === "checkin").length;
    const arrivals = rooms.filter((r) => r.status === "checkin").length;
    const departures = rooms.filter((r) => r.status === "checkout").length;
    return { occupied, arrivals, departures, total: rooms.length };
  }, [rooms]);

  const monthExpenseByCat = useMemo(() => {
    const map = {};
    for (const t of txns) {
      if (t.type !== "expense" || t.status === "pending") continue;
      if ((t.currency || "LKR") !== "LKR") continue;
      if (!t.date || !t.date.startsWith(month)) continue;
      map[t.category] = (map[t.category] || 0) + Number(t.amount);
    }
    return map;
  }, [txns]);

  const budgetAlerts = useMemo(() => {
    const alerts = [];
    for (const [cat, limit] of Object.entries(budgets)) {
      const spent = monthExpenseByCat[cat] || 0;
      if (!limit) continue;
      const pct = (spent / limit) * 100;
      if (pct >= 100) alerts.push({ level: "over", text: cat + " over budget — " + fmt(spent) + " of " + fmt(limit) });
      else if (pct >= 80) alerts.push({ level: "warn", text: cat + " at " + Math.round(pct) + "% of budget (" + fmt(spent) + " / " + fmt(limit) + ")" });
    }
    return alerts;
  }, [budgets, monthExpenseByCat]);

  const dayReport = useMemo(() => {
    const inc = {}, exp = {}, methods = {}, fx = {};
    let ti = 0, te = 0, pendCount = 0, bcom = 0, cardCom = 0, onlineCom = 0;
    for (const t of txns) {
      if (t.date !== reportDate) continue;
      const v = Number(t.amount) || 0;
      const cur = t.currency || "LKR";
      if (t.status === "pending") { pendCount++; continue; }
      if (cur !== "LKR") {
        if (!fx[cur]) fx[cur] = { in: 0, out: 0, com: 0, catInc: {} };
        t.type === "income" ? (fx[cur].in += v) : (fx[cur].out += v);
        if (t.type === "income") {
          if (CARD_PROVIDERS.includes(t.method)) fx[cur].com += v * CARD_COMMISSION_RATE;
          if (t.method === "Online") fx[cur].com += v * ONLINE_COMMISSION_RATE;
          fx[cur].catInc[t.category] = (fx[cur].catInc[t.category] || 0) + v;
        }
        continue;
      }
      if (t.type === "income") {
        inc[t.category] = (inc[t.category] || 0) + v; ti += v;
        if (t.category === BCOM_CAT) bcom += v * BCOM_RATE;
        if (CARD_PROVIDERS.includes(t.method)) cardCom += v * CARD_COMMISSION_RATE;
        if (t.method === "Online") onlineCom += v * ONLINE_COMMISSION_RATE;
      }
      else { exp[t.category] = (exp[t.category] || 0) + v; te += v; }
      if (!methods[t.method]) methods[t.method] = { in: 0, out: 0 };
      t.type === "income" ? (methods[t.method].in += v) : (methods[t.method].out += v);
    }
    const sort = (o) => Object.entries(o).sort((a, b) => b[1] - a[1]);
    return { inc: sort(inc), exp: sort(exp), ti, te, bcom, cardCom, onlineCom, methods: Object.entries(methods), pendCount, fx: Object.entries(fx) };
  }, [txns, reportDate]);

  const dayFxConverted = useMemo(() => {
    let total = 0; const missingCur = [];
    for (const [cur, v] of dayReport.fx) {
      const rate = rates[cur];
      if (!rate) { if (v.in) missingCur.push(cur); continue; }
      total += (v.in - v.out - v.com) * rate;
    }
    return { total, missingCur };
  }, [dayReport.fx, rates]);

  const drillDownTxns = useMemo(() => {
    if (!expandedCat) return [];
    const { scope, type, category } = expandedCat;
    return txns.filter((t) => {
      if (t.type !== type || t.category !== category) return false;
      if (t.status === "pending") return false;
      if ((t.currency || "LKR") !== "LKR") return false;
      if (!t.date) return false;
      if (scope === "day") return t.date === reportDate;
      if (scope === "month") return t.date.startsWith(month);
      if (scope === "year") return t.date.startsWith(year);
      return false;
    }).sort((a, b) => (a.date < b.date ? 1 : -1));
  }, [expandedCat, txns, reportDate, month, year]);

  const monthByCat = useMemo(() => {
    const map = {};
    for (const t of txns) {
      if (!t.date || !t.date.startsWith(month) || t.status === "pending") continue;
      if ((t.currency || "LKR") !== "LKR") continue;
      const k = t.type + "|" + t.category;
      map[k] = (map[k] || 0) + Number(t.amount);
    }
    return Object.entries(map).map(([k, v]) => { const [type, cat] = k.split("|"); return { type, cat, total: v }; }).sort((a, b) => b.total - a.total);
  }, [txns]);

  const yearByCat = useMemo(() => {
    const map = {};
    for (const t of txns) {
      if (!t.date || !t.date.startsWith(year) || t.status === "pending") continue;
      if ((t.currency || "LKR") !== "LKR") continue;
      const k = t.type + "|" + t.category;
      map[k] = (map[k] || 0) + Number(t.amount);
    }
    return Object.entries(map).map(([k, v]) => { const [type, cat] = k.split("|"); return { type, cat, total: v }; }).sort((a, b) => b.total - a.total);
  }, [txns]);

  const monthMethodBreakdown = useMemo(() => {
    const map = {};
    for (const t of txns) {
      if (!t.date || !t.date.startsWith(month) || t.status === "pending") continue;
      if ((t.currency || "LKR") !== "LKR") continue;
      if (!map[t.method]) map[t.method] = { in: 0, out: 0 };
      t.type === "income" ? (map[t.method].in += Number(t.amount)) : (map[t.method].out += Number(t.amount));
    }
    return Object.entries(map).sort((a, b) => b[1].in - a[1].in);
  }, [txns]);

  const yearMethodBreakdown = useMemo(() => {
    const map = {};
    for (const t of txns) {
      if (!t.date || !t.date.startsWith(year) || t.status === "pending") continue;
      if ((t.currency || "LKR") !== "LKR") continue;
      if (!map[t.method]) map[t.method] = { in: 0, out: 0 };
      t.type === "income" ? (map[t.method].in += Number(t.amount)) : (map[t.method].out += Number(t.amount));
    }
    return Object.entries(map).sort((a, b) => b[1].in - a[1].in);
  }, [txns]);

  const filtered = useMemo(() => {
    if (filter === "all") return txns;
    if (filter === "pending") return txns.filter((t) => t.status === "pending");
    return txns.filter((t) => t.type === filter);
  }, [txns, filter]);

  const filteredByDate = useMemo(() => {
    const groups = {};
    for (const t of filtered) {
      const key = t.date || "Undated";
      if (!groups[key]) groups[key] = [];
      groups[key].push(t);
    }
    return Object.entries(groups)
      .sort((a, b) => (a[0] < b[0] ? 1 : -1))
      .map(([date, items]) => {
        let dayIn = 0, dayOut = 0;
        for (const t of items) {
          if (t.status === "pending" || (t.currency || "LKR") !== "LKR") continue;
          t.type === "income" ? (dayIn += Number(t.amount)) : (dayOut += Number(t.amount));
        }
        return { date, items, dayIn, dayOut };
      });
  }, [filtered]);

  const maxCat = Math.max(1, ...monthByCat.map((c) => c.total));
  const yearMaxCat = Math.max(1, ...yearByCat.map((c) => c.total));
  const hasCustomView = incomeCats.length + expenseCats.length > 0;

  const Nav = ({ id, label, icon }) => (
    <button onClick={() => { setView(id); if (id !== "rooms") setSelectedRoom(null); }} className={"flex-1 flex flex-col items-center gap-0.5 py-2 text-[11px] " + (view === id ? "text-teal-700 font-semibold" : "text-slate-400")}>
      <span className="text-base leading-none" aria-hidden="true">{icon}</span>{label}
    </button>
  );

  if (loadError) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center p-6">
        <div className="max-w-sm bg-white border border-rose-200 rounded-xl p-4 text-sm">
          <p className="font-semibold text-rose-700 mb-1">Couldn't reach the database</p>
          <p className="text-slate-600 mb-3">{loadError}</p>
          <p className="text-xs text-slate-400">Check your .env file has the correct VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY, and that the schema SQL has been run.</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 flex justify-center" style={{ fontFamily: "ui-sans-serif, system-ui, sans-serif" }}>
      <div className="w-full max-w-md bg-slate-50 min-h-screen flex flex-col relative">

        <header className="bg-white pt-4 pb-0 sticky top-0 z-10">
          <div className="px-4 flex items-center justify-between pb-3">
            <div>
              <div className="text-base font-bold text-teal-900 tracking-tight">Beach &amp; Bliss</div>
              <div className="text-[11px] uppercase tracking-widest text-teal-700/70">Mirissa · finance</div>
            </div>
            <div className="flex items-center gap-2">
              {role === "staff" ? (
                <button onClick={() => { setShowUnlock(true); setPinError(false); setPinInput(""); }} className="text-[11px] font-medium text-slate-500 border border-slate-200 rounded-full px-3 py-1.5">Staff · unlock admin</button>
              ) : (
                <button onClick={lockToStaff} className="text-[11px] font-medium text-teal-800 bg-teal-50 border border-teal-200 rounded-full px-3 py-1.5">Admin · lock</button>
              )}
              {view !== "add" && (
                <button onClick={() => { setAddPrefill(null); setEditingTxn(null); setView("add"); }} className="bg-teal-700 text-white text-sm font-semibold px-4 py-2 rounded-lg">+ Add</button>
              )}
            </div>
          </div>
          <Wave />
        </header>

        <main className="flex-1 pb-20">
          {!loaded && <div className="p-6 text-sm text-slate-400">Loading…</div>}

          {loaded && showUnlock && role === "staff" && (
            <div className="px-4 pt-4">
              <div className="bg-white rounded-xl border border-teal-300 p-4 mb-3">
                <h2 className="text-sm font-semibold text-slate-800 mb-1">Admin unlock</h2>
                <p className="text-xs text-slate-500 mb-3">Enter the admin PIN to view totals, entries, reports, and settings.</p>
                <input type="password" inputMode="numeric" autoFocus value={pinInput} onChange={(e) => setPinInput(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && tryUnlock()} placeholder="PIN"
                  className="w-full border border-slate-200 rounded-lg px-3 py-2.5 mb-2 text-center text-lg tracking-widest" />
                {pinError && <p className="text-xs text-rose-600 mb-2 text-center">Wrong PIN. Try again.</p>}
                <div className="flex gap-2">
                  <button onClick={() => setShowUnlock(false)} className="flex-1 py-2.5 rounded-lg border border-slate-200 text-sm font-medium text-slate-600">Cancel</button>
                  <button onClick={tryUnlock} className="flex-1 py-2.5 rounded-lg bg-teal-700 text-white text-sm font-semibold">Unlock</button>
                </div>
              </div>
            </div>
          )}

          {loaded && view === "dashboard" && role === "staff" && !showUnlock && (
            <div className="px-4 pt-4">
              <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3 text-center">
                <p className="text-sm text-slate-600 mb-3">Log payments and expenses as they happen.</p>
                <button onClick={() => { setAddPrefill(null); setEditingTxn(null); setView("add"); }} className="w-full py-3.5 rounded-lg bg-teal-700 text-white text-base font-semibold mb-2">+ New entry</button>
                <p className="text-xs text-slate-400">{txns.filter((t) => t.date === today).length} entr{txns.filter((t) => t.date === today).length === 1 ? "y" : "ies"} recorded today</p>
              </div>
              <button onClick={() => setView("rooms")} className="w-full bg-white rounded-xl border border-slate-200 p-3 flex items-center justify-between text-left">
                <div>
                  <div className="text-xs text-slate-500 mb-0.5">Rooms</div>
                  <div className="text-sm font-semibold text-slate-800">{roomStats.occupied}/{roomStats.total} occupied</div>
                </div>
                <div className="flex gap-2 text-[11px]">
                  {roomStats.arrivals > 0 && <span className="bg-sky-50 text-sky-700 border border-sky-200 rounded-full px-2 py-1 font-medium">{roomStats.arrivals} arriving</span>}
                  {roomStats.departures > 0 && <span className="bg-amber-50 text-amber-700 border border-amber-200 rounded-full px-2 py-1 font-medium">{roomStats.departures} checking out</span>}
                  <span className="text-teal-700 font-semibold self-center">→</span>
                </div>
              </button>
            </div>
          )}

          {loaded && view === "dashboard" && role === "admin" && (
            <div className="px-4 pt-4">
              <div className="grid grid-cols-2 gap-3 mb-3">
                <StatCard label="Today's income" value={fmt(stats.ti)} tone="up" />
                <StatCard label="Today's expenses" value={fmt(stats.te)} tone="down" />
                <StatCard label="Net today (after commissions)" value={fmt(stats.ti - stats.te - stats.bcomToday - stats.cardComToday - stats.onlineComToday)} />
                <StatCard label="Month net (after commissions)" value={fmt(stats.mi - stats.me - stats.bcomMonth - stats.cardComMonth - stats.onlineComMonth)} />
              </div>

              {Object.keys(stats.fx).length > 0 && (
                <div className="grid grid-cols-2 gap-3 mb-3">
                  <StatCard label="Month net — LKR only" value={fmt(stats.mi - stats.me - stats.bcomMonth - stats.cardComMonth - stats.onlineComMonth)} />
                  <StatCard label="Month net — LKR + foreign (converted)" value={fxConverted.month || fxConverted.missingCur.length ? fmt(stats.mi - stats.me - stats.bcomMonth - stats.cardComMonth - stats.onlineComMonth + fxConverted.month) : "—"} tone="up" />
                </div>
              )}
              {fxConverted.missingCur.length > 0 && (
                <p className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 mb-3">
                  No exchange rate set for {fxConverted.missingCur.join(", ")} — that income isn't included in the combined figure above. Set rates in Settings.
                </p>
              )}

              {stats.bcomAll > 0 && (
                <div className="bg-white rounded-xl border border-slate-200 p-3 mb-3">
                  <div className="text-xs text-slate-500 mb-1.5">Booking.com commission accrued (18%)</div>
                  <div className="grid grid-cols-3 gap-2 text-center">
                    <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Today</div><div className="text-sm font-semibold tabular-nums text-orange-700">{fmt(stats.bcomToday)}</div></div>
                    <div><div className="text-[10px] uppercase tracking-wide text-slate-400">This month</div><div className="text-sm font-semibold tabular-nums text-orange-700">{fmt(stats.bcomMonth)}</div></div>
                    <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Total accrued</div><div className="text-sm font-semibold tabular-nums text-orange-700">{fmt(stats.bcomAll)}</div></div>
                  </div>
                </div>
              )}

              {stats.cardComAll > 0 && (
                <div className="bg-white rounded-xl border border-slate-200 p-3 mb-3">
                  <div className="text-xs text-slate-500 mb-1.5">Card commission accrued (3% — Com Bank / DFCC / NTB / Global)</div>
                  <div className="grid grid-cols-3 gap-2 text-center">
                    <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Today</div><div className="text-sm font-semibold tabular-nums text-orange-700">{fmt(stats.cardComToday)}</div></div>
                    <div><div className="text-[10px] uppercase tracking-wide text-slate-400">This month</div><div className="text-sm font-semibold tabular-nums text-orange-700">{fmt(stats.cardComMonth)}</div></div>
                    <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Total accrued</div><div className="text-sm font-semibold tabular-nums text-orange-700">{fmt(stats.cardComAll)}</div></div>
                  </div>
                </div>
              )}

              {stats.onlineComAll > 0 && (
                <div className="bg-white rounded-xl border border-slate-200 p-3 mb-3">
                  <div className="text-xs text-slate-500 mb-1.5">Online payment commission accrued (2%)</div>
                  <div className="grid grid-cols-3 gap-2 text-center">
                    <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Today</div><div className="text-sm font-semibold tabular-nums text-orange-700">{fmt(stats.onlineComToday)}</div></div>
                    <div><div className="text-[10px] uppercase tracking-wide text-slate-400">This month</div><div className="text-sm font-semibold tabular-nums text-orange-700">{fmt(stats.onlineComMonth)}</div></div>
                    <div><div className="text-[10px] uppercase tracking-wide text-slate-400">Total accrued</div><div className="text-sm font-semibold tabular-nums text-orange-700">{fmt(stats.onlineComAll)}</div></div>
                  </div>
                </div>
              )}

              {Object.keys(stats.fx).length > 0 && (
                <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
                  <h2 className="text-sm font-semibold text-slate-800 mb-2">Foreign currency income (this month, not converted)</h2>
                  {Object.entries(stats.fx).map(([cur, v]) => (
                    <div key={cur} className="mb-3 last:mb-0 pb-3 last:pb-0 border-b border-slate-50 last:border-0">
                      <div className="flex items-center justify-between text-xs mb-1.5">
                        <span className="font-semibold text-indigo-700">{cur}</span>
                        <span className="text-emerald-700 tabular-nums">+{fmtCur(v.month.in, cur).replace(/^\D+\s?/, "")}</span>
                        <span className="text-rose-700 tabular-nums">−{fmtCur(v.month.out, cur).replace(/^\D+\s?/, "")}</span>
                      </div>
                      {Object.entries(v.catInc).sort((a, b) => b[1] - a[1]).map(([cat, amt]) => (
                        <div key={cat} className="flex justify-between text-[11px] text-slate-500 py-0.5 pl-2">
                          <span>{cat}</span><span className="tabular-nums">{fmtCur(amt, cur).replace(/^\D+\s?/, "")}</span>
                        </div>
                      ))}
                      {v.monthCom > 0 && (
                        <div className="flex items-center justify-between text-[11px] text-orange-700 mt-1">
                          <span>Card/Online commission</span>
                          <span className="tabular-nums">−{fmtCur(v.monthCom, cur).replace(/^\D+\s?/, "")}</span>
                        </div>
                      )}
                      <div className="flex items-center justify-between text-xs font-semibold pt-0.5">
                        <span>Net (after commission)</span>
                        <span className="tabular-nums">{fmtCur(v.month.in - v.month.out - v.monthCom, cur).replace(/^\D+\s?/, "")}</span>
                      </div>
                    </div>
                  ))}
                  <p className="text-[10px] text-slate-400 mt-1.5">Kept separate from LKR totals above — no exchange rate applied.</p>
                </div>
              )}

              <button onClick={() => setView("rooms")} className="w-full bg-white rounded-xl border border-slate-200 p-3 mb-3 flex items-center justify-between text-left">
                <div>
                  <div className="text-xs text-slate-500 mb-0.5">Rooms</div>
                  <div className="text-sm font-semibold text-slate-800">{roomStats.occupied}/{roomStats.total} occupied</div>
                </div>
                <div className="flex gap-2 text-[11px]">
                  {roomStats.arrivals > 0 && <span className="bg-sky-50 text-sky-700 border border-sky-200 rounded-full px-2 py-1 font-medium">{roomStats.arrivals} arriving</span>}
                  {roomStats.departures > 0 && <span className="bg-amber-50 text-amber-700 border border-amber-200 rounded-full px-2 py-1 font-medium">{roomStats.departures} checking out</span>}
                  <span className="text-teal-700 font-semibold self-center">→</span>
                </div>
              </button>

              {(budgetAlerts.length > 0 || stats.recv > 0 || stats.pay > 0) && (
                <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
                  <h2 className="text-sm font-semibold text-slate-800 mb-2">Alerts</h2>
                  {budgetAlerts.map((a, i) => (
                    <div key={i} className={"text-xs rounded-lg px-3 py-2 mb-2 border " + (a.level === "over" ? "bg-rose-50 text-rose-700 border-rose-200" : "bg-amber-50 text-amber-700 border-amber-200")}>{a.text}</div>
                  ))}
                  {stats.recv > 0 && (
                    <button onClick={() => { setFilter("pending"); setView("transactions"); }} className="w-full text-left text-xs rounded-lg px-3 py-2 mb-2 border bg-sky-50 text-sky-700 border-sky-200">Receivables pending: {fmt(stats.recv)} — tap to view</button>
                  )}
                  {stats.pay > 0 && (
                    <button onClick={() => { setFilter("pending"); setView("transactions"); }} className="w-full text-left text-xs rounded-lg px-3 py-2 border bg-amber-50 text-amber-700 border-amber-200">Payables pending: {fmt(stats.pay)} — tap to view</button>
                  )}
                </div>
              )}

              <div className="bg-white rounded-xl border border-slate-200 p-4">
                <div className="flex items-center justify-between mb-1">
                  <h2 className="text-sm font-semibold text-slate-800">Recent activity</h2>
                  <button onClick={() => { setFilter("all"); setView("transactions"); }} className="text-xs text-teal-700 font-medium">View all</button>
                </div>
                {txns.length === 0 ? (
                  <p className="text-sm text-slate-500 py-6 text-center">No entries yet. Add your first income or expense.</p>
                ) : txns.slice(0, 5).map((t) => <TxnRow key={t.id} t={t} onMarkPaid={markPaid} history={histories[t.id]} />)}
              </div>
            </div>
          )}

          {loaded && view === "rooms" && (
            <div className="px-4 pt-4">
              <div className="flex items-center justify-between mb-3">
                <h2 className="text-sm font-semibold text-slate-800">Rooms — {roomStats.occupied}/{roomStats.total} occupied</h2>
                <div className="flex gap-2 text-[10px]">
                  <span className="flex items-center gap-1 text-sky-700"><span className="w-2 h-2 rounded-full bg-sky-400 inline-block" />In</span>
                  <span className="flex items-center gap-1 text-teal-700"><span className="w-2 h-2 rounded-full bg-teal-400 inline-block" />Staying</span>
                  <span className="flex items-center gap-1 text-amber-700"><span className="w-2 h-2 rounded-full bg-amber-400 inline-block" />Out</span>
                </div>
              </div>
              {[...new Set(rooms.map((r) => r.no[0]))].map((floor) => (
                <div key={floor} className="mb-3">
                  <div className="text-[11px] font-semibold uppercase tracking-wider text-slate-400 mb-1.5">Floor {floor}</div>
                  <div className="grid grid-cols-5 gap-1.5">
                    {rooms.filter((r) => r.no[0] === floor).map((r) => (
                      <button key={r.no} onClick={() => setSelectedRoom(r.no)}
                        className={"rounded-lg border p-1.5 text-left min-h-[64px] " + STATUS_STYLE[r.status] + (selectedRoom === r.no ? " ring-2 ring-teal-500" : "")}>
                        <div className="text-sm font-bold leading-none mb-1">{r.no}</div>
                        <div className="text-[9px] font-medium uppercase tracking-wide">{ROOM_STATUSES.find((s) => s.id === r.status)?.label}</div>
                        {r.guest && <div className="text-[9px] truncate mt-0.5 opacity-80">{r.guest}</div>}
                        {r.out && r.status !== "vacant" && <div className="text-[9px] mt-0.5 opacity-60">out {r.out.slice(5)}</div>}
                      </button>
                    ))}
                  </div>
                </div>
              ))}
              {selectedRoom && <RoomEditor room={rooms.find((r) => r.no === selectedRoom)} onSave={saveRoom} onAddIncome={addIncomeForRoom} onClose={() => setSelectedRoom(null)} />}
            </div>
          )}

          {loaded && view === "transactions" && role === "admin" && (
            <div className="px-4 pt-4">
              <div className="flex gap-2 mb-3 flex-wrap">
                {["all", "income", "expense", "pending"].map((f) => (
                  <button key={f} onClick={() => setFilter(f)} className={"px-3 py-1.5 rounded-full text-xs font-medium border capitalize " + (filter === f ? "bg-teal-700 text-white border-teal-700" : "bg-white text-slate-600 border-slate-200")}>{f}</button>
                ))}
              </div>
              {filteredByDate.length === 0 ? (
                <div className="bg-white rounded-xl border border-slate-200 px-4 py-8">
                  <p className="text-sm text-slate-400 text-center">Nothing here yet</p>
                </div>
              ) : filteredByDate.map(({ date, items, dayIn, dayOut }) => (
                <div key={date} className="mb-3">
                  <div className="flex items-center justify-between px-1 mb-1.5">
                    <span className="text-xs font-semibold text-slate-500">{date}</span>
                    <span className="text-[11px] tabular-nums text-slate-400">
                      {dayIn > 0 && <span className="text-emerald-600">+{fmt(dayIn)}</span>}
                      {dayIn > 0 && dayOut > 0 && <span> · </span>}
                      {dayOut > 0 && <span className="text-rose-500">−{fmt(dayOut)}</span>}
                    </span>
                  </div>
                  <div className="bg-white rounded-xl border border-slate-200 px-4 py-1">
                    {items.map((t) => <TxnRow key={t.id} t={t} onDelete={deleteTxn} onMarkPaid={markPaid} history={histories[t.id]} onEdit={(id) => { setEditingTxn(txns.find((x) => x.id === id)); setView("add"); }} />)}
                  </div>
                </div>
              ))}
            </div>
          )}

          {loaded && view === "transactions" && role === "staff" && !showUnlock && (
            <div className="px-4 pt-4">
              <p className="text-xs text-slate-500 mb-3">Entries for today, so you can check what's already been logged and confirm your own entries are correct.</p>
              <div className="bg-white rounded-xl border border-slate-200 px-4 py-1">
                {txns.filter((t) => t.date === today).length === 0 ? (
                  <p className="text-sm text-slate-400 py-8 text-center">Nothing logged today yet</p>
                ) : txns.filter((t) => t.date === today).map((t) => (
                  <TxnRow key={t.id} t={t} onMarkPaid={markPaid} history={histories[t.id]} onEdit={(id) => { setEditingTxn(txns.find((x) => x.id === id)); setView("add"); }} />
                ))}
              </div>
            </div>
          )}

          {loaded && view === "add" && (
            <AddForm onSave={addTxn} onCancel={() => { setAddPrefill(null); setEditingTxn(null); setView(role === "admin" ? "dashboard" : "transactions"); }}
              rooms={rooms} prefill={addPrefill} editTxn={editingTxn} incomeCats={incomeCats} expenseCats={expenseCats} onAddCategory={addCategory} />
          )}

          {loaded && view === "reports" && role === "admin" && (
            <div className="px-4 pt-4">
              <div className="flex rounded-lg overflow-hidden border border-slate-200 mb-3">
                <button onClick={() => setReportTab("daily")} className={"flex-1 py-2 text-sm font-medium " + (reportTab === "daily" ? "bg-teal-700 text-white" : "bg-white text-slate-600")}>Daily</button>
                <button onClick={() => setReportTab("monthly")} className={"flex-1 py-2 text-sm font-medium " + (reportTab === "monthly" ? "bg-teal-700 text-white" : "bg-white text-slate-600")}>Monthly</button>
                <button onClick={() => setReportTab("yearly")} className={"flex-1 py-2 text-sm font-medium " + (reportTab === "yearly" ? "bg-teal-700 text-white" : "bg-white text-slate-600")}>Yearly</button>
              </div>

              {reportTab === "daily" && (
                <>
                  <div className="flex items-center gap-2 mb-3">
                    <input type="date" value={reportDate} onChange={(e) => setReportDate(e.target.value)} className="flex-1 border border-slate-200 rounded-lg px-3 py-2 text-sm bg-white" />
                    {reportDate !== today && <button onClick={() => setReportDate(today)} className="text-xs text-teal-700 font-medium border border-teal-200 rounded-lg px-3 py-2">Today</button>}
                  </div>
                  <div className="grid grid-cols-3 gap-2 mb-3">
                    <StatCard label="Income" value={fmt(dayReport.ti)} tone="up" />
                    <StatCard label="Expenses" value={fmt(dayReport.te)} tone="down" />
                    <StatCard label="Net (after commissions)" value={fmt(dayReport.ti - dayReport.te - dayReport.bcom - dayReport.cardCom - dayReport.onlineCom)} />
                  </div>
                  {dayReport.fx.length > 0 && (
                    <div className="grid grid-cols-2 gap-2 mb-3">
                      <StatCard label="Net — LKR only" value={fmt(dayReport.ti - dayReport.te - dayReport.bcom - dayReport.cardCom - dayReport.onlineCom)} />
                      <StatCard label="Net — LKR + foreign (converted)" value={fmt(dayReport.ti - dayReport.te - dayReport.bcom - dayReport.cardCom - dayReport.onlineCom + dayFxConverted.total)} tone="up" />
                    </div>
                  )}
                  {dayFxConverted.missingCur.length > 0 && (
                    <p className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 mb-3">No exchange rate set for {dayFxConverted.missingCur.join(", ")} — set it in Settings to include it in the combined figure.</p>
                  )}
                  <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
                    <h2 className="text-sm font-semibold text-emerald-700 mb-2">Income by category</h2>
                    {dayReport.inc.length === 0 ? <p className="text-xs text-slate-400 py-2">No income recorded for this day</p> : dayReport.inc.map(([cat, v]) => {
                      const isOpen = expandedCat && expandedCat.scope === "day" && expandedCat.type === "income" && expandedCat.category === cat;
                      return (
                        <div key={cat}>
                          <button onClick={() => setExpandedCat(isOpen ? null : { scope: "day", type: "income", category: cat })}
                            className="w-full flex justify-between text-xs py-1.5 border-b border-slate-50 last:border-0 text-left">
                            <span className="text-slate-600 flex items-center gap-1">{cat} <span className="text-slate-300">{isOpen ? "▲" : "▾"}</span></span>
                            <span className="font-semibold tabular-nums text-emerald-700">{fmt(v)}</span>
                          </button>
                          {isOpen && <DrillDown label={cat} txns={drillDownTxns} />}
                        </div>
                      );
                    })}
                    {dayReport.inc.length > 0 && (
                      <>
                        <div className="flex justify-between text-xs pt-2 mt-1 border-t border-slate-200 font-semibold">
                          <span className="text-slate-800">Total income</span><span className="tabular-nums text-emerald-700">{fmt(dayReport.ti)}</span>
                        </div>
                        {dayReport.bcom > 0 && (
                          <div className="flex justify-between text-xs pt-1.5">
                            <span className="text-orange-700">Less: Booking.com commission (18%)</span><span className="tabular-nums text-orange-700 font-semibold">−{fmt(dayReport.bcom)}</span>
                          </div>
                        )}
                        {dayReport.cardCom > 0 && (
                          <div className="flex justify-between text-xs pt-1.5">
                            <span className="text-orange-700">Less: Card commission (3%)</span><span className="tabular-nums text-orange-700 font-semibold">−{fmt(dayReport.cardCom)}</span>
                          </div>
                        )}
                        {dayReport.onlineCom > 0 && (
                          <div className="flex justify-between text-xs pt-1.5">
                            <span className="text-orange-700">Less: Online payment commission (2%)</span><span className="tabular-nums text-orange-700 font-semibold">−{fmt(dayReport.onlineCom)}</span>
                          </div>
                        )}
                        {(dayReport.bcom > 0 || dayReport.cardCom > 0 || dayReport.onlineCom > 0) && (
                          <div className="flex justify-between text-xs pt-1.5 font-semibold">
                            <span className="text-slate-800">Income after commissions</span><span className="tabular-nums text-emerald-800">{fmt(dayReport.ti - dayReport.bcom - dayReport.cardCom - dayReport.onlineCom)}</span>
                          </div>
                        )}
                      </>
                    )}
                  </div>
                  <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
                    <h2 className="text-sm font-semibold text-rose-700 mb-2">Expenses by category</h2>
                    {dayReport.exp.length === 0 ? <p className="text-xs text-slate-400 py-2">No expenses recorded for this day</p> : dayReport.exp.map(([cat, v]) => {
                      const isOpen = expandedCat && expandedCat.scope === "day" && expandedCat.type === "expense" && expandedCat.category === cat;
                      return (
                        <div key={cat}>
                          <button onClick={() => setExpandedCat(isOpen ? null : { scope: "day", type: "expense", category: cat })}
                            className="w-full flex justify-between text-xs py-1.5 border-b border-slate-50 last:border-0 text-left">
                            <span className="text-slate-600 flex items-center gap-1">{cat} <span className="text-slate-300">{isOpen ? "▲" : "▾"}</span></span>
                            <span className="font-semibold tabular-nums text-rose-700">{fmt(v)}</span>
                          </button>
                          {isOpen && <DrillDown label={cat} txns={drillDownTxns} />}
                        </div>
                      );
                    })}
                    {dayReport.exp.length > 0 && (
                      <div className="flex justify-between text-xs pt-2 mt-1 border-t border-slate-200 font-semibold">
                        <span className="text-slate-800">Total expenses</span><span className="tabular-nums text-rose-700">{fmt(dayReport.te)}</span>
                      </div>
                    )}
                  </div>
                  {dayReport.methods.length > 0 && (
                    <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
                      <h2 className="text-sm font-semibold text-slate-800 mb-2">By payment method</h2>
                      <div className="grid grid-cols-4 text-[10px] font-semibold uppercase tracking-wide text-slate-400 pb-1 border-b border-slate-100">
                        <span>Method</span><span className="text-right">In</span><span className="text-right">Out</span><span className="text-right">Net</span>
                      </div>
                      {dayReport.methods.map(([m, v]) => (
                        <div key={m} className="grid grid-cols-4 text-xs py-1.5 border-b border-slate-50 last:border-0">
                          <span className="text-slate-600">{m}</span>
                          <span className="text-right tabular-nums text-emerald-700">{v.in ? fmt(v.in) : "–"}</span>
                          <span className="text-right tabular-nums text-rose-700">{v.out ? fmt(v.out) : "–"}</span>
                          <span className="text-right tabular-nums font-semibold">{fmt(v.in - v.out)}</span>
                        </div>
                      ))}
                    </div>
                  )}
                  {dayReport.pendCount > 0 && (
                    <p className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 mb-3">{dayReport.pendCount} pending entr{dayReport.pendCount === 1 ? "y" : "ies"} on this date excluded until marked paid.</p>
                  )}
                  {dayReport.fx.length > 0 && (
                    <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
                      <h2 className="text-sm font-semibold text-slate-800 mb-2">Foreign currency income this day (not converted)</h2>
                      {dayReport.fx.map(([cur, v]) => (
                        <div key={cur} className="mb-3 last:mb-0 pb-3 last:pb-0 border-b border-slate-50 last:border-0">
                          <div className="flex items-center justify-between text-xs mb-1.5">
                            <span className="font-semibold text-indigo-700">{cur}</span>
                            <span className="text-emerald-700 tabular-nums">+{fmtCur(v.in, cur).replace(/^\D+\s?/, "")}</span>
                            <span className="text-rose-700 tabular-nums">−{fmtCur(v.out, cur).replace(/^\D+\s?/, "")}</span>
                          </div>
                          {Object.entries(v.catInc).sort((a, b) => b[1] - a[1]).map(([cat, amt]) => (
                            <div key={cat} className="flex justify-between text-[11px] text-slate-500 py-0.5 pl-2">
                              <span>{cat}</span><span className="tabular-nums">{fmtCur(amt, cur).replace(/^\D+\s?/, "")}</span>
                            </div>
                          ))}
                          {v.com > 0 && (
                            <div className="flex items-center justify-between text-[11px] text-orange-700 mt-1">
                              <span>Card/Online commission</span>
                              <span className="tabular-nums">−{fmtCur(v.com, cur).replace(/^\D+\s?/, "")}</span>
                            </div>
                          )}
                          <div className="flex items-center justify-between text-xs font-semibold pt-0.5">
                            <span>Net (after commission)</span>
                            <span className="tabular-nums">{fmtCur(v.in - v.out - v.com, cur).replace(/^\D+\s?/, "")}</span>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </>
              )}

              {reportTab === "monthly" && (
                <>
                  <div className="grid grid-cols-2 gap-3 mb-3">
                    <StatCard label="Month income" value={fmt(stats.mi)} tone="up" />
                    <StatCard label="Month expenses" value={fmt(stats.me)} tone="down" />
                    <StatCard label="B.com commission (18%)" value={"−" + fmt(stats.bcomMonth)} tone="down" />
                    <StatCard label="Card commission (3%)" value={"−" + fmt(stats.cardComMonth)} tone="down" />
                    <StatCard label="Online commission (2%)" value={"−" + fmt(stats.onlineComMonth)} tone="down" />
                  </div>
                  <div className="grid grid-cols-1 gap-3 mb-3">
                    <StatCard label="Month net (after commissions)" value={fmt(stats.mi - stats.me - stats.bcomMonth - stats.cardComMonth - stats.onlineComMonth)} />
                  </div>
                  {Object.keys(stats.fx).length > 0 && (
                    <div className="grid grid-cols-2 gap-2 mb-3">
                      <StatCard label="Net — LKR only" value={fmt(stats.mi - stats.me - stats.bcomMonth - stats.cardComMonth - stats.onlineComMonth)} />
                      <StatCard label="Net — LKR + foreign (converted)" value={fmt(stats.mi - stats.me - stats.bcomMonth - stats.cardComMonth - stats.onlineComMonth + fxConverted.month)} tone="up" />
                    </div>
                  )}
                  {fxConverted.missingCur.length > 0 && (
                    <p className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 mb-3">No exchange rate set for {fxConverted.missingCur.join(", ")} — set it in Settings to include it in the combined figure.</p>
                  )}
                  {(() => {
                    const renderCatRow = (c, scope, scaleMax) => {
                      const isOpen = expandedCat && expandedCat.scope === scope && expandedCat.type === c.type && expandedCat.category === c.cat;
                      return (
                        <div key={c.type + c.cat} className="mb-3 last:mb-0">
                          <button onClick={() => setExpandedCat(isOpen ? null : { scope, type: c.type, category: c.cat })} className="w-full text-left">
                            <div className="flex justify-between text-xs mb-1">
                              <span className="text-slate-600 flex items-center gap-1">{c.cat} <span className="text-slate-300">{isOpen ? "▲" : "▾"}</span></span>
                              <span className={"font-semibold tabular-nums " + (c.type === "income" ? "text-emerald-700" : "text-rose-700")}>{fmt(c.total)}</span>
                            </div>
                            <div className="h-1.5 bg-slate-100 rounded-full overflow-hidden">
                              <div className={"h-full rounded-full " + (c.type === "income" ? "bg-emerald-500" : "bg-rose-400")} style={{ width: (c.total / scaleMax) * 100 + "%" }} />
                            </div>
                          </button>
                          {isOpen && <DrillDown label={c.cat} txns={drillDownTxns} />}
                        </div>
                      );
                    };
                    const monthInc = monthByCat.filter((c) => c.type === "income");
                    const monthExp = monthByCat.filter((c) => c.type === "expense");
                    return (
                      <>
                        <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
                          <h2 className="text-sm font-semibold text-emerald-700 mb-2">Income by category</h2>
                          {monthInc.length === 0 ? <p className="text-xs text-slate-400 py-2">No income this month</p> : monthInc.map((c) => renderCatRow(c, "month", maxCat))}
                        </div>
                        <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
                          <h2 className="text-sm font-semibold text-rose-700 mb-2">Expenses by category</h2>
                          {monthExp.length === 0 ? <p className="text-xs text-slate-400 py-2">No expenses this month</p> : monthExp.map((c) => renderCatRow(c, "month", maxCat))}
                        </div>
                      </>
                    );
                  })()}
                  {monthMethodBreakdown.length > 0 && (
                    <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
                      <h2 className="text-sm font-semibold text-slate-800 mb-2">Income by payment method — {month}</h2>
                      <div className="grid grid-cols-3 text-[10px] font-semibold uppercase tracking-wide text-slate-400 pb-1 border-b border-slate-100">
                        <span>Method</span><span className="text-right">Income</span><span className="text-right">Expense</span>
                      </div>
                      {monthMethodBreakdown.map(([m, v]) => (
                        <div key={m} className="grid grid-cols-3 text-xs py-1.5 border-b border-slate-50 last:border-0">
                          <span className="text-slate-700 font-medium">{m}</span>
                          <span className="text-right tabular-nums text-emerald-700">{v.in ? fmt(v.in) : "–"}</span>
                          <span className="text-right tabular-nums text-rose-700">{v.out ? fmt(v.out) : "–"}</span>
                        </div>
                      ))}
                      <div className="grid grid-cols-3 text-xs pt-2 mt-1 border-t border-slate-200 font-semibold">
                        <span className="text-slate-800">Total</span>
                        <span className="text-right tabular-nums text-emerald-700">{fmt(monthMethodBreakdown.reduce((s, [, v]) => s + v.in, 0))}</span>
                        <span className="text-right tabular-nums text-rose-700">{fmt(monthMethodBreakdown.reduce((s, [, v]) => s + v.out, 0))}</span>
                      </div>
                    </div>
                  )}
                  {Object.keys(budgets).length > 0 && (
                    <div className="bg-white rounded-xl border border-slate-200 p-4">
                      <h2 className="text-sm font-semibold text-slate-800 mb-3">Budget usage</h2>
                      {Object.entries(budgets).map(([cat, limit]) => {
                        const spent = monthExpenseByCat[cat] || 0;
                        const pct = Math.min(100, (spent / limit) * 100);
                        const over = spent > limit;
                        return (
                          <div key={cat} className="mb-3 last:mb-0">
                            <div className="flex justify-between text-xs mb-1">
                              <span className="text-slate-600">{cat}</span>
                              <span className={"font-semibold tabular-nums " + (over ? "text-rose-700" : "text-slate-700")}>{fmt(spent)} / {fmt(limit)}</span>
                            </div>
                            <div className="h-1.5 bg-slate-100 rounded-full overflow-hidden">
                              <div className={"h-full rounded-full " + (over ? "bg-rose-500" : pct >= 80 ? "bg-amber-400" : "bg-teal-500")} style={{ width: pct + "%" }} />
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </>
              )}

              {reportTab === "yearly" && (
                <>
                  <p className="text-xs text-slate-500 mb-3">Showing totals for {year}.</p>
                  <div className="grid grid-cols-2 gap-3 mb-3">
                    <StatCard label="Year income (LKR)" value={fmt(stats.yi)} tone="up" />
                    <StatCard label="Year expenses (LKR)" value={fmt(stats.ye)} tone="down" />
                    <StatCard label="B.com commission (18%)" value={"−" + fmt(stats.bcomYear)} tone="down" />
                    <StatCard label="Card + Online commission" value={"−" + fmt(stats.cardComYear + stats.onlineComYear)} tone="down" />
                  </div>
                  <div className="grid grid-cols-1 gap-3 mb-3">
                    <StatCard label="Year net — LKR only (after commissions)" value={fmt(stats.yi - stats.ye - stats.bcomYear - stats.cardComYear - stats.onlineComYear)} />
                  </div>
                  {(() => {
                    const renderCatRow = (c, scope, scaleMax) => {
                      const isOpen = expandedCat && expandedCat.scope === scope && expandedCat.type === c.type && expandedCat.category === c.cat;
                      return (
                        <div key={c.type + c.cat} className="mb-3 last:mb-0">
                          <button onClick={() => setExpandedCat(isOpen ? null : { scope, type: c.type, category: c.cat })} className="w-full text-left">
                            <div className="flex justify-between text-xs mb-1">
                              <span className="text-slate-600 flex items-center gap-1">{c.cat} <span className="text-slate-300">{isOpen ? "▲" : "▾"}</span></span>
                              <span className={"font-semibold tabular-nums " + (c.type === "income" ? "text-emerald-700" : "text-rose-700")}>{fmt(c.total)}</span>
                            </div>
                            <div className="h-1.5 bg-slate-100 rounded-full overflow-hidden">
                              <div className={"h-full rounded-full " + (c.type === "income" ? "bg-emerald-500" : "bg-rose-400")} style={{ width: (c.total / scaleMax) * 100 + "%" }} />
                            </div>
                          </button>
                          {isOpen && <DrillDown label={c.cat} txns={drillDownTxns} />}
                        </div>
                      );
                    };
                    const yearInc = yearByCat.filter((c) => c.type === "income");
                    const yearExp = yearByCat.filter((c) => c.type === "expense");
                    return (
                      <>
                        <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
                          <h2 className="text-sm font-semibold text-emerald-700 mb-2">Income by category — {year}</h2>
                          {yearInc.length === 0 ? <p className="text-xs text-slate-400 py-2">No income this year</p> : yearInc.map((c) => renderCatRow(c, "year", yearMaxCat))}
                        </div>
                        <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
                          <h2 className="text-sm font-semibold text-rose-700 mb-2">Expenses by category — {year}</h2>
                          {yearExp.length === 0 ? <p className="text-xs text-slate-400 py-2">No expenses this year</p> : yearExp.map((c) => renderCatRow(c, "year", yearMaxCat))}
                        </div>
                      </>
                    );
                  })()}
                  {Object.keys(stats.fx).length > 0 && (
                    <>
                      <div className="grid grid-cols-1 gap-3 mb-3">
                        <StatCard label="Year net — LKR + foreign (converted)" value={fmt(stats.yi - stats.ye - stats.bcomYear - stats.cardComYear - stats.onlineComYear + fxConverted.year)} tone="up" />
                      </div>
                      {fxConverted.missingCur.length > 0 && (
                        <p className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 mb-3">No exchange rate set for {fxConverted.missingCur.join(", ")} — set it in Settings to include it in the combined figure.</p>
                      )}
                      <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
                        <h2 className="text-sm font-semibold text-slate-800 mb-2">Foreign currency income — {year} (not converted)</h2>
                        {Object.entries(stats.fx).map(([cur, v]) => (
                          <div key={cur} className="mb-2 last:mb-0 pb-2 last:pb-0 border-b border-slate-50 last:border-0">
                            <div className="flex items-center justify-between text-xs">
                              <span className="font-semibold text-indigo-700">{cur}</span>
                              <span className="text-emerald-700 tabular-nums">+{fmtCur(v.year.in, cur).replace(/^\D+\s?/, "")}</span>
                              <span className="text-rose-700 tabular-nums">−{fmtCur(v.year.out, cur).replace(/^\D+\s?/, "")}</span>
                              {rates[cur] && <span className="text-slate-400 tabular-nums">≈ {fmt((v.year.in - v.year.out - v.yearCom) * rates[cur])}</span>}
                            </div>
                          </div>
                        ))}
                      </div>
                    </>
                  )}
                  {yearMethodBreakdown.length > 0 && (
                    <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
                      <h2 className="text-sm font-semibold text-slate-800 mb-2">Income by payment method — {year}</h2>
                      <div className="grid grid-cols-3 text-[10px] font-semibold uppercase tracking-wide text-slate-400 pb-1 border-b border-slate-100">
                        <span>Method</span><span className="text-right">Income</span><span className="text-right">Expense</span>
                      </div>
                      {yearMethodBreakdown.map(([m, v]) => (
                        <div key={m} className="grid grid-cols-3 text-xs py-1.5 border-b border-slate-50 last:border-0">
                          <span className="text-slate-700 font-medium">{m}</span>
                          <span className="text-right tabular-nums text-emerald-700">{v.in ? fmt(v.in) : "–"}</span>
                          <span className="text-right tabular-nums text-rose-700">{v.out ? fmt(v.out) : "–"}</span>
                        </div>
                      ))}
                      <div className="grid grid-cols-3 text-xs pt-2 mt-1 border-t border-slate-200 font-semibold">
                        <span className="text-slate-800">Total</span>
                        <span className="text-right tabular-nums text-emerald-700">{fmt(yearMethodBreakdown.reduce((s, [, v]) => s + v.in, 0))}</span>
                        <span className="text-right tabular-nums text-rose-700">{fmt(yearMethodBreakdown.reduce((s, [, v]) => s + v.out, 0))}</span>
                      </div>
                    </div>
                  )}
                </>
              )}
            </div>
          )}

          {loaded && view === "settings" && role === "admin" && (
            <div className="px-4 pt-4">
              {savedMsg && <div className="text-xs text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-lg px-3 py-2 mb-3">{savedMsg}</div>}
              <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
                <h2 className="text-sm font-semibold text-slate-800 mb-1">Admin PIN</h2>
                <p className="text-xs text-slate-500 mb-2">Staff can only add entries and manage rooms. This PIN unlocks totals, reports, and settings.</p>
                <div className="flex gap-2">
                  <input type="password" inputMode="numeric" value={newPin} onChange={(e) => setNewPin(e.target.value)} placeholder="New PIN (4–6 digits)" className="flex-1 border border-slate-200 rounded-lg px-3 py-2 text-sm" />
                  <button onClick={savePin} className="px-4 py-2 rounded-lg bg-teal-700 text-white text-sm font-semibold">Save</button>
                </div>
              </div>
              <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
                <h2 className="text-sm font-semibold text-slate-800 mb-1">Rooms</h2>
                <p className="text-xs text-slate-500 mb-2">List your room numbers separated by commas.</p>
                <input value={roomListDraft} onChange={(e) => setRoomListDraft(e.target.value)} placeholder="101, 102, 201, 202" className="w-full border border-slate-200 rounded-lg px-3 py-2 mb-2 text-sm" />
                <button onClick={saveRoomList} className="w-full py-2.5 rounded-lg bg-teal-700 text-white text-sm font-semibold">Save room list</button>
              </div>
              {hasCustomView && (
                <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
                  <h2 className="text-sm font-semibold text-slate-800 mb-1">Categories</h2>
                  <p className="text-xs text-slate-500 mb-2">Remove ones you no longer need — existing entries keep their category label.</p>
                  <div className="max-h-60 overflow-y-auto">
                    {[...incomeCats.map((c) => ["income", c]), ...expenseCats.map((c) => ["expense", c])].map(([type, c]) => (
                      <div key={type + c} className="flex items-center justify-between py-1.5 border-b border-slate-100 last:border-0">
                        <span className="text-xs text-slate-700">{c} <span className="text-slate-400">({type})</span></span>
                        <button onClick={() => removeCategory(type, c)} className="text-slate-300 hover:text-rose-500 text-xs" aria-label="Remove category">✕</button>
                      </div>
                    ))}
                  </div>
                </div>
              )}
              <div className="bg-white rounded-xl border border-slate-200 p-4 mb-3">
                <h2 className="text-sm font-semibold text-slate-800 mb-1">Exchange rates</h2>
                <p className="text-xs text-slate-500 mb-2">How many LKR equal 1 unit of each currency today. Update daily as needed — used to show a combined LKR + foreign net on the dashboard and reports.</p>
                {RATE_CURRENCIES.map((cur) => (
                  <div key={cur} className="flex items-center gap-2 mb-2">
                    <span className="text-xs text-slate-600 w-28">1 {cur} =</span>
                    <input type="number" inputMode="decimal" value={rateDraft[cur] ?? ""} onChange={(e) => setRateDraft({ ...rateDraft, [cur]: e.target.value === "" ? "" : Number(e.target.value) })}
                      placeholder="e.g. 300" className="flex-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm tabular-nums" />
                    <span className="text-xs text-slate-400">LKR</span>
                  </div>
                ))}
                <button onClick={saveRates} className="mt-1 w-full py-2.5 rounded-lg bg-teal-700 text-white text-sm font-semibold">Save exchange rates</button>
              </div>
              <ImportCSV incomeCats={incomeCats} expenseCats={expenseCats} onDone={loadAll} />
              <BudgetEditor budgets={budgets} onSave={saveBudgets} expenseCats={expenseCats} />
            </div>
          )}
        </main>

        <nav className="fixed bottom-0 w-full max-w-md bg-white border-t border-slate-200 flex">
          <Nav id="dashboard" label="Home" icon="⌂" />
          <Nav id="rooms" label="Rooms" icon="◫" />
          <Nav id="add" label="Add" icon="+" />
          <Nav id="transactions" label="Entries" icon="≡" />
          {role === "admin" && (
            <>
              <Nav id="reports" label="Reports" icon="▤" />
              <Nav id="settings" label="Settings" icon="⚙" />
            </>
          )}
        </nav>
      </div>
    </div>
  );
}

"""
Chiusure Cassa Bot — Railway deployment
Riceve messaggi di chiusura dal gruppo Telegram, aggiorna i dati,
elimina i messaggi e invia riepilogo giornaliero a Ricardo.
"""
import os, json, re, logging, io
from flask import Flask, request, jsonify
from datetime import datetime, date, timedelta
import requests
from apscheduler.schedulers.background import BackgroundScheduler
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── Configurazione ────────────────────────────────────────────────────────
TOKEN      = os.environ["BOT_TOKEN"]
GROUP_ID   = int(os.environ.get("GROUP_CHAT_ID",  "-5249608111"))
RICARDO_ID = int(os.environ.get("RICARDO_CHAT_ID", "5590933344"))
DATA_FILE  = "data.json"
API        = f"https://api.telegram.org/bot{TOKEN}"

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

app = Flask(__name__)

# ── Storage in memoria (+ file JSON per persistenza tra restart) ──────────
def load_data():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE) as f:
                return json.load(f)
        except Exception:
            pass
    return {"closures": []}

def save_data(data):
    with open(DATA_FILE, "w") as f:
        json.dump(data, f, default=str, ensure_ascii=False, indent=2)

db = load_data()

# ── Helpers Telegram ──────────────────────────────────────────────────────
def tg(method, **kwargs):
    try:
        r = requests.post(f"{API}/{method}", json=kwargs, timeout=10)
        return r.json()
    except Exception as e:
        log.error(f"Telegram error ({method}): {e}")
        return {}

def send(chat_id, text, parse_mode="HTML"):
    return tg("sendMessage", chat_id=chat_id, text=text, parse_mode=parse_mode)

def delete_msg(chat_id, message_id):
    return tg("deleteMessage", chat_id=chat_id, message_id=message_id)

def send_document(chat_id, filename, content_bytes, caption=""):
    try:
        r = requests.post(
            f"{API}/sendDocument",
            data={"chat_id": chat_id, "caption": caption},
            files={"document": (filename, content_bytes,
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            timeout=30
        )
        return r.json()
    except Exception as e:
        log.error(f"send_document error: {e}")
        return {}

# ── Parser numeri (formato italiano) ─────────────────────────────────────
def parse_num(s):
    s = str(s).strip().replace(" ", "")
    if not s or s == "-": return None
    has_k = s.lower().endswith("k")
    if has_k: s = s[:-1]
    if "," in s:
        parts = s.split(",")
        try:
            v = float(parts[0].replace(".", "") + "." + parts[1])
            return v * 1000 if has_k else v
        except: return None
    if "." in s:
        last = s.split(".")[-1]
        try:
            v = float(s) if len(last) <= 2 else float(s.replace(".", ""))
            return v * 1000 if has_k else v
        except: return None
    try:
        v = float(s)
        return v * 1000 if has_k else v
    except: return None

def extract_cash(text):
    found = {}
    deductions = []
    for line in text.split("\n"):
        line = line.strip()
        if not line: continue
        m = re.match(r"^f2[\s\.]+([\d\.,]+)\s*k?", line, re.I)
        if m:
            has_k = bool(re.search(r"k", line[m.start():m.end()+2], re.I))
            v = parse_num(m.group(1) + ("k" if has_k else ""))
            if v and v > 0: found["f2"] = v; continue
        m = re.match(r"^(?:tot\.?|t\.)\s*([\d\.,]+)", line, re.I)
        if m:
            v = parse_num(m.group(1))
            if v and v > 5: found["tot"] = v; continue
        m = re.match(r"^f1?[\s\.]+([\d\.,]+)", line, re.I)
        if m:
            v = parse_num(m.group(1))
            if v and v > 5: found["f1"] = v; continue
        m = re.match(r"^b[\s\.]+([\d\.,]+)", line, re.I)
        if m:
            v = parse_num(m.group(1))
            if v is not None: found["b"] = v; continue
        if re.match(r"^-\s*[\d]", line):
            deductions.append(line)
    if deductions:
        found["deductions"] = " | ".join(deductions)
    return found

def detect_location(text):
    for line in text.split("\n")[:4]:
        l = line.strip().lower()
        if "bologna" in l: return "Bologna"
        if "padova"  in l: return "Padova"
    return None

def is_correction(text):
    return any(k in text.lower() for k in
               ["correzione", "errore", "sbagliato", "corretto",
                "modifica", "rettifica"])

def effective_date(timestamp):
    dt = datetime.fromtimestamp(timestamp)
    return (dt - timedelta(days=1)).date() if dt.hour < 6 else dt.date()

def fmt_eur(v):
    if v is None: return "–"
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{s} €"

# ── Genera Excel ──────────────────────────────────────────────────────────
def generate_excel(closures):
    wb = Workbook()
    ws = wb.active
    ws.title = "Chiusure"
    thin = Side(style="thin", color="BFBFBF")
    def bdr(): return Border(left=thin, right=thin, top=thin, bottom=thin)
    def cs(cell, bg=None, bold=False, align="center", sz=10):
        cell.font = Font(name="Arial", bold=bold, size=sz)
        cell.alignment = Alignment(horizontal=align, vertical="center")
        if bg: cell.fill = PatternFill("solid", start_color=bg)
        cell.border = bdr()

    headers = ["Data", "Sede", "Totale", "B", "F1 (Fondo Cassa)",
               "F2 (Fondo 2)", "Coperti Pranzo", "Coperti Sera", "Deduzioni"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=ci, value=h)
        cs(c, bg="1F3864", bold=True)
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)

    BG = {"Bologna": ("DEEAF1","C5DCF0"), "Padova": ("E2EFDA","C6E0B4")}
    EUR = '#,##0.00\\ €'

    for ri, rec in enumerate(
        sorted(closures, key=lambda x: (x["date"], x["location"])), 2):
        bg = BG.get(rec["location"], ("FFFFFF","F0F0F0"))[(ri%2)]
        row_vals = [
            datetime.strptime(rec["date"], "%Y-%m-%d").date(),
            rec["location"],
            rec.get("tot"), rec.get("b"), rec.get("f1"), rec.get("f2"),
            None, None,
            rec.get("deductions","")
        ]
        for ci, val in enumerate(row_vals, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cs(cell, bg=bg, align="right" if isinstance(val, float) else
               ("left" if ci in (2, 9) else "center"))
            if ci == 1: cell.number_format = "DD/MM/YYYY"
            elif ci in (3,4,5,6) and isinstance(val, float):
                cell.number_format = EUR

    col_widths = [12,10,14,12,18,14,14,12,45]
    from openpyxl.utils import get_column_letter
    for ci, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ── Job giornaliero (ore 9:00) ────────────────────────────────────────────
def daily_summary():
    log.info("Invio riepilogo giornaliero...")
    yesterday = (date.today() - timedelta(days=1)).isoformat()
    today_recs = [r for r in db["closures"] if r["date"] == yesterday]

    bol = next((r for r in today_recs if r["location"] == "Bologna"), None)
    pad = next((r for r in today_recs if r["location"] == "Padova"),  None)

    lines = [f"📊 <b>Chiusure {date.fromisoformat(yesterday).strftime('%d/%m/%Y')}</b>\n"]

    if bol:
        lines.append("🔵 <b>BOLOGNA</b>")
        lines.append(f"Tot: {fmt_eur(bol.get('tot'))}  |  F1: {fmt_eur(bol.get('f1'))}  |  B: {fmt_eur(bol.get('b'))}")
        if bol.get("f2"): lines.append(f"F2: {fmt_eur(bol['f2'])}")
        if bol.get("deductions"): lines.append(f"📝 {bol['deductions']}")
    else:
        lines.append("🔵 <b>BOLOGNA</b> — nessun dato")

    lines.append("")

    if pad:
        lines.append("🟢 <b>PADOVA</b>")
        lines.append(f"Tot: {fmt_eur(pad.get('tot'))}  |  F1: {fmt_eur(pad.get('f1'))}  |  B: {fmt_eur(pad.get('b'))}")
        if pad.get("f2"): lines.append(f"F2: {fmt_eur(pad['f2'])}")
        if pad.get("deductions"): lines.append(f"📝 {pad['deductions']}")
    else:
        lines.append("🟢 <b>PADOVA</b> — nessun dato")

    tot = (bol.get("tot") or 0) + (pad.get("tot") or 0) if (bol or pad) else 0
    if tot: lines.append(f"\n💰 <b>TOTALE: {fmt_eur(tot)}</b>")

    send(RICARDO_ID, "\n".join(lines))

    # Invia anche il file Excel aggiornato
    if db["closures"]:
        xlsx = generate_excel(db["closures"])
        send_document(RICARDO_ID, "chiusure_cassa.xlsx", xlsx,
                      caption="📎 Excel aggiornato con tutte le chiusure")

    log.info("Riepilogo inviato.")

# ── Webhook handler ───────────────────────────────────────────────────────
@app.route("/webhook", methods=["POST"])
def webhook():
    update = request.get_json(silent=True) or {}
    msg = update.get("message")
    if not msg:
        return jsonify(ok=True)

    chat_id = msg.get("chat", {}).get("id")
    text     = msg.get("text", "")
    msg_id   = msg.get("message_id")
    ts       = msg.get("date", 0)
    sender   = msg.get("from", {}).get("first_name", "?")

    # Comandi da Ricardo in privato
    if chat_id == RICARDO_ID:
        if text.strip() == "/excel":
            if db["closures"]:
                xlsx = generate_excel(db["closures"])
                send_document(RICARDO_ID, "chiusure_cassa.xlsx", xlsx,
                              caption="📎 Tutte le chiusure registrate")
            else:
                send(RICARDO_ID, "Nessuna chiusura registrata.")
        elif text.strip() == "/oggi":
            daily_summary()
        return jsonify(ok=True)

    # Messaggi dal gruppo
    if chat_id != GROUP_ID:
        return jsonify(ok=True)

    if not text:
        return jsonify(ok=True)

    # Correzione?
    if is_correction(text):
        alert = (f"⚠️ <b>Correzione ricevuta</b>\n\n"
                 f"Da: {sender}\n"
                 f"Testo:\n<code>{text}</code>\n\n"
                 f"Verifica e aggiorna l'Excel manualmente se necessario.")
        send(RICARDO_ID, alert)
        return jsonify(ok=True)

    # Chiusura cassa?
    location = detect_location(text)
    if not location:
        return jsonify(ok=True)

    cash = extract_cash(text)
    if not cash:
        return jsonify(ok=True)

    eff_d = effective_date(ts).isoformat()

    # Controlla se già esiste una chiusura per questo giorno/sede
    existing = next((r for r in db["closures"]
                     if r["date"] == eff_d and r["location"] == location), None)
    if existing:
        # Aggiorna con i nuovi dati (più completi)
        if len(cash) >= len({k:v for k,v in existing.items()
                              if k not in ("date","location","sender")}):
            existing.update({**cash, "sender": sender})
            log.info(f"Aggiornata chiusura esistente: {location} {eff_d}")
    else:
        record = {"date": eff_d, "location": location, "sender": sender, **cash}
        db["closures"].append(record)
        log.info(f"Nuova chiusura: {location} {eff_d}")

    save_data(db)

    # Elimina il messaggio dal gruppo
    delete_msg(GROUP_ID, msg_id)

    # Notifica immediata a Ricardo
    d_fmt = date.fromisoformat(eff_d).strftime("%d/%m")
    notif = (f"✅ <b>{location} — {d_fmt}</b>\n"
             f"Tot: {fmt_eur(cash.get('tot'))}  |  "
             f"F1: {fmt_eur(cash.get('f1'))}  |  "
             f"B: {fmt_eur(cash.get('b'))}")
    if cash.get("deductions"):
        notif += f"\n📝 {cash['deductions']}"
    send(RICARDO_ID, notif)

    return jsonify(ok=True)

@app.route("/health")
def health():
    return jsonify(ok=True, closures=len(db["closures"]))

# ── Avvio ─────────────────────────────────────────────────────────────────
scheduler.add_job(daily_summary, "cron", hour=9, minute=0,
                  timezone="Europe/Rome")
scheduler.start()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

import os, json, re, io, logging
from flask import Flask, request, jsonify
from datetime import datetime, date, timedelta
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

app = Flask(__name__)

# Variabili lette dentro le funzioni, mai al livello del modulo
def cfg():
    return {
        "token":      os.environ.get("BOT_TOKEN", ""),
        "group_id":   int(os.environ.get("GROUP_CHAT_ID",  "-5249608111")),
        "ricardo_id": int(os.environ.get("RICARDO_CHAT_ID", "5590933344")),
    }

DATA_FILE = "/tmp/data.json"

def load_db():
    try:
        with open(DATA_FILE) as f:
            return json.load(f)
    except Exception:
        return {"closures": [], "daily_sent": ""}

def save_db(db):
    with open(DATA_FILE, "w") as f:
        json.dump(db, f, default=str, ensure_ascii=False, indent=2)

# ── Telegram ──────────────────────────────────────────────────────────────
def api(method, **kw):
    t = cfg()["token"]
    if not t:
        log.error("BOT_TOKEN non impostato")
        return {}
    try:
        r = requests.post(f"https://api.telegram.org/bot{t}/{method}",
                          json=kw, timeout=15)
        return r.json()
    except Exception as e:
        log.error(f"api {method}: {e}")
        return {}

def send(chat_id, text, parse_mode="HTML"):
    return api("sendMessage", chat_id=chat_id, text=text, parse_mode=parse_mode)

def delete_msg(chat_id, message_id):
    return api("deleteMessage", chat_id=chat_id, message_id=message_id)

def send_file(chat_id, filename, data_bytes, caption=""):
    t = cfg()["token"]
    if not t: return
    try:
        requests.post(
            f"https://api.telegram.org/bot{t}/sendDocument",
            data={"chat_id": chat_id, "caption": caption},
            files={"document": (filename, data_bytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            timeout=30)
    except Exception as e:
        log.error(f"send_file: {e}")

# ── Parser ────────────────────────────────────────────────────────────────
def parse_num(s):
    s = str(s).strip().replace(" ", "")
    if not s or s == "-": return None
    has_k = s.lower().endswith("k")
    if has_k: s = s[:-1]
    if "," in s:
        p = s.split(",")
        try: v = float(p[0].replace(".", "") + "." + p[1]); return v*1000 if has_k else v
        except: return None
    if "." in s:
        last = s.split(".")[-1]
        try:
            v = float(s) if len(last) <= 2 else float(s.replace(".", ""))
            return v*1000 if has_k else v
        except: return None
    try: v = float(s); return v*1000 if has_k else v
    except: return None

def extract_cash(text):
    found, deds = {}, []
    for line in text.split("\n"):
        line = line.strip()
        if not line: continue
        m = re.match(r"^f2[\s\.]+([\d\.,]+)\s*k?", line, re.I)
        if m:
            hk = bool(re.search(r"k", line[m.start():m.end()+2], re.I))
            v = parse_num(m.group(1)+("k" if hk else ""))
            if v: found["f2"] = v; continue
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
        if re.match(r"^-\s*[\d]", line): deds.append(line)
    if deds: found["deductions"] = " | ".join(deds)
    return found

def detect_location(text):
    for line in text.split("\n")[:4]:
        l = line.strip().lower()
        if "bologna" in l: return "Bologna"
        if "padova"  in l: return "Padova"
    return None

def is_correction(text):
    return any(k in text.lower() for k in
               ["correzione","errore","sbagliato","corretto","modifica","rettifica"])

def eff_date(ts):
    dt = datetime.fromtimestamp(ts)
    return (dt - timedelta(days=1)).date() if dt.hour < 6 else dt.date()

def fmt(v):
    if v is None: return "–"
    return f"{v:,.2f}".replace(",","X").replace(".",",").replace("X",".") + " €"

# ── Excel ─────────────────────────────────────────────────────────────────
def make_excel(closures):
    wb = Workbook(); ws = wb.active; ws.title = "Chiusure"
    thin = Side(style="thin", color="BFBFBF")
    def bdr(): return Border(left=thin,right=thin,top=thin,bottom=thin)
    def cs(c, bg=None, bold=False, fc="000000", sz=10, al="center"):
        c.font = Font(name="Arial", bold=bold, color=fc, size=sz)
        c.alignment = Alignment(horizontal=al, vertical="center")
        if bg: c.fill = PatternFill("solid", start_color=bg)
        c.border = bdr()
    hdrs = ["Data","Sede","Totale","B","F1","F2","Coperti Pranzo","Coperti Sera","Deduzioni"]
    for ci, h in enumerate(hdrs, 1):
        c = ws.cell(row=1, column=ci, value=h)
        cs(c, bg="1F3864", bold=True, fc="FFFFFF")
    BG = {"Bologna":["DEEAF1","C5DCF0"], "Padova":["E2EFDA","C6E0B4"]}
    EUR = '#,##0.00\\ €'
    for ri, rec in enumerate(sorted(closures, key=lambda x:(x["date"],x["location"])), 2):
        bg = BG.get(rec["location"],["FFFFFF","F2F2F2"])[ri%2]
        vals = [datetime.strptime(rec["date"],"%Y-%m-%d").date(), rec["location"],
                rec.get("tot"), rec.get("b"), rec.get("f1"), rec.get("f2"),
                None, None, rec.get("deductions","")]
        for ci, val in enumerate(vals, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cs(cell, bg=bg, al="right" if isinstance(val,float) else ("left" if ci in(2,9) else "center"))
            if ci==1: cell.number_format="DD/MM/YYYY"
            elif ci in(3,4,5,6) and isinstance(val,float): cell.number_format=EUR
    for ci, w in enumerate([12,10,13,12,13,13,14,12,40],1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "A2"
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()

# ── Riepilogo ─────────────────────────────────────────────────────────────
def send_summary(db, for_date=None):
    c = cfg()
    if for_date is None:
        for_date = (date.today() - timedelta(days=1)).isoformat()
    recs = [r for r in db["closures"] if r["date"] == for_date]
    bol  = next((r for r in recs if r["location"]=="Bologna"), None)
    pad  = next((r for r in recs if r["location"]=="Padova"),  None)
    d_fmt = date.fromisoformat(for_date).strftime("%d/%m/%Y")
    lines = [f"📊 <b>Chiusure {d_fmt}</b>\n"]
    if bol:
        lines += ["🔵 <b>BOLOGNA</b>",
                  f"Tot: {fmt(bol.get('tot'))}  |  F1: {fmt(bol.get('f1'))}  |  B: {fmt(bol.get('b'))}"]
        if bol.get("f2"): lines.append(f"F2: {fmt(bol['f2'])}")
        if bol.get("deductions"): lines.append(f"📝 {bol['deductions']}")
    else:
        lines.append("🔵 <b>BOLOGNA</b> — nessun dato")
    lines.append("")
    if pad:
        lines += ["🟢 <b>PADOVA</b>",
                  f"Tot: {fmt(pad.get('tot'))}  |  F1: {fmt(pad.get('f1'))}  |  B: {fmt(pad.get('b'))}"]
        if pad.get("f2"): lines.append(f"F2: {fmt(pad['f2'])}")
        if pad.get("deductions"): lines.append(f"📝 {pad['deductions']}")
    else:
        lines.append("🟢 <b>PADOVA</b> — nessun dato")
    tot = ((bol or {}).get("tot") or 0) + ((pad or {}).get("tot") or 0)
    if tot: lines.append(f"\n💰 <b>TOTALE: {fmt(tot)}</b>")
    send(c["ricardo_id"], "\n".join(lines))
    if db["closures"]:
        send_file(c["ricardo_id"], "chiusure_cassa.xlsx",
                  make_excel(db["closures"]), "📎 Excel aggiornato")
    db["daily_sent"] = for_date
    save_db(db)

# ── Routes ────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return jsonify(ok=True, service="Chiusure Cassa Bot")

@app.route("/health")
def health():
    db = load_db()
    return jsonify(ok=True, closures=len(db.get("closures",[])))

@app.route("/daily", methods=["GET","POST"])
def daily():
    db = load_db()
    yesterday = (date.today()-timedelta(days=1)).isoformat()
    if db.get("daily_sent") == yesterday:
        return jsonify(ok=True, msg="già inviato")
    send_summary(db)
    return jsonify(ok=True, msg=f"riepilogo inviato per {yesterday}")

@app.route("/webhook", methods=["POST"])
def webhook():
    c = cfg()
    if not c["token"]:
        return jsonify(ok=False, error="BOT_TOKEN non impostato"), 500

    db  = load_db()
    upd = request.get_json(silent=True) or {}
    msg = upd.get("message")
    if not msg:
        return jsonify(ok=True)

    chat_id = msg.get("chat",{}).get("id")
    text    = msg.get("text","")
    msg_id  = msg.get("message_id")
    ts      = msg.get("date",0)
    sender  = msg.get("from",{}).get("first_name","?")

    # Comandi da Ricardo in privato
    if chat_id == c["ricardo_id"]:
        cmd = text.strip().split()[0] if text.strip() else ""
        if cmd == "/excel":
            if db["closures"]:
                send_file(c["ricardo_id"],"chiusure_cassa.xlsx",
                          make_excel(db["closures"]),"📎 Tutte le chiusure")
            else:
                send(c["ricardo_id"],"Nessuna chiusura registrata.")
        elif cmd in ("/oggi","/ieri","/summary"):
            td = date.today().isoformat() if cmd=="/oggi" else None
            send_summary(db, td)
        elif cmd == "/stato":
            n = len(db["closures"])
            last = db["closures"][-1]["date"] if db["closures"] else "—"
            send(c["ricardo_id"],f"✅ Bot attivo\n📦 {n} chiusure\n📅 Ultima: {last}")
        return jsonify(ok=True)

    # Messaggi dal gruppo
    if chat_id != c["group_id"] or not text:
        return jsonify(ok=True)

    if is_correction(text):
        send(c["ricardo_id"],
             f"⚠️ <b>Correzione ricevuta</b>\n\nDa: <b>{sender}</b>\n"
             f"Testo:\n<pre>{text}</pre>\n\nVerifica e aggiorna l'Excel.")
        return jsonify(ok=True)

    location = detect_location(text)
    if not location: return jsonify(ok=True)
    cash = extract_cash(text)
    if not cash: return jsonify(ok=True)

    eff_d = eff_date(ts).isoformat()
    existing = next((r for r in db["closures"]
                     if r["date"]==eff_d and r["location"]==location), None)
    if existing:
        existing.update({**cash,"sender":sender})
    else:
        db["closures"].append({"date":eff_d,"location":location,"sender":sender,**cash})
    save_db(db)
    delete_msg(c["group_id"], msg_id)

    d_fmt = date.fromisoformat(eff_d).strftime("%d/%m")
    notif = (f"✅ <b>{location} — {d_fmt}</b>\n"
             f"Tot: {fmt(cash.get('tot'))}  |  F1: {fmt(cash.get('f1'))}  |  B: {fmt(cash.get('b'))}")
    if cash.get("f2"): notif += f"  |  F2: {fmt(cash['f2'])}"
    if cash.get("deductions"): notif += f"\n📝 {cash['deductions']}"
    send(c["ricardo_id"], notif)
    log.info(f"Salvata chiusura: {location} {eff_d}")
    return jsonify(ok=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)

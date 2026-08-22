# -*- coding: utf-8 -*-
"""
कार्य-प्रगति वृत्त — Flask बैकएंड
==================================
- /form            → सार्वजनिक हिंदी फॉर्म (शेयर लिंक, कोई ईमेल/लॉगिन नहीं)
- /submit          → उत्तर सहेजना (POST)
- /dhanyavaad      → धन्यवाद पृष्ठ
- /admin           → एडमिन डैशबोर्ड (टेबल + इन्फोग्राफिक्स + Excel)
- /admin/export    → Excel (.xlsx) डाउनलोड — हमेशा ताज़ा
- /admin/clear     → सारा डेटा साफ़ करना
- /admin/api/data  → डैशबोर्ड द्वारा 10 सेकंड में ऑटो-रिफ्रेश वाला JSON
"""
import io
import os
import json
import time
import hashlib
import sqlite3
import hmac
import secrets
from datetime import datetime, date, timedelta

from flask import (
    Flask, request, render_template, redirect, url_for,
    session, jsonify, send_file, flash, abort
)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from survey_config import (
    SECTIONS, FORM_TITLE, FORM_DESC, INSTRUCTION, THANK_YOU_MSG,
    ADMIN_PASSWORD, SITE_TITLE, ORG_NAME, ORG_TAGLINE, ORG_SUB, EXPECTED_PROVINCES
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "data", "responses.db")
LOGO_EXISTS = os.path.exists(os.path.join(BASE_DIR, "static", "logo.png"))

app = Flask(__name__)


def _load_secret_key():
    """SECRET_KEY: पर्यावरण-चर से, वरना स्थायी फ़ाइल से (रीस्टार्ट पर भी वही रहे)"""
    env = os.environ.get("SECRET_KEY")
    if env:
        return env
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    path = os.path.join(os.path.dirname(DB_PATH), "secret_key.txt")
    try:
        with open(path) as fh:
            k = fh.read().strip()
            if k:
                return k
    except FileNotFoundError:
        pass
    k = secrets.token_hex(32)
    with open(path, "w") as fh:
        fh.write(k)
    return k


app.secret_key = _load_secret_key()


# ---------------------------------------------------------------- सपाट फ़ील्ड सूची
def flat_fields():
    """सभी फ़ील्ड एक सपाट सूची में — क्रम: प्रारंभिक → क → ख → ग → घ → ड"""
    out = []
    for sec in SECTIONS:
        for item in sec.get("items", []):
            for f in item.get("fields", []):
                out.append({
                    "id": f["id"],
                    "label": f.get("label", ""),
                    "short": f.get("short", ""),
                    "type": f.get("type", "text"),
                    "grp": item.get("grp", ""),
                    "item_label": item.get("label", ""),
                    "required": bool(f.get("required", True)),
                    "full": bool(f.get("full", False)),
                    "options": f.get("options") or [],
                    "show_if": f.get("show_if"),
                    "show_if_value": f.get("show_if_value"),
                })
    return out


FLAT = flat_fields()

# Excel/तालिका के हेडर — "क.1 — जिलों की वर्तमान संख्या" जैसा
def col_header(f):
    return f"{f['grp']} — {f['label']}" if f["grp"] else f["label"]


# ---------------------------------------------------------------- डेटाबेस
# स्थानीय (SQLite) चलता रहता है; होस्टिंग पर DATABASE_URL (PostgreSQL) दें →
# डेटा स्थायी क्लाउड डेटाबेस में रहेगा (नीचे README/deploy निर्देश)
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()


def get_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _pg():
    import psycopg2
    return psycopg2.connect(DATABASE_URL, sslmode="require")


def db_init():
    if DATABASE_URL:
        conn = _pg()
        try:
            with conn:
                with conn.cursor() as c:
                    c.execute("""
                        CREATE TABLE IF NOT EXISTS responses (
                            id SERIAL PRIMARY KEY,
                            created_at TEXT NOT NULL,
                            data TEXT NOT NULL
                        )
                    """)
        finally:
            conn.close()
    else:
        with get_db() as db:
            db.execute("""
                CREATE TABLE IF NOT EXISTS responses (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    created_at TEXT NOT NULL,
                    data TEXT NOT NULL
                )
            """)


def db_insert(created_at, data_json):
    if DATABASE_URL:
        conn = _pg()
        try:
            with conn:
                with conn.cursor() as c:
                    c.execute("INSERT INTO responses (created_at, data) VALUES (%s, %s)",
                              (created_at, data_json))
        finally:
            conn.close()
    else:
        with get_db() as db:
            db.execute("INSERT INTO responses (created_at, data) VALUES (?, ?)",
                       (created_at, data_json))


def db_all():
    """सभी उत्तर — नए पहले (id DESC)"""
    if DATABASE_URL:
        conn = _pg()
        try:
            with conn.cursor() as c:
                c.execute("SELECT id, created_at, data FROM responses ORDER BY id DESC")
                cols = [d[0] for d in c.description]
                return [dict(zip(cols, row)) for row in c.fetchall()]
        finally:
            conn.close()
    else:
        rows = []
        with get_db() as db:
            for r in db.execute("SELECT id, created_at, data FROM responses ORDER BY id DESC"):
                rows.append({"id": r["id"], "created_at": r["created_at"], "data": r["data"]})
        return rows


def db_clear():
    if DATABASE_URL:
        conn = _pg()
        try:
            with conn:
                with conn.cursor() as c:
                    c.execute("DELETE FROM responses")
        finally:
            conn.close()
    else:
        with get_db() as db:
            db.execute("DELETE FROM responses")


db_init()


def parse_data(row):
    return json.loads(row["data"])


# ---------------------------------------------------------------- सार्वजनिक फॉर्म
@app.route("/")
def index():
    return redirect(url_for("form_page"))


@app.route("/form")
def form_page():
    return render_template(
        "form.html",
        title=FORM_TITLE, desc=FORM_DESC, instruction=INSTRUCTION,
        sections=SECTIONS, site=SITE_TITLE, org_name=ORG_NAME,
        org_tagline=ORG_TAGLINE, org_sub=ORG_SUB, logo=LOGO_EXISTS,
    )


@app.route("/submit", methods=["POST"])
def submit():
    # हनीपॉट — स्पैम बॉट यहाँ लिखते हैं
    if request.form.get("website"):
        return redirect(url_for("thank_you"))

    payload = {}
    for f in FLAT:
        payload[f["id"]] = (request.form.get(f["id"]) or "").strip()

    ok = True
    for f in FLAT:
        if not f["required"] or payload[f["id"]]:
            continue
        # सशर्त फ़ील्ड: शर्त पूरी होने पर ही ज़रूरी (जैसे "नहीं" चुनने पर ही तिथि)
        if f.get("show_if") and payload.get(f["show_if"]) != f.get("show_if_value"):
            continue
        ok = False

    if not ok:
        flash("कृपया सभी ज़रूरी (*) प्रश्न भरें।", "error")
        return redirect(url_for("form_page"))

    db_insert(datetime.now().isoformat(timespec="seconds"),
              json.dumps(payload, ensure_ascii=False))
    return redirect(url_for("thank_you"))


@app.route("/dhanyavaad")
def thank_you():
    return render_template("thank_you.html", msg=THANK_YOU_MSG, site=SITE_TITLE,
                           org_name=ORG_NAME, org_sub=ORG_SUB, logo=LOGO_EXISTS)


# ---------------------------------------------------------------- एडमिन प्रमाणीकरण
# कुकी-मुक्त प्रमाणीकरण: सही पासवर्ड पर हस्ताक्षरित कुंजी (URL में) मिलती है।
# यह प्रीव्यू iframe / मोबाइल ब्राउज़रों में भी चलता है जहाँ कुकीज़ ब्लॉक होती हैं।

def make_admin_key(hours=24):
    ts = int(time.time()) + hours * 3600
    body = f"admin:{ts}".encode()
    sig = hmac.new(app.secret_key.encode(), body, hashlib.sha256).hexdigest()[:32]
    return body.decode() + "." + sig


def check_admin_key(key):
    if not key or "." not in key:
        return False
    body, sig = key.rsplit(".", 1)
    expected = hmac.new(app.secret_key.encode(), body.encode(), hashlib.sha256).hexdigest()[:32]
    if not hmac.compare_digest(sig, expected):
        return False
    try:
        parts = body.split(":")
        if parts[0] != "admin":
            return False
        if int(parts[1]) < time.time():
            return False  # समय-सीमा समाप्त
    except Exception:
        return False
    return True


def is_admin():
    """कुकी सत्र या URL कुंजी — दोनों में से कोई भी मान्य"""
    if session.get("admin"):
        return True
    k = request.args.get("key") or request.headers.get("X-Admin-Key")
    return bool(k) and check_admin_key(k)


@app.route("/admin/login", methods=["POST"])
def admin_login():
    pw = request.form.get("password", "")
    expected = os.environ.get("ADMIN_PASSWORD", ADMIN_PASSWORD)
    if hmac.compare_digest(pw, expected):
        session["admin"] = True
        return redirect(url_for("admin", key=make_admin_key()))
    flash("गलत पासवर्ड!", "error")
    return redirect(url_for("admin"))


@app.route("/admin/logout", methods=["POST"])
def admin_logout():
    session.pop("admin", None)
    return redirect(url_for("admin"))


@app.route("/admin/clear", methods=["POST"])
def admin_clear():
    if not session.get("admin"):
        abort(403)
    db_clear()
    return redirect(url_for("admin"))


# ---------------------------------------------------------------- एडमिन डैशबोर्ड
@app.route("/admin")
def admin():
    key = request.args.get("key")
    if key and check_admin_key(key):
        session["admin"] = True  # कुकी चले तो आगे के लिए सत्र भी रखें
    elif not session.get("admin"):
        return render_template("admin_login.html", site=SITE_TITLE,
                               org_name=ORG_NAME, logo=LOGO_EXISTS)
    # Excel बटन में वही कुंजी जुड़े (कुकी ब्लॉक होने पर भी काम करे)
    export_url = url_for("admin_export", key=key) if key else url_for("admin_export")
    return render_template(
        "admin.html",
        title=FORM_TITLE, site=SITE_TITLE, export_url=export_url, org_name=ORG_NAME,
        org_tagline=ORG_TAGLINE, org_sub=ORG_SUB, logo=LOGO_EXISTS,
    )


def to_num(v):
    """स्ट्रिंग → संख्या (जितना हो सके)"""
    v = (v or "").strip()
    if not v:
        return None
    try:
        f = float(v)
        return int(f) if f.is_integer() else f
    except ValueError:
        return None


def build_summary():
    """सभी उत्तरों से आँकड़े बनाओ"""
    rows = db_all()

    total = len(rows)
    today_str = date.today().isoformat()
    week_start = (date.today() - timedelta(days=6)).isoformat()
    today_count = sum(1 for r in rows if r["created_at"][:10] == today_str)
    week_count = sum(1 for r in rows if r["created_at"][:10] >= week_start)
    last_time = rows[0]["created_at"] if rows else None

    # पिछले 14 दिन की समय-श्रृंखला
    series = {}
    for i in range(13, -1, -1):
        d = (date.today() - timedelta(days=i)).strftime("%d/%m")
        series[d] = 0
    for r in rows:
        key = r["created_at"][8:10] + "/" + r["created_at"][5:7]  # dd/mm
        if key in series:
            series[key] += 1

    # प्रति-फ़ील्ड योग (संख्यात्मक)
    field_sums = {}
    for f in FLAT:
        if f["type"] == "number":
            vals = [to_num(parse_data(r).get(f["id"])) for r in rows]
            vals = [v for v in vals if v is not None]
            field_sums[f["id"]] = sum(vals)
        else:
            field_sums[f["id"]] = None

    # ---------- उत्तर-सूची ट्रैकर: किस प्रांत ने भरा / नहीं भरा ----------
    def _norm(s):
        return (s or "").replace(" ", "").replace("\u3000", "").lower()

    filled = {}  # सामान्यीकृत नाम → जानकारी (rows नए पहले हैं, इसलिए पहला = नवीनतम)
    for r in rows:
        d = parse_data(r)
        prov = (d.get("prov") or "").strip()
        if not prov:
            continue
        k = _norm(prov)
        if k in filled:
            continue
        filled[k] = {"name": prov, "at": r["created_at"], "head": d.get("head", "")}

    tracker = {"expected": EXPECTED_PROVINCES, "status": [], "extra": []}
    for name in EXPECTED_PROVINCES:
        k = _norm(name)
        st = filled.pop(k, None)
        tracker["status"].append({
            "name": name, "filled": bool(st),
            "at": st["at"] if st else None, "head": st["head"] if st else None,
        })
    tracker["extra"] = [filled[k] for k in filled]  # सूची से बाहर के प्रांत (जमा हुए)
    tracker["filled"] = sum(1 for s in tracker["status"] if s["filled"])
    tracker["total"] = len(tracker["status"])
    tracker["pending"] = [s["name"] for s in tracker["status"] if not s["filled"]]

    return {
        "total": total,
        "today": today_count,
        "week": week_count,
        "last_time": last_time,
        "series": [{"label": k, "value": v} for k, v in series.items()],
        "fields": [
            {"id": f["id"], "label": f["label"], "short": f.get("short", ""),
             "type": f["type"],
             "grp": f["grp"], "item_label": f["item_label"],
             "header": col_header(f),
             "show_if": f.get("show_if"), "show_if_value": f.get("show_if_value")}
            for f in FLAT
        ],
        "field_sums": field_sums,
        "tracker": tracker,
        "rows": [
            {"id": r["id"], "created_at": r["created_at"], "data": parse_data(r)}
            for r in rows
        ],
        "sections": [
            {"title": sec["title"],
             "items": [{"grp": it.get("grp", ""), "label": it.get("label", "")}
                       for it in sec.get("items", [])]}
            for sec in SECTIONS
        ],
    }


@app.route("/admin/api/data")
def admin_api_data():
    if not is_admin():
        return jsonify({"error": "unauthorized"}), 401
    return jsonify(build_summary())


# ---------------------------------------------------------------- Excel एक्सपोर्ट
def build_workbook(s):
    """सारा Excel (2 शीट) बनाओ — डाउनलोड और workspace-सेव दोनों इसी से"""
    wb = Workbook()
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor="1D4ED8")
    thin = Side(style="thin", color="D1D5DB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ---- शीट 1: सारे उत्तर (हेडर बिल्कुल हिंदी में, प्रश्नों के अनुसार)
    ws = wb.active
    ws.title = "उत्तर"
    headers = ["क्रमांक", "जमा करने का समय (IST)"] + [col_header(f) for f in FLAT]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    for i, r in enumerate(s["rows"], start=1):
        ts = r["created_at"].replace("T", "  ")
        row_vals = [i, ts]
        for f in FLAT:
            raw = r["data"].get(f["id"], "")
            if f.get("show_if") and r["data"].get(f["show_if"]) != f.get("show_if_value"):
                row_vals.append("लागू नहीं")  # शर्त पूरी न होने पर (जैसे "हाँ" चुनने पर)
            elif f["type"] == "number":
                n = to_num(raw)
                row_vals.append(n if n is not None else "")
            else:
                row_vals.append(raw)
        ws.append(row_vals)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 21
    col = 3
    for f in FLAT:
        ws.column_dimensions[get_column_letter(col)].width = min(34, 8 + len(f["label"]))
        col += 1
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(ws.max_row, 2)}"

    # ---- शीट 2: सारांश (कुल योग + प्रांत-वार)
    ws2 = wb.create_sheet("सारांश")
    ws2.append([FORM_TITLE])
    ws2["A1"].font = Font(bold=True, size=14, color="1D4ED8")
    ws2.append(["कुल उत्तर (प्रांत)", s["total"]])
    ws2.append(["आज के उत्तर", s["today"]])
    ws2.append(["पिछले 7 दिन", s["week"]])
    ws2.append(["अंतिम उत्तर", s["last_time"] or "—"])
    ws2.append([])

    # खंड-वार कुल योग
    for sec in s["sections"]:
        items = [it for it in sec["items"] if it["grp"]]
        if not items:
            continue
        ws2.append([sec["title"]])
        ws2.cell(row=ws2.max_row, column=1).font = Font(bold=True, color="1D4ED8")
        for it in items:
            ws2.append([f"{it['grp']} {it['label']}"])
            ws2.cell(row=ws2.max_row, column=1).font = Font(bold=True)
            for f in FLAT:
                if f["grp"] == it["grp"]:
                    v = s["field_sums"].get(f["id"])
                    txt = str(v) if v is not None else "—"
                    ws2.append(["", f"{f['label']} — {txt}"])
        ws2.append([])

    # प्रांत-वार एक नज़र
    ws2.append(["प्रांत-वार (मुख्य आँकड़े)"])
    ws2.cell(row=ws2.max_row, column=1).font = Font(bold=True, color="1D4ED8")
    key_fields = ["prov", "head", "k1_cur", "k1_tgt", "k3", "gh1", "d1_done", "d2_pres"]
    ws2.append(["प्रांत", "अध्यक्ष", "वर्तमान जिले", "सितंबर लक्ष्य (जिले)",
                "नगर", "सितंबर यज्ञ स्थल", "मिलन सम्पन्न (स्थान)", "मिलन उपस्थिति"])
    for c in range(1, 9):
        ws2.cell(row=ws2.max_row, column=c).font = Font(bold=True)
    for r in s["rows"]:
        ws2.append([r["data"].get("prov", ""), r["data"].get("head", "")] +
                   [r["data"].get(k, "") for k in key_fields[2:]])

    ws2.column_dimensions["A"].width = 34
    ws2.column_dimensions["B"].width = 24
    for c in range(3, 9):
        ws2.column_dimensions[get_column_letter(c)].width = 20

    return wb


@app.route("/admin/export")
def admin_export():
    """सीधा डाउनलोड — असली होस्टिंग (Render/Railway) पर यही बटन चलता है"""
    if not is_admin():
        abort(403)
    wb = build_workbook(build_summary())
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = f"kary-pragati-vritt_{date.today().isoformat()}.xlsx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=fname,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------- रन
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port, debug=False)

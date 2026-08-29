# -*- coding: utf-8 -*-
"""
कार्य-प्रगति वृत्त — बहु-फॉर्म Flask बैकएंड
==========================================
- /                    → मुख्य पेज (सभी फॉर्मों के कार्ड)
- /form                → पुराना लिंक → /f/progress पर भेजता है (पीछे-संगतता)
- /f/<survey>          → सार्वजनिक हिंदी फॉर्म (शेयर लिंक, कोई ईमेल/लॉगिन नहीं)
- /submit/<survey>     → उत्तर सहेजना (POST; माह भी)
- /dhanyavaad/<survey> → धन्यवाद पृष्ठ
- /admin?survey=&month=→ एडमिन डैशबोर्ड (टैब + माह-फ़िल्टर + ट्रैकर + Excel)
- /admin/export        → Excel (.xlsx) — चुने फॉर्म/माह का, हमेशा ताज़ा
- /admin/clear         → केवल चुने फॉर्म का डेटा साफ़
- /admin/api/data      → 10 सेकंड ऑटो-रिफ्रेश JSON (?survey=&month=)
डेटाबेस में माइग्रेशन अपने-आप: survey / report_month कॉलम जुड़ जाते हैं,
पुराने उत्तर survey='progress' मानकर सुरक्षित रहते हैं (माह: "सभी माह" में दिखते हैं)।
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
    SURVEYS, MONTHS, CURRENT_MONTH, ADMIN_PASSWORD, SITE_TITLE,
    ORG_NAME, ORG_TAGLINE, ORG_SUB
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "data", "responses.db")
LOGO_EXISTS = os.path.exists(os.path.join(BASE_DIR, "static", "logo.png"))
DEFAULT_SURVEY = "progress"

app = Flask(__name__)


def _load_secret_key():
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


# ---------------------------------------------------------------- सर्वे सहायक
def get_survey(key):
    cfg = SURVEYS.get(key)
    if not cfg:
        abort(404)
    return cfg


def flat_fields(cfg):
    """सर्वे के सभी फ़ील्ड सपाट सूची में (प्रांत-ड्रॉपडाउन पहले)"""
    out = []
    if cfg.get("prov_dropdown"):
        # "prov" खंडों में नहीं, फॉर्म के ऊपर ड्रॉपडाउन से चुना जाता है —
        # पर सहेजना/Excel/ट्रैकर पहले जैसा ही (पहला कॉलम)
        out.append({
            "id": "prov", "label": "प्रांत का नाम", "short": "प्रांत",
            "type": "select", "grp": "", "item_label": "प्रारंभिक जानकारी",
            "required": True, "full": True,
            "options": cfg.get("expected_provinces") or [],
            "show_if": None, "show_if_value": None,
            "auto_sum": None, "readonly": False,
        })
    for sec in cfg.get("sections", []):
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
                    "auto_sum": f.get("auto_sum"),
                    "readonly": bool(f.get("readonly")),
                })
    return out


FLATS = {k: flat_fields(v) for k, v in SURVEYS.items()}


def col_header(f):
    return f"{f['grp']} — {f['label']}" if f["grp"] else f["label"]


def org_ctx():
    return {"org_name": ORG_NAME, "org_tagline": ORG_TAGLINE,
            "org_sub": ORG_SUB, "logo": LOGO_EXISTS, "site": SITE_TITLE}


# ---------------------------------------------------------------- डेटाबेस
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
    """तालिका बनाओ + पुरानी तालिका में survey/report_month कॉलम जोड़ो (माइग्रेशन)"""
    create_sqlite = """
        CREATE TABLE IF NOT EXISTS responses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT NOT NULL,
            survey TEXT NOT NULL DEFAULT 'progress',
            report_month TEXT,
            data TEXT NOT NULL
        )"""
    create_pg = create_sqlite.replace(
        "INTEGER PRIMARY KEY AUTOINCREMENT", "SERIAL PRIMARY KEY")

    if DATABASE_URL:
        conn = _pg()
        try:
            with conn:
                with conn.cursor() as c:
                    c.execute(create_pg)
                    # माइग्रेशन — पुरानी तालिका में कॉलम न हों तो जोड़ें
                    c.execute("ALTER TABLE responses ADD COLUMN IF NOT EXISTS survey TEXT NOT NULL DEFAULT 'progress'")
                    c.execute("ALTER TABLE responses ADD COLUMN IF NOT EXISTS report_month TEXT")
        finally:
            conn.close()
    else:
        with get_db() as db:
            db.execute(create_sqlite)
            cols = {r["name"] for r in db.execute("PRAGMA table_info(responses)")}
            if "survey" not in cols:
                db.execute("ALTER TABLE responses ADD COLUMN survey TEXT NOT NULL DEFAULT 'progress'")
            if "report_month" not in cols:
                db.execute("ALTER TABLE responses ADD COLUMN report_month TEXT")


def db_insert(created_at, survey, report_month, data_json):
    if DATABASE_URL:
        conn = _pg()
        try:
            with conn:
                with conn.cursor() as c:
                    c.execute(
                        "INSERT INTO responses (created_at, survey, report_month, data) VALUES (%s, %s, %s, %s)",
                        (created_at, survey, report_month, data_json))
        finally:
            conn.close()
    else:
        with get_db() as db:
            db.execute(
                "INSERT INTO responses (created_at, survey, report_month, data) VALUES (?, ?, ?, ?)",
                (created_at, survey, report_month, data_json))


def db_rows(survey=None, month=None):
    """फ़िल्टर किए उत्तर — नए पहले। month=None/'all' → सभी माह"""
    sql = "SELECT id, created_at, survey, report_month, data FROM responses"
    cond, args = [], []
    if survey:
        cond.append("survey = %s" if DATABASE_URL else "survey = ?")
        args.append(survey)
    if month and month != "all":
        cond.append("report_month = %s" if DATABASE_URL else "report_month = ?")
        args.append(month)
    if cond:
        sql += " WHERE " + " AND ".join(cond)
    sql += " ORDER BY id DESC"

    if DATABASE_URL:
        conn = _pg()
        try:
            with conn.cursor() as c:
                c.execute(sql, args)
                cols = [d[0] for d in c.description]
                return [dict(zip(cols, row)) for row in c.fetchall()]
        finally:
            conn.close()
    else:
        with get_db() as db:
            return [dict(r) for r in db.execute(sql, args).fetchall()]


def db_clear(survey):
    """केवल उसी फॉर्म का डेटा साफ़"""
    if DATABASE_URL:
        conn = _pg()
        try:
            with conn:
                with conn.cursor() as c:
                    c.execute("DELETE FROM responses WHERE survey = %s", (survey,))
        finally:
            conn.close()
    else:
        with get_db() as db:
            db.execute("DELETE FROM responses WHERE survey = ?", (survey,))


db_init()


def parse_data(row):
    return json.loads(row["data"])


# ---------------------------------------------------------------- सार्वजनिक पेज
@app.route("/")
def index():
    return render_template("landing.html", surveys=SURVEYS, **org_ctx())


# पुराने लिंक (पीछे-संगतता): /form → पहला फॉर्म
@app.route("/form")
def form_compat():
    return redirect(url_for("form_page", survey=DEFAULT_SURVEY))


@app.route("/f/<survey>")
def form_page(survey):
    cfg = get_survey(survey)
    others = [(k, v["title"]) for k, v in SURVEYS.items() if k != survey]
    return render_template(
        "form.html",
        survey_key=survey, cfg=cfg, title=cfg["title"], desc=cfg["desc"],
        instruction=cfg["instruction"], sections=cfg["sections"],
        month_field=cfg.get("month_field", False), months=MONTHS,
        current_month=CURRENT_MONTH,
        prov_dropdown=cfg.get("prov_dropdown", False),
        provinces=cfg.get("expected_provinces") or [],
        others=others, **org_ctx(),
    )


@app.route("/submit/<survey>", methods=["POST"])
def submit(survey):
    cfg = get_survey(survey)
    # हनीपॉट — स्पैम बॉट यहाँ लिखते हैं
    if request.form.get("website"):
        return redirect(url_for("thank_you", survey=survey))

    flat = FLATS[survey]
    payload = {f["id"]: (request.form.get(f["id"]) or "").strip() for f in flat}

    # ऑटो-सम ("योग" फ़ील्ड): खाली हो तो सर्वर ही जोड़ दे (JS न चले तो भी सही)
    for f in flat:
        if f.get("auto_sum") and not payload[f["id"]]:
            tot, have = 0.0, False
            for src in f["auto_sum"]:
                v = to_num(payload.get(src))
                if v is not None:
                    tot += v
                    have = True
            if have:
                payload[f["id"]] = str(int(tot)) if float(tot).is_integer() else str(tot)

    # रिपोर्टिंग माह (यदि सक्षम)
    report_month = None
    if cfg.get("month_field"):
        report_month = (request.form.get("__month") or "").strip()
        if report_month not in MONTHS:
            flash("कृपया रिपोर्टिंग माह चुनें।", "error")
            return redirect(url_for("form_page", survey=survey))

    # प्रांत ड्रॉपडाउन (यदि सक्षम) — सूची के बाहर कुछ न आए
    if cfg.get("prov_dropdown"):
        plist = cfg.get("expected_provinces") or []
        if payload.get("prov", "") not in plist:
            flash("कृपया सूची से अपना प्रांत चुनें।", "error")
            return redirect(url_for("form_page", survey=survey))

    ok = True
    for f in flat:
        if not f["required"] or payload[f["id"]]:
            continue
        # सशर्त फ़ील्ड: शर्त पूरी होने पर ही ज़रूरी
        if f.get("show_if") and payload.get(f["show_if"]) != f.get("show_if_value"):
            continue
        ok = False

    if not ok:
        flash("कृपया सभी ज़रूरी (*) प्रश्न भरें।", "error")
        return redirect(url_for("form_page", survey=survey))

    db_insert(datetime.now().isoformat(timespec="seconds"),
              survey, report_month, json.dumps(payload, ensure_ascii=False))
    return redirect(url_for("thank_you", survey=survey))


# पुराना /submit (पुराने कैश किए पेज भेजें तो भी सुरक्षित)
@app.route("/submit", methods=["POST"])
def submit_compat():
    return submit(DEFAULT_SURVEY)


@app.route("/dhanyavaad/<survey>")
def thank_you(survey):
    cfg = get_survey(survey)
    return render_template("thank_you.html", msg=cfg["thank_you"],
                           survey_key=survey, **org_ctx())


@app.route("/dhanyavaad")
def thank_you_compat():
    return redirect(url_for("thank_you", survey=DEFAULT_SURVEY))


# ---------------------------------------------------------------- एडमिन प्रमाणीकरण
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
            return False
    except Exception:
        return False
    return True


def is_admin():
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
    if not is_admin():
        abort(403)
    survey = request.form.get("survey", DEFAULT_SURVEY)
    get_survey(survey)  # अमान्य हो तो 404
    db_clear(survey)
    return redirect(url_for("admin", survey=survey,
                            key=request.form.get("key") or None))


# ---------------------------------------------------------------- एडमिन डैशबोर्ड
def _sel_survey_month():
    survey = request.args.get("survey", DEFAULT_SURVEY)
    get_survey(survey)
    cfg = SURVEYS[survey]
    if cfg.get("month_field"):
        month = request.args.get("month") or CURRENT_MONTH
        if month != "all" and month not in MONTHS:
            month = CURRENT_MONTH
    else:
        month = "all"
    return survey, cfg, month


@app.route("/admin")
def admin():
    key = request.args.get("key")
    if key and check_admin_key(key):
        session["admin"] = True
    elif not session.get("admin"):
        return render_template("admin_login.html", **org_ctx())

    survey, cfg, month = _sel_survey_month()

    def q(extra):
        return {**extra, "key": key} if key else extra

    return render_template(
        "admin.html",
        survey_key=survey, cfg=cfg, month=month, months=MONTHS,
        surveys=SURVEYS, url_key=key or "",
        export_url=url_for("admin_export", **q({"survey": survey, "month": month})),
        data_url=url_for("admin_api_data", **q({"survey": survey, "month": month})),
        title=cfg["title"], **org_ctx(),
    )


def to_num(v):
    v = (v or "").strip()
    if not v:
        return None
    try:
        f = float(v)
        return int(f) if f.is_integer() else f
    except ValueError:
        return None


def build_summary(survey, cfg, month):
    rows = db_rows(survey=survey, month=month)
    flat = FLATS[survey]

    total = len(rows)
    today_str = date.today().isoformat()
    week_start = (date.today() - timedelta(days=6)).isoformat()
    today_count = sum(1 for r in rows if r["created_at"][:10] == today_str)
    week_count = sum(1 for r in rows if r["created_at"][:10] >= week_start)
    last_time = rows[0]["created_at"] if rows else None

    series = {}
    for i in range(13, -1, -1):
        series[(date.today() - timedelta(days=i)).strftime("%d/%m")] = 0
    for r in rows:
        key = r["created_at"][8:10] + "/" + r["created_at"][5:7]
        if key in series:
            series[key] += 1

    field_sums = {}
    for f in flat:
        if f["type"] == "number":
            vals = [to_num(parse_data(r).get(f["id"])) for r in rows]
            field_sums[f["id"]] = sum(v for v in vals if v is not None)
        else:
            field_sums[f["id"]] = None

    # ---------- उत्तर-सूची ट्रैकर ----------
    expected = cfg.get("expected_provinces") or []

    def _norm(s):
        return (s or "").replace(" ", "").replace("　", "").lower()

    filled = {}
    for r in rows:
        d = parse_data(r)
        prov = (d.get("prov") or "").strip()
        if not prov:
            continue
        k = _norm(prov)
        if k in filled:
            continue
        filled[k] = {"name": prov, "at": r["created_at"], "head": d.get("head", "")}

    tracker = {"expected": expected, "status": [], "extra": []}
    for name in expected:
        st = filled.pop(_norm(name), None)
        tracker["status"].append({
            "name": name, "filled": bool(st),
            "at": st["at"] if st else None, "head": st["head"] if st else None})
    tracker["extra"] = [filled[k] for k in filled]
    tracker["filled"] = sum(1 for s in tracker["status"] if s["filled"])
    tracker["total"] = len(tracker["status"])
    tracker["pending"] = [s["name"] for s in tracker["status"] if not s["filled"]]

    return {
        "total": total, "today": today_count, "week": week_count,
        "last_time": last_time,
        "show_month_col": bool(cfg.get("month_field")) and month == "all",
        "series": [{"label": k, "value": v} for k, v in series.items()],
        "fields": [
            {"id": f["id"], "label": f["label"], "short": f.get("short", ""),
             "type": f["type"], "grp": f["grp"], "item_label": f["item_label"],
             "header": col_header(f),
             "show_if": f.get("show_if"), "show_if_value": f.get("show_if_value")}
            for f in flat
        ],
        "field_sums": field_sums,
        "tracker": tracker,
        "rows": [
            {"id": r["id"], "created_at": r["created_at"],
             "report_month": r.get("report_month"), "data": parse_data(r)}
            for r in rows
        ],
        "sections": [
            {"title": sec["title"],
             "items": [{"grp": it.get("grp", ""), "label": it.get("label", "")}
                       for it in sec.get("items", [])]}
            for sec in cfg.get("sections", [])
        ],
    }


@app.route("/admin/api/data")
def admin_api_data():
    if not is_admin():
        return jsonify({"error": "unauthorized"}), 401
    survey, cfg, month = _sel_survey_month()
    return jsonify(build_summary(survey, cfg, month))


# ---------------------------------------------------------------- Excel एक्सपोर्ट
def build_workbook(s, survey, cfg, month):
    wb = Workbook()
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor="1D4ED8")
    thin = Side(style="thin", color="D1D5DB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    flat = FLATS[survey]
    month_on = bool(cfg.get("month_field"))

    # ---- शीट 1: सारे उत्तर
    ws = wb.active
    ws.title = "उत्तर"
    headers = ["क्रमांक", "जमा करने का समय (IST)"]
    if month_on:
        headers.append("रिपोर्टिंग माह")
    headers += [col_header(f) for f in flat]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    for i, r in enumerate(s["rows"], start=1):
        row_vals = [i, r["created_at"].replace("T", "  ")]
        if month_on:
            row_vals.append(r.get("report_month") or "—")
        for f in flat:
            raw = r["data"].get(f["id"], "")
            if f.get("show_if") and r["data"].get(f["show_if"]) != f.get("show_if_value"):
                row_vals.append("लागू नहीं")
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

    col = 1
    ws.column_dimensions[get_column_letter(col)].width = 8; col += 1
    ws.column_dimensions[get_column_letter(col)].width = 21; col += 1
    if month_on:
        ws.column_dimensions[get_column_letter(col)].width = 14; col += 1
    for f in flat:
        ws.column_dimensions[get_column_letter(col)].width = min(34, 8 + len(f["label"]))
        col += 1
    ws.freeze_panes = "D2" if month_on else "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(ws.max_row, 2)}"

    # ---- शीट 2: सारांश
    ws2 = wb.create_sheet("सारांश")
    title_line = cfg["title"] + (f" — {month}" if month_on and month != "all" else
                                 " — सभी माह" if month_on else "")
    ws2.append([title_line])
    ws2["A1"].font = Font(bold=True, size=14, color="1D4ED8")
    ws2.append(["कुल उत्तर (प्रांत)", s["total"]])
    ws2.append(["आज के उत्तर", s["today"]])
    ws2.append(["पिछले 7 दिन", s["week"]])
    ws2.append(["अंतिम उत्तर", s["last_time"] or "—"])
    t = s["tracker"]
    if t["total"]:
        ws2.append(["प्रांत भर चुके", f"{t['filled']} / {t['total']}"])
        ws2.append(["बाकी प्रांत", ", ".join(t["pending"]) or "—"])
    ws2.append([])

    for sec in s["sections"]:
        items = [it for it in sec["items"] if it["grp"]]
        if not items:
            continue
        ws2.append([sec["title"]])
        ws2.cell(row=ws2.max_row, column=1).font = Font(bold=True, color="1D4ED8")
        for it in items:
            ws2.append([f"{it['grp']} {it['label']}"])
            ws2.cell(row=ws2.max_row, column=1).font = Font(bold=True)
            for f in flat:
                if f["grp"] == it["grp"]:
                    v = s["field_sums"].get(f["id"])
                    txt = str(v) if v is not None else "—"
                    ws2.append(["", f"{f['label']} — {txt}"])
        ws2.append([])

    # प्रांत-वार एक नज़र (यदि कॉन्फ़िग किया हो)
    quick = cfg.get("excel_quick") or []
    if quick:
        pre = cfg.get("excel_quick_pre") or [
            {"id": "prov", "label": "प्रांत"}, {"id": "head", "label": "अध्यक्ष"}]
        ws2.append(["प्रांत-वार (मुख्य आँकड़े)"])
        ws2.cell(row=ws2.max_row, column=1).font = Font(bold=True, color="1D4ED8")
        qhead = [q["label"] for q in pre] + [q["label"] for q in quick]
        ws2.append(qhead)
        for c in range(1, len(qhead) + 1):
            ws2.cell(row=ws2.max_row, column=c).font = Font(bold=True)
        for r in s["rows"]:
            ws2.append([r["data"].get(q["id"], "") for q in pre] +
                       [r["data"].get(q["id"], "") for q in quick])

    ws2.column_dimensions["A"].width = 34
    ws2.column_dimensions["B"].width = 24
    for c in range(3, 12):
        ws2.column_dimensions[get_column_letter(c)].width = 20

    return wb


def _month_slug(m):
    """माह → फ़ाइलनाम के लिए साफ़ रूप (जैसे 2026-08)"""
    if not m or m == "all":
        return "all-months"
    try:
        i = MONTHS.index(m)
        return f"{2026 + (7 + i) // 12}-{(7 + i) % 12 + 1:02d}"
    except ValueError:
        return "month"


@app.route("/admin/export")
def admin_export():
    if not is_admin():
        abort(403)
    survey, cfg, month = _sel_survey_month()
    wb = build_workbook(build_summary(survey, cfg, month), survey, cfg, month)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = f"{survey}_{_month_slug(month)}_{date.today().isoformat()}.xlsx"
    return send_file(
        buf, as_attachment=True, download_name=fname,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------- रन
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port, debug=False)

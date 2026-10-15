#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Мини-система: три участка + панель директора.
Запуск:  pip install flask openpyxl  &&  python app.py
Откройте http://127.0.0.1:5000
Данные один раз импортируются из трёх xlsx в SQLite (data.db),
после этого xlsx можно удалить — система работает автономно.
"""
import io
import os
import sqlite3
from datetime import datetime
from pathlib import Path

from flask import Flask, request, redirect, url_for, flash, send_file, g

APP_DIR = Path(__file__).parent
DB_PATH = APP_DIR / "data.db"

# Кандидаты на расположение исходных xlsx (первичный импорт — один раз)
CANDIDATE_XLSX_DIRS = [APP_DIR, APP_DIR.parent]

# Хост/порт можно переопределить через окружение: PORT=5055 python app.py
HOST = os.environ.get("HOST", "127.0.0.1")
PORT = int(os.environ.get("PORT", "5000"))

app = Flask(__name__)
# Ключ для flash-сообщений: в проде задайте SECRET_KEY, иначе — случайный на запуск.
app.secret_key = os.environ.get("SECRET_KEY") or os.urandom(24)


# ---------------- DB ----------------
def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(str(DB_PATH))
        g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def close_db(exc=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    db = get_db()
    db.executescript(
        """
        CREATE TABLE IF NOT EXISTS zag (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            record_id TEXT, date TEXT, batch_id TEXT,
            employee TEXT, plan_qty INTEGER, fact_qty INTEGER
        );
        CREATE TABLE IF NOT EXISTS weld (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            record_id TEXT, date TEXT, batch_id TEXT, need_qty INTEGER
        );
        CREATE TABLE IF NOT EXISTS docs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            record_id TEXT, date TEXT, employee TEXT,
            plan_docs INTEGER, fact_docs INTEGER
        );
        """
    )
    db.commit()


def table_empty(table):
    cur = get_db().execute(f"SELECT COUNT(*) c FROM {table}")
    return cur.fetchone()["c"] == 0


# ---------------- Валидация и расчёты ----------------
# Заготовка: ошибки, исключающие запись из KPI: duplicate, missing, negative, outlier.
# zero (факт=0) — предупреждение, в KPI участвует.
# Сварка: missing/negative исключают. Документация: missing/negative исключают, zero/overplan — предупреждения.
OUTLIER_MULT = 3.0

ISSUE_LABELS = {
    "duplicate": "дубликат record_id",
    "missing": "нет факта",
    "negative": "отрицательный факт",
    "zero": "факт = 0 (простой?)",
    "outlier": "выброс: факт > 3×план",
    "bad_plan": "нет/плохой план",
    "overplan": "факт > плана",
}


def validate_zag(rows):
    """rows: list of dicts. Возвращает {id: [коды]}."""
    counts = {}
    for r in rows:
        counts[r["record_id"]] = counts.get(r["record_id"], 0) + 1
    out = {}
    for r in rows:
        iss = []
        if counts.get(r["record_id"], 0) > 1:
            iss.append("duplicate")
        f, p = r["fact_qty"], r["plan_qty"]
        if p is None or (isinstance(p, int) and p <= 0):
            iss.append("bad_plan")
        if f is None:
            iss.append("missing")
        elif f < 0:
            iss.append("negative")
        elif f == 0:
            iss.append("zero")
        elif isinstance(p, int) and p > 0 and f > OUTLIER_MULT * p:
            iss.append("outlier")
        out[r["id"]] = iss
    return out


ZAG_EXCLUDE = {"duplicate", "missing", "negative", "outlier", "bad_plan"}


def validate_weld(rows):
    out = {}
    for r in rows:
        iss = []
        n = r["need_qty"]
        if n is None:
            iss.append("missing")
        elif n < 0:
            iss.append("negative")
        elif n == 0:
            iss.append("zero")
        out[r["id"]] = iss
    return out


WELD_EXCLUDE = {"missing", "negative"}


def validate_docs(rows):
    out = {}
    for r in rows:
        iss = []
        f, p = r["fact_docs"], r["plan_docs"]
        if p is None or (isinstance(p, int) and p <= 0):
            iss.append("bad_plan")
        if f is None:
            iss.append("missing")
        elif f < 0:
            iss.append("negative")
        elif f == 0:
            iss.append("zero")
        elif isinstance(p, int) and p > 0 and f > p:
            iss.append("overplan")
        out[r["id"]] = iss
    return out


DOCS_EXCLUDE = {"missing", "negative", "bad_plan"}


def parse_date(s):
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except Exception:
        return None


def in_period(date_str, d_from, d_to):
    d = parse_date(date_str)
    if d is None:
        return False
    if d_from and d < d_from:
        return False
    if d_to and d > d_to:
        return False
    return True


def compute_kpi(d_from, d_to):
    """Все расчёты панели директора. Возвращает dict."""
    db = get_db()
    zag = [dict(r) for r in db.execute("SELECT * FROM zag ORDER BY date, id")]
    weld = [dict(r) for r in db.execute("SELECT * FROM weld ORDER BY date, id")]
    docs = [dict(r) for r in db.execute("SELECT * FROM docs ORDER BY date, id")]

    vz, vw, vd = validate_zag(zag), validate_weld(weld), validate_docs(docs)

    def filt(rows, vmap, excl):
        kept, bad = [], []
        for r in rows:
            if not in_period(r["date"], d_from, d_to):
                continue
            if set(vmap.get(r["id"], [])) & excl:
                bad.append(r)
            else:
                kept.append(r)
        return kept, bad

    zag_ok, zag_bad = filt(zag, vz, ZAG_EXCLUDE)
    weld_ok, weld_bad = filt(weld, vw, WELD_EXCLUDE)
    docs_ok, docs_bad = filt(docs, vd, DOCS_EXCLUDE)

    plan_sum = sum(r["plan_qty"] or 0 for r in zag_ok)
    fact_sum = sum(r["fact_qty"] or 0 for r in zag_ok)
    need_sum = sum(r["need_qty"] or 0 for r in weld_ok)
    dplan = sum(r["plan_docs"] or 0 for r in docs_ok)
    dfact = sum(r["fact_docs"] or 0 for r in docs_ok)

    zag_exec = (fact_sum / plan_sum) if plan_sum else None
    docs_exec = (dfact / dplan) if dplan else None

    # Связка по партиям: покрытие = факт_заготовки / потребность_сварки (только валидные с обеих сторон)
    fact_by_batch, need_by_batch = {}, {}
    for r in zag_ok:
        fact_by_batch[r["batch_id"]] = fact_by_batch.get(r["batch_id"], 0) + (r["fact_qty"] or 0)
    for r in weld_ok:
        need_by_batch[r["batch_id"]] = need_by_batch.get(r["batch_id"], 0) + (r["need_qty"] or 0)

    batches = sorted(set(fact_by_batch) | set(need_by_batch))
    coverage_rows = []
    for b in batches:
        f = fact_by_batch.get(b)
        n = need_by_batch.get(b)
        zag_bad_b = [r for r in zag_bad if r["batch_id"] == b]
        weld_bad_b = [r for r in weld_bad if r["batch_id"] == b]
        has_error = bool(zag_bad_b or weld_bad_b)
        if f is None or n is None or n == 0:
            cov = None
            # пара есть, но одна сторона исключена из-за ошибки -> недостоверно, а не «нет пары»
            status = "tainted" if has_error else "unknown"
        else:
            cov = f / n
            status = "tainted" if has_error else ("ok" if cov >= 1 else "deficit")
        coverage_rows.append(
            {"batch": b, "fact": f, "need": n, "cov": cov, "status": status,
             "zag_bad": zag_bad_b, "weld_bad": weld_bad_b}
        )

    # Динамика по дням (валидные)
    def daily(rows, val_key):
        agg = {}
        for r in rows:
            agg[r["date"]] = agg.get(r["date"], 0) + (r[val_key] or 0)
        return sorted(agg.items())

    daily_zag_p = daily(zag_ok, "plan_qty")
    daily_zag_f = daily(zag_ok, "fact_qty")
    daily_docs_p = daily(docs_ok, "plan_docs")
    daily_docs_f = daily(docs_ok, "fact_docs")

    # Все проблемы периода (и вне периода — для контроля качества данных).
    # Дубликат record_id даёт одинаковый код на каждой копии: сворачиваем такие
    # повторы в одну запись проблемы, чтобы список не задваивался.
    all_issues = []
    for section, rows, vmap in (("Заготовка", zag, vz), ("Сварка", weld, vw), ("Документация", docs, vd)):
        seen = set()
        for r in rows:
            for code in vmap.get(r["id"], []):
                key = (section, r["record_id"], code)
                if key in seen:
                    continue
                seen.add(key)
                all_issues.append({"участок": section, "row": r, "code": code,
                                   "in_period": in_period(r["date"], d_from, d_to)})

    return {
        "zag_ok": zag_ok, "zag_bad": zag_bad, "weld_ok": weld_ok, "weld_bad": weld_bad,
        "docs_ok": docs_ok, "docs_bad": docs_bad,
        "plan_sum": plan_sum, "fact_sum": fact_sum, "need_sum": need_sum,
        "dplan": dplan, "dfact": dfact, "zag_exec": zag_exec, "docs_exec": docs_exec,
        "coverage_rows": coverage_rows, "daily_zag_p": daily_zag_p,
        "daily_zag_f": daily_zag_f, "daily_docs_p": daily_docs_p,
        "daily_docs_f": daily_docs_f, "issues": all_issues,
        "vz": vz, "vw": vw, "vd": vd,
    }


# ---------------- Импорт ----------------
def find_xlsx(prefix):
    """Ищет файл вида '01_*.xlsx'. Игнорирует временные файлы Excel ('~$...')."""
    for d in CANDIDATE_XLSX_DIRS:
        if not d.exists():
            continue
        for p in sorted(d.glob("*.xlsx")):
            if p.name.startswith("~$"):
                continue
            if p.name.startswith(prefix + "_") or p.name.startswith(prefix + " "):
                return p
    return None


def parse_excel_date(v):
    if v is None or v == "":
        return None
    if hasattr(v, "isoformat"):
        try:
            return v.date().isoformat() if hasattr(v, "date") else str(v)[:10]
        except Exception:
            return str(v)[:10]
    s = str(v).strip()[:10]
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except Exception:
            pass
    return s


def to_int_or_none(v):
    if v is None or v == "":
        return None
    try:
        return int(float(v))
    except Exception:
        return None


def parse_optional_int(raw, label="Значение"):
    """Пусто/None -> None. Иначе int, иначе ValueError с понятным текстом."""
    if raw is None or str(raw).strip() == "":
        return None
    try:
        return int(float(raw))
    except Exception:
        raise ValueError(f"{label} должен быть числом (или пусто = нет данных)")


def do_import():
    """Однократный импорт из трёх xlsx. Возвращает (ok_count, messages)."""
    import openpyxl

    msgs = []
    total = 0
    db = get_db()

    def load_sheet(path, sheet="Данные"):
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[sheet] if sheet in wb.sheetnames else wb.active
        rows = list(ws.iter_rows(values_only=True))
        return rows[0], rows[1:]

    # 1. Заготовка
    p = find_xlsx("01")
    if p and table_empty("zag"):
        header, data = load_sheet(p)
        idx = {h: i for i, h in enumerate(header)}
        n = 0
        for row in data:
            if row[idx["record_id"]] in (None, ""):
                continue
            db.execute(
                "INSERT INTO zag (record_id,date,batch_id,employee,plan_qty,fact_qty) VALUES (?,?,?,?,?,?)",
                (str(row[idx["record_id"]]), parse_excel_date(row[idx["date"]]),
                 str(row[idx["batch_id"]]), str(row[idx["employee"]]),
                 to_int_or_none(row[idx["plan_qty"]]), to_int_or_none(row[idx["fact_qty"]])),
            )
            n += 1
        total += n
        msgs.append(f"Заготовка: загружено {n} записей из {p.name}")
    elif not table_empty("zag"):
        msgs.append("Заготовка: уже есть данные — повторный импорт пропущен")
    else:
        msgs.append("Заготовка: файл 01_*.xlsx не найден")

    # 2. Сварка
    p = find_xlsx("02")
    if p and table_empty("weld"):
        header, data = load_sheet(p)
        idx = {h: i for i, h in enumerate(header)}
        n = 0
        for row in data:
            if row[idx["record_id"]] in (None, ""):
                continue
            db.execute(
                "INSERT INTO weld (record_id,date,batch_id,need_qty) VALUES (?,?,?,?)",
                (str(row[idx["record_id"]]), parse_excel_date(row[idx["date"]]),
                 str(row[idx["batch_id"]]), to_int_or_none(row[idx["need_qty"]])),
            )
            n += 1
        total += n
        msgs.append(f"Сварка: загружено {n} записей из {p.name}")
    elif not table_empty("weld"):
        msgs.append("Сварка: уже есть данные — повторный импорт пропущен")
    else:
        msgs.append("Сварка: файл 02_*.xlsx не найден")

    # 3. Документация
    p = find_xlsx("03")
    if p and table_empty("docs"):
        header, data = load_sheet(p)
        idx = {h: i for i, h in enumerate(header)}
        n = 0
        for row in data:
            if row[idx["record_id"]] in (None, ""):
                continue
            db.execute(
                "INSERT INTO docs (record_id,date,employee,plan_docs,fact_docs) VALUES (?,?,?,?,?)",
                (str(row[idx["record_id"]]), parse_excel_date(row[idx["date"]]),
                 str(row[idx["employee"]]), to_int_or_none(row[idx["plan_docs"]]),
                 to_int_or_none(row[idx["fact_docs"]])),
            )
            n += 1
        total += n
        msgs.append(f"Документация: загружено {n} записей из {p.name}")
    elif not table_empty("docs"):
        msgs.append("Документация: уже есть данные — повторный импорт пропущен")
    else:
        msgs.append("Документация: файл 03_*.xlsx не найден")

    db.commit()
    return total, msgs


# ---------------- SVG-графики (без внешних библиотек) ----------------
def svg_grouped_bars(title, labels, s1, s2, name1="План", name2="Факт"):
    W, H, pad = 640, 260, 40
    vals = (s1 or []) + (s2 or [])
    vmax = max(vals) if vals else 1
    if vmax <= 0:
        vmax = 1
    n = max(len(labels), 1)
    gw = (W - pad - 20) / n
    bw = min(28, (gw - 14) / 2)
    out = [f'<svg viewBox="0 0 {W} {H}" style="width:100%;background:#fff;border:1px solid #e5e7eb;border-radius:8px">']
    out.append(f'<text x="12" y="20" font-size="13" font-weight="bold">{title}</text>')
    for i, lab in enumerate(labels):
        x = pad + i * gw
        for j, (series, color) in enumerate([(s1, "#93c5fd"), (s2, "#2563eb")]):
            v = (series[i] if series and i < len(series) else 0) or 0
            h = (H - 70) * v / vmax
            bx = x + 6 + j * (bw + 4)
            by = H - 34 - h
            out.append(f'<rect x="{bx:.1f}" y="{by:.1f}" width="{bw:.1f}" height="{h:.1f}" fill="{color}" rx="2"><title>{lab}: {v}</title></rect>')
        short = (lab or "")[5:]  # ММ-ДД
        out.append(f'<text x="{x + 6:.1f}" y="{H - 18}" font-size="10" fill="#6b7280">{short}</text>')
    out.append(f'<text x="12" y="{H - 4}" font-size="10" fill="#6b7280">max={vmax}</text>')
    out.append(f'<circle cx="{W-170}" cy="16" r="5" fill="#93c5fd"/><text x="{W-160}" y="20" font-size="11">{name1}</text>')
    out.append(f'<circle cx="{W-90}" cy="16" r="5" fill="#2563eb"/><text x="{W-80}" y="20" font-size="11">{name2}</text>')
    out.append("</svg>")
    return "".join(out)


def svg_coverage(rows):
    W, H, pad = 640, 260, 40
    labels = [r["batch"] for r in rows]
    n = max(len(labels), 1)
    gw = (W - pad - 20) / n
    bw = min(34, gw - 20)
    vmax = 1
    vals = []
    for r in rows:
        v = r["cov"] if r["cov"] is not None else 0
        vals.append(v)
        vmax = max(vmax, v)
    vmax = max(vmax, 1.0)
    out = [f'<svg viewBox="0 0 {W} {H}" style="width:100%;background:#fff;border:1px solid #e5e7eb;border-radius:8px">']
    out.append('<text x="12" y="20" font-size="13" font-weight="bold">Покрытие потребности сварки фактом заготовки (по партиям)</text>')
    y100 = (H - 34) - (H - 70) * 1.0 / vmax
    out.append(f'<line x1="{pad}" y1="{y100:.1f}" x2="{W-10}" y2="{y100:.1f}" stroke="#16a34a" stroke-dasharray="5 3"/>')
    out.append(f'<text x="{W-60}" y="{y100 - 5:.1f}" font-size="10" fill="#16a34a">100%</text>')
    colors = {"ok": "#16a34a", "deficit": "#dc2626", "tainted": "#f59e0b", "unknown": "#9ca3af"}
    for i, r in enumerate(rows):
        v = vals[i]
        h = (H - 70) * v / vmax
        x = pad + i * gw + 8
        c = colors.get(r["status"], "#9ca3af")
        cov_txt = "—" if r["cov"] is None else f'{r["cov"]:.0%}'
        out.append(f'<rect x="{x:.1f}" y="{H - 34 - h:.1f}" width="{bw:.1f}" height="{h:.1f}" fill="{c}" rx="2"><title>{r["batch"]}: покрытие={cov_txt} статус={r["status"]}</title></rect>')
        out.append(f'<text x="{x:.1f}" y="{H - 18}" font-size="10" fill="#374151">{r["batch"]}</text>')
        txt = "—" if r["cov"] is None else f'{r["cov"]:.0%}'
        out.append(f'<text x="{x:.1f}" y="{H - 40 - h:.1f}" font-size="10" font-weight="bold" fill="#111827">{txt}</text>')
    out.append("</svg>")
    return "".join(out)


# ---------------- Шаблоны ----------------
BASE = """
<!doctype html><html lang="ru"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{{title}}</title>
<style>
body{font-family:-apple-system,'Segoe UI',Roboto,Arial,sans-serif;margin:0;background:#f3f4f6;color:#111827}
nav{background:#111827;color:#fff;padding:10px 16px;display:flex;gap:14px;flex-wrap:wrap;align-items:center}
nav a{color:#d1d5db;text-decoration:none;padding:6px 10px;border-radius:6px}
nav a.active,nav a:hover{background:#374151;color:#fff}
.wrap{max-width:1100px;margin:0 auto;padding:18px}
.card{background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:14px;margin:12px 0}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px}
.kpi{background:#f9fafb;border:1px solid #e5e7eb;border-radius:8px;padding:10px}
.kpi b{font-size:20px;display:block}
table{border-collapse:collapse;width:100%;font-size:13px}
th,td{border:1px solid #e5e7eb;padding:6px 8px;text-align:left}
th{background:#f9fafb}
.badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:12px;margin:1px}
.err{background:#fee2e2;color:#991b1b}.warn{background:#fef3c7;color:#92400e}.okb{background:#dcfce7;color:#166534}.neut{background:#e5e7eb;color:#374151}
input,select{padding:6px 8px;border:1px solid #d1d5db;border-radius:6px}
button,.btn{background:#2563eb;color:#fff;border:0;border-radius:6px;padding:7px 12px;cursor:pointer;text-decoration:none;display:inline-block}
button.danger{background:#dc2626}button.gray{background:#6b7280}
.flash{background:#fef9c3;border:1px solid #facc15;padding:8px 12px;border-radius:8px;margin:8px 0}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:12px}@media(max-width:800px){.grid2{grid-template-columns:1fr}}
.small{font-size:12px;color:#6b7280}
form.inline{display:inline}
</style></head><body>
<nav>
<b>🏭 Мини-система</b>
<a href="/director" class="{{'active' if active=='director' else ''}}">Панель директора</a>
<a href="/zag" class="{{'active' if active=='zag' else ''}}">Заготовка</a>
<a href="/weld" class="{{'active' if active=='weld' else ''}}">Сварка</a>
<a href="/docs" class="{{'active' if active=='docs' else ''}}">Документация</a>
<span style="margin-left:auto" class="small">SQLite · авто-пересчёт · после импорта xlsx не нужны</span>
</nav>
<div class="wrap">
{% with m=get_flashed_messages() %}{% if m %}{% for x in m %}<div class="flash">{{x}}</div>{% endfor %}{% endif %}{% endwith %}
{{body|safe}}
</div></body></html>
"""

from jinja2 import Template


def page(title, active, body):
    return Template(BASE).render(title=title, active=active, body=body, get_flashed_messages=__import__("flask").get_flashed_messages)


def badge(code):
    if code in ("duplicate", "missing", "negative", "outlier", "bad_plan"):
        cls = "err"
    else:
        cls = "warn"
    return f'<span class="badge {cls}">{ISSUE_LABELS.get(code, code)}</span>'


def period_form(d_from, d_to):
    return f"""
    <div class="card"><form method="get">
    Период: <input type="date" name="from" value="{d_from or ''}"> — <input type="date" name="to" value="{d_to or ''}">
    <button>Показать</button>
    <a class="btn" style="background:#059669" href="/export?from={d_from or ''}&to={d_to or ''}&format=xlsx">Скачать отчёт Excel</a>
    <a class="btn" style="background:#6b7280" href="/export?from={d_from or ''}&to={d_to or ''}&format=csv">CSV</a>
    </form></div>"""


def parse_period(args):
    f = args.get("from") or None
    t = args.get("to") or None
    df = parse_date(f) if f else None
    dt = parse_date(t) if t else None
    return (f or "", t or "", df, dt)


# ---------------- Маршруты ----------------
@app.route("/")
def index():
    return redirect(url_for("director"))


@app.route("/director")
def director():
    init_db()
    if table_empty("zag"):
        do_import()
    fs, ts, df, dt = parse_period(request.args)
    k = compute_kpi(df, dt)
    focus = request.args.get("batch") or ""

    def pct(x):
        return "—" if x is None else f"{x:.0%}"

    # KPI-карточки
    cards = f"""
    <div class="card"><h2>Панель генерального директора</h2>
    <div class="small">Расчёты — только по <b>достоверным</b> записям периода.
    Ошибочные записи исключены из итогов (а не обнулены): итог помечается недостоверным, если в его партии есть ошибка.
    Документация — независимый участок: её ошибки не влияют на заготовку/сварку и наоборот.</div></div>
    {period_form(fs, ts)}
    <div class="kpis">
      <div class="kpi">Заготовка: план<b>{k['plan_sum']}</b><span class="small">сумма планов достоверных</span></div>
      <div class="kpi">Заготовка: факт<b>{k['fact_sum']}</b><span class="small">сумма фактов достоверных</span></div>
      <div class="kpi">Выполнение плана<b>{pct(k['zag_exec'])}</b><span class="small">факт / план (не среднее %)</span></div>
      <div class="kpi">Потребность сварки<b>{k['need_sum']}</b><span class="small">сумма потребностей</span></div>
      <div class="kpi">Документы: план / факт<b>{k['dplan']} / {k['dfact']}</b><span class="small">{pct(k['docs_exec'])}</span></div>
      <div class="kpi">Исключено из итогов<b>{len(k['zag_bad'])} + {len(k['weld_bad'])} + {len(k['docs_bad'])}</b><span class="small">заг + свар + док</span></div>
    </div>"""

    # графики
    zl = [d for d, _ in k["daily_zag_p"]]
    # объединить даты двух рядов
    zdates = sorted(set([d for d, _ in k["daily_zag_p"]] + [d for d, _ in k["daily_zag_f"]]))
    m1 = dict(k["daily_zag_p"]); m2 = dict(k["daily_zag_f"])
    ch1 = svg_grouped_bars("Заготовка по дням: план vs факт (достоверные)", zdates,
                           [m1.get(d, 0) for d in zdates], [m2.get(d, 0) for d in zdates])
    ddates = sorted(set([d for d, _ in k["daily_docs_p"]] + [d for d, _ in k["daily_docs_f"]]))
    n1 = dict(k["daily_docs_p"]); n2 = dict(k["daily_docs_f"])
    ch2 = svg_grouped_bars("Документация по дням: план vs факт", ddates,
                           [n1.get(d, 0) for d in ddates], [n2.get(d, 0) for d in ddates],
                           "План док.", "Факт док.")
    ch3 = svg_coverage(k["coverage_rows"])

    # сводная таблица по партиям
    rows_html = []
    for r in k["coverage_rows"]:
        st = {"ok": '<span class="badge okb">покрыта</span>', "deficit": '<span class="badge err">дефицит</span>',
              "tainted": '<span class="badge warn">недостоверно — есть ошибка</span>',
              "unknown": '<span class="badge neut">нет пары</span>'}[r["status"]]
        cov = "—" if r["cov"] is None else f'{r["cov"]:.0%}'
        hl = " style='background:#eff6ff'" if focus == r["batch"] else ""
        rows_html.append(f"<tr{hl}><td><a href='/director?from={fs}&to={ts}&batch={r['batch']}'>{r['batch']}</a></td>"
                         f"<td>{r['fact'] if r['fact'] is not None else '—'}</td><td>{r['need'] if r['need'] is not None else '—'}</td>"
                         f"<td>{cov}</td><td>{st}</td></tr>")
    summary = f"""<div class="card"><h3>Сводная таблица по партиям (из каких записей получился итог — нажмите на партию)</h3>
    <table><tr><th>Партия</th><th>Факт заготовки</th><th>Потребность сварки</th><th>Покрытие</th><th>Статус</th></tr>
    {''.join(rows_html)}</table>
    <div class="small">Легенда: <span class="badge okb">покрыта ≥100%</span>
    <span class="badge err">дефицит &lt;100%</span>
    <span class="badge warn">недостоверно — в партии есть исключённая запись</span>
    <span class="badge neut">нет пары — связи нет, не придумываем</span></div></div>"""

    # детализация по выбранной партии
    detail = ""
    if focus:
        db = get_db()
        zr = [dict(x) for x in db.execute("SELECT * FROM zag WHERE batch_id=? ORDER BY id", (focus,))]
        wr = [dict(x) for x in db.execute("SELECT * FROM weld WHERE batch_id=? ORDER BY id", (focus,))]
        # Дубликаты считаются по всей таблице, поэтому валидируем полные выборки,
        # а затем показываем флаги только для строк выбранной партии.
        allz = [dict(x) for x in db.execute("SELECT * FROM zag")]
        allw = [dict(x) for x in db.execute("SELECT * FROM weld")]
        vz = validate_zag(allz)
        vw = validate_weld(allw)
        OK_BADGE = '<span class="badge okb">ок</span>'
        zhtml = "".join(
            "<tr><td>" + str(r['record_id']) + "</td><td>" + str(r['date']) + "</td><td>" + str(r['employee']) + "</td>"
            + "<td>" + str(r['plan_qty']) + "</td><td>" + (str(r['fact_qty']) if r['fact_qty'] is not None else '—') + "</td>"
            + "<td>" + (' '.join(badge(c) for c in vz.get(r['id'], [])) or OK_BADGE) + "</td></tr>"
            for r in zr)
        whtml = "".join(
            "<tr><td>" + str(r['record_id']) + "</td><td>" + str(r['date']) + "</td><td>" + str(r['need_qty']) + "</td>"
            + "<td>" + (' '.join(badge(c) for c in vw.get(r['id'], [])) or OK_BADGE) + "</td></tr>"
            for r in wr)
        detail = f"""<div class="card"><h3>Детализация партии {focus}</h3>
        <b>Заготовка:</b><table><tr><th>ID</th><th>Дата</th><th>Сотрудник</th><th>План</th><th>Факт</th><th>Проверка</th></tr>{zhtml}</table><br>
        <b>Сварка:</b><table><tr><th>ID</th><th>Дата</th><th>Потребность</th><th>Проверка</th></tr>{whtml}</table>
        <div class="small">Исключённые записи в итог партии не входят — поэтому статус партии «недостоверно», а не «достоверный дефицит/покрытие».</div></div>"""

    # качество данных
    per_issues = [i for i in k["issues"] if i["in_period"]]
    iss_html = []
    for i in per_issues:
        r = i["row"]
        rid = r.get("record_id")
        link = {"Заготовка": "/zag", "Сварка": "/weld", "Документация": "/docs"}[i["участок"]]
        iss_html.append(f"<tr><td>{i['участок']}</td><td>{rid}</td><td>{r.get('date')}</td><td>{badge(i['code'])}</td>"
                        f"<td><a href='{link}'>исправить</a></td></tr>")
    quality = f"""<div class="card"><h3>⚠️ Проверка ошибок ({len(per_issues)} в периоде)</h3>
    <table><tr><th>Участок</th><th>Запись</th><th>Дата</th><th>Проблема</th><th></th></tr>
    {''.join(iss_html) or '<tr><td colspan=5>Ошибок в периоде нет</td></tr>'}</table>
    <div class="small">Правило: ошибка одного участка исключает только свои записи.
    Связанный показатель (покрытие) помечается «недостоверно», независимый участок (документация ↔ заготовка/сварка) не меняется.</div></div>"""

    body = (cards + f'<div class="grid2"><div>{ch1}</div><div>{ch2}</div></div>'
            + f'<div style="margin-top:12px">{ch3}</div>' + summary + detail + quality)
    return page("Панель директора", "director", body)


def editable_table(title, active, columns, rows, vmap, update_url, add_fields,
                   extra_note="", edit_fields=(("fact", "факт"),)):
    """edit_fields: список пар (имя_поля_формы, placeholder) для inline-правки."""
    trs = []
    for r in rows:
        flags = " ".join(badge(c) for c in vmap.get(r["id"], [])) or '<span class="badge okb">ок</span>'
        tds = "".join(f"<td>{r[c] if r[c] is not None else '—'}</td>" for c in columns)
        inputs = ""
        for field, ph in edit_fields:
            val = r.get(field + "_val")
            inputs += (f"<input name='{field}' size='6' placeholder='{ph}' "
                       f"value='{val if val is not None else ''}'> ")
        trs.append(f"<tr id='row{r['id']}'><td>{r['id']}</td>{tds}<td>{flags}</td>"
                   f"<td><form class='inline' method='post' action='{update_url}'>"
                   f"<input type='hidden' name='id' value='{r['id']}'>{inputs}"
                   f"<button>💾</button></form> "
                   f"<form class='inline' method='post' action='{update_url.replace('update', 'delete')}' "
                   f"onsubmit=\"return confirm('Удалить запись?')\">"
                   f"<input type='hidden' name='id' value='{r['id']}'><button class='danger'>✕</button></form></td></tr>")
    add_html = " ".join(f"<input name='{n}' placeholder='{n}' size='10'>" for n in add_fields)
    edit_hdr = " / ".join(ph for _, ph in edit_fields)
    return f"""<div class="card"><h2>{title}</h2><div class="small">{extra_note}</div></div>
    <div class="card"><table><tr><th>#</th>{''.join(f'<th>{c}</th>' for c in columns)}<th>Проверка</th><th>Правка ({edit_hdr}) / удаление</th></tr>
    {''.join(trs)}</table></div>
    <div class="card"><h3>＋ Новая запись</h3><form method="post" action="{update_url.replace('update', 'add')}">
    {add_html} <button>Добавить</button></form></div>"""


@app.route("/zag", methods=["GET"])
def zag_list():
    init_db()
    if table_empty("zag"):
        do_import()
    db = get_db()
    rows = [dict(r) for r in db.execute("SELECT * FROM zag ORDER BY date, id")]
    vz = validate_zag(rows)
    flat = [{"id": r["id"], "record_id": r["record_id"], "date": r["date"], "batch_id": r["batch_id"],
             "employee": r["employee"], "plan_qty": r["plan_qty"], "fact_qty": r["fact_qty"],
             "plan_val": r["plan_qty"], "fact_val": r["fact_qty"]} for r in rows]
    body = editable_table("🔧 Заготовка (план/факт по партиям)", "zag",
                          ["record_id", "date", "batch_id", "employee", "plan_qty", "fact_qty"],
                          flat, vz, "/zag/update",
                          ["record_id", "date", "batch_id", "employee", "plan_qty", "fact_qty"],
                          "Ошибки (дубликат / нет факта / отрицательный / выброс &gt;3×план) исключаются из итогов директора. Факт = 0 — предупреждение, в итогах участвует.",
                          edit_fields=(("plan", "план"), ("fact", "факт")))
    return page("Заготовка", "zag", body)


@app.route("/zag/update", methods=["POST"])
def zag_update():
    db = get_db()
    rid = request.form.get("id")
    try:
        plan = parse_optional_int(request.form.get("plan"), "План")
        fact = parse_optional_int(request.form.get("fact"), "Факт")
    except ValueError as e:
        flash(str(e))
        return redirect(url_for("zag_list"))
    db.execute("UPDATE zag SET plan_qty=?, fact_qty=? WHERE id=?", (plan, fact, rid))
    db.commit()
    flash(f"Запись #{rid}: план → {plan if plan is not None else '—'}, "
          f"факт → {fact if fact is not None else '—'}. Панель директора пересчитана.")
    return redirect(url_for("zag_list"))


@app.route("/zag/add", methods=["POST"])
def zag_add():
    db = get_db()
    f = request.form
    db.execute("INSERT INTO zag (record_id,date,batch_id,employee,plan_qty,fact_qty) VALUES (?,?,?,?,?,?)",
               (f.get("record_id"), f.get("date"), f.get("batch_id"), f.get("employee"),
                to_int_or_none(f.get("plan_qty")), to_int_or_none(f.get("fact_qty"))))
    db.commit()
    flash("Запись добавлена.")
    return redirect(url_for("zag_list"))


@app.route("/zag/delete", methods=["POST"])
def zag_delete():
    db = get_db()
    db.execute("DELETE FROM zag WHERE id=?", (request.form.get("id"),))
    db.commit()
    flash("Запись удалена.")
    return redirect(url_for("zag_list"))


@app.route("/weld", methods=["GET"])
def weld_list():
    init_db()
    if table_empty("weld"):
        do_import()
    db = get_db()
    rows = [dict(r) for r in db.execute("SELECT * FROM weld ORDER BY date, id")]
    vw = validate_weld(rows)
    flat = [{"id": r["id"], "record_id": r["record_id"], "date": r["date"], "batch_id": r["batch_id"],
             "need_qty": r["need_qty"], "fact_val": r["need_qty"]} for r in rows]
    body = editable_table("🔥 Сварка (потребность по партиям)", "weld",
                          ["record_id", "date", "batch_id", "need_qty"],
                          flat, vw, "/weld/update",
                          ["record_id", "date", "batch_id", "need_qty"],
                          "Связь с заготовкой — только по batch_id. Ошибка сварки не меняет итоги заготовки, а покрытие партии становится «недостоверно».",
                          edit_fields=(("fact", "потребность"),))
    return page("Сварка", "weld", body)


@app.route("/weld/update", methods=["POST"])
def weld_update():
    db = get_db()
    try:
        need = parse_optional_int(request.form.get("fact"), "Потребность")
    except ValueError as e:
        flash(str(e))
        return redirect(url_for("weld_list"))
    db.execute("UPDATE weld SET need_qty=? WHERE id=?", (need, request.form.get("id")))
    db.commit()
    flash("Потребность обновлена, панель директора пересчитана.")
    return redirect(url_for("weld_list"))


@app.route("/weld/add", methods=["POST"])
def weld_add():
    db = get_db()
    f = request.form
    db.execute("INSERT INTO weld (record_id,date,batch_id,need_qty) VALUES (?,?,?,?)",
               (f.get("record_id"), f.get("date"), f.get("batch_id"), to_int_or_none(f.get("need_qty"))))
    db.commit()
    flash("Запись добавлена.")
    return redirect(url_for("weld_list"))


@app.route("/weld/delete", methods=["POST"])
def weld_delete():
    db = get_db()
    db.execute("DELETE FROM weld WHERE id=?", (request.form.get("id"),))
    db.commit()
    flash("Запись удалена.")
    return redirect(url_for("weld_list"))


@app.route("/docs", methods=["GET"])
def docs_list():
    init_db()
    if table_empty("docs"):
        do_import()
    db = get_db()
    rows = [dict(r) for r in db.execute("SELECT * FROM docs ORDER BY date, id")]
    vd = validate_docs(rows)
    flat = [{"id": r["id"], "record_id": r["record_id"], "date": r["date"], "employee": r["employee"],
             "plan_docs": r["plan_docs"], "fact_docs": r["fact_docs"],
             "plan_val": r["plan_docs"], "fact_val": r["fact_docs"]} for r in rows]
    body = editable_table("📄 Документация (независимый участок)", "docs",
                          ["record_id", "date", "employee", "plan_docs", "fact_docs"],
                          flat, vd, "/docs/update",
                          ["record_id", "date", "employee", "plan_docs", "fact_docs"],
                          "Участок независим: его ошибки и правки не влияют на заготовку и сварку.",
                          edit_fields=(("plan", "план"), ("fact", "факт")))
    return page("Документация", "docs", body)


@app.route("/docs/update", methods=["POST"])
def docs_update():
    db = get_db()
    try:
        plan = parse_optional_int(request.form.get("plan"), "План")
        fact = parse_optional_int(request.form.get("fact"), "Факт")
    except ValueError as e:
        flash(str(e))
        return redirect(url_for("docs_list"))
    db.execute("UPDATE docs SET plan_docs=?, fact_docs=? WHERE id=?", (plan, fact, request.form.get("id")))
    db.commit()
    flash("Документы обновлены, панель директора пересчитана.")
    return redirect(url_for("docs_list"))


@app.route("/docs/add", methods=["POST"])
def docs_add():
    db = get_db()
    f = request.form
    db.execute("INSERT INTO docs (record_id,date,employee,plan_docs,fact_docs) VALUES (?,?,?,?,?)",
               (f.get("record_id"), f.get("date"), f.get("employee"),
                to_int_or_none(f.get("plan_docs")), to_int_or_none(f.get("fact_docs"))))
    db.commit()
    flash("Запись добавлена.")
    return redirect(url_for("docs_list"))


@app.route("/docs/delete", methods=["POST"])
def docs_delete():
    db = get_db()
    db.execute("DELETE FROM docs WHERE id=?", (request.form.get("id"),))
    db.commit()
    flash("Запись удалена.")
    return redirect(url_for("docs_list"))


@app.route("/export")
def export():
    init_db()
    fs, ts, df, dt = parse_period(request.args)
    fmt = (request.args.get("format", "xlsx") or "xlsx").lower()
    if fmt not in ("xlsx", "csv"):
        flash(f"Неизвестный формат отчёта: «{fmt}». Доступны: xlsx, csv.")
        return redirect(url_for("director", **{"from": fs, "to": ts}))
    k = compute_kpi(df, dt)
    label = f"{fs or 'all'}_{ts or 'all'}"

    if fmt == "csv":
        import csv as csvm
        buf = io.StringIO()
        w = csvm.writer(buf)
        w.writerow(["batch", "fact_zag", "need_weld", "coverage", "status"])
        for r in k["coverage_rows"]:
            w.writerow([r["batch"], r["fact"], r["need"],
                        ("" if r["cov"] is None else round(r["cov"], 4)), r["status"]])
        w.writerow([])
        w.writerow(["KPI", "value"])
        w.writerow(["plan_zag", k["plan_sum"]])
        w.writerow(["fact_zag", k["fact_sum"]])
        w.writerow(["need_weld", k["need_sum"]])
        w.writerow(["plan_docs", k["dplan"]])
        w.writerow(["fact_docs", k["dfact"]])
        mem = io.BytesIO(buf.getvalue().encode("utf-8-sig"))
        return send_file(mem, as_attachment=True, download_name=f"report_{label}.csv", mimetype="text/csv")

    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "KPI"
    ws.append(["Показатель", "Значение", "Комментарий"])
    ws.append(["Период", f"{fs} — {ts}", ""])
    ws.append(["Заготовка план", k["plan_sum"], "сумма по достоверным"])
    ws.append(["Заготовка факт", k["fact_sum"], "сумма по достоверным"])
    ws.append(["Выполнение плана", round(k["zag_exec"], 4) if k["zag_exec"] is not None else "—", "факт/план"])
    ws.append(["Потребность сварки", k["need_sum"], ""])
    ws.append(["Документы план/факт", f"{k['dplan']}/{k['dfact']}", "независимый участок"])
    ws.append(["Исключено записей (заг/свар/док)", f"{len(k['zag_bad'])}/{len(k['weld_bad'])}/{len(k['docs_bad'])}", ""])
    ws2 = wb.create_sheet("Покрытие_по_партиям")
    ws2.append(["batch", "fact_zag", "need_weld", "coverage", "status"])
    for r in k["coverage_rows"]:
        ws2.append([r["batch"], r["fact"], r["need"],
                    (None if r["cov"] is None else round(r["cov"], 4)), r["status"]])
    ws3 = wb.create_sheet("Заготовка_достоверные")
    ws3.append(["record_id", "date", "batch", "employee", "plan", "fact"])
    for r in k["zag_ok"]:
        ws3.append([r["record_id"], r["date"], r["batch_id"], r["employee"], r["plan_qty"], r["fact_qty"]])
    ws4 = wb.create_sheet("Ошибки")
    ws4.append(["участок", "record_id", "date", "проблема"])
    for i in k["issues"]:
        if i["in_period"]:
            ws4.append([i["участок"], i["row"].get("record_id"), i["row"].get("date"), ISSUE_LABELS.get(i["code"], i["code"])])
    mem = io.BytesIO()
    wb.save(mem)
    mem.seek(0)
    return send_file(mem, as_attachment=True, download_name=f"report_{label}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/health")
def health():
    return "ok"


if __name__ == "__main__":
    with app.app_context():
        init_db()
        if table_empty("zag"):
            try:
                total, msgs = do_import()
                print("Импорт:", *msgs, sep="\n  ")
            except Exception as e:
                print("Импорт пропущен:", e)
    print(f"Откройте http://{HOST}:{PORT}  (порт меняется: PORT=5055 python app.py)")
    # Проверяем порт заранее: Flask/SocketServer при занятом порте роняет
    # SystemExit, и внутри app.run() это не всегда перехватывается предсказуемо.
    import socket
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
        probe.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        busy = probe.connect_ex((HOST if HOST not in ("", "0.0.0.0") else "127.0.0.1", PORT)) == 0
    if busy:
        print(f"\nПорт {HOST}:{PORT} занят (на macOS порт 5000 занимает AirPlay Receiver).")
        print(f"Запустите на другом порту:  PORT={PORT + 1} python app.py")
        raise SystemExit(1)
    try:
        app.run(host=HOST, port=PORT, debug=False)
    except KeyboardInterrupt:
        print("\nОстановлено.")


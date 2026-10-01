"""WRS Mobile — single-page FastAPI app."""
from __future__ import annotations

import io
import os
from datetime import datetime
from typing import Any, Dict, List, Optional

import pandas as pd
from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app.config import (
    ADDRESS_FILE,
    APP_TITLE,
    APP_VER,
    CORRECTIONS_FILE,
    DATA_DIR,
    DVORNIK_FILE,
    HISTORY_FILE,
    RECON_FILE,
    SETTINGS_FILE,
    TRACKER_FILE,
    UPLOAD_DIR,
)
from app.services.dvornik import (
    build_dvornik_report,
    detect_weekday_from_shifts,
    evaluate_dvornik,
    extract_adj_from_dataframe,
    parse_schedule_excel,
)
from app.services.shifts_processor import process_shifts
from app.services.excel_export import build_shifts_excel, msk_now
from app.services.wrs_processor import (
    build_basic_report,
    process_wrs,
    region_for,
)
from app.utils.address import normalize_address
from app.utils.storage import default_settings, load_json, now_str, save_json

app = FastAPI(title=APP_TITLE, version=APP_VER)

BASE = os.path.dirname(os.path.abspath(__file__))
templates = Jinja2Templates(directory=os.path.join(BASE, "templates"))
static_dir = os.path.join(BASE, "static")
os.makedirs(static_dir, exist_ok=True)
app.mount("/static", StaticFiles(directory=static_dir), name="static")

_LAST: Dict[str, Any] = {}


def get_settings() -> dict:
    data = load_json(SETTINGS_FILE, default_settings())
    if not isinstance(data, dict):
        data = default_settings()
    data.setdefault("managers", default_settings()["managers"])
    data.setdefault("regions", default_settings()["regions"])
    return data


def get_ab() -> dict:
    raw = load_json(ADDRESS_FILE, {"rows": [], "source_file": "", "updated_at": ""})
    if isinstance(raw, list):
        raw = {"rows": raw, "source_file": "", "updated_at": ""}
    if not isinstance(raw, dict):
        raw = {"rows": [], "source_file": "", "updated_at": ""}
    rows = []
    for item in raw.get("rows") or []:
        if not isinstance(item, dict):
            continue
        addr = str(item.get("Адрес магазина", "") or "").strip()
        mgr = str(item.get("Менеджер", "") or "").strip()
        if addr:
            rows.append({"Адрес магазина": addr, "Менеджер": mgr})
    raw["rows"] = rows
    return raw


def save_ab(store: dict) -> None:
    store["updated_at"] = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    save_json(ADDRESS_FILE, store)


def ab_to_df(ab: dict) -> pd.DataFrame:
    rows = ab.get("rows") or []
    if not rows:
        return pd.DataFrame(columns=["Адрес магазина", "Менеджер"])
    return pd.DataFrame(rows)


def get_corrections() -> dict:
    data = load_json(CORRECTIONS_FILE, {})
    return data if isinstance(data, dict) else {}


def save_corrections(data: dict) -> None:
    save_json(CORRECTIONS_FILE, data)


def get_recon() -> list:
    data = load_json(RECON_FILE, [])
    return data if isinstance(data, list) else []


def get_dvornik() -> list:
    data = load_json(DVORNIK_FILE, [])
    return data if isinstance(data, list) else []


def save_dvornik(rows: list) -> None:
    save_json(DVORNIK_FILE, rows)


def append_history(entry: dict) -> None:
    hist = load_json(HISTORY_FILE, [])
    if not isinstance(hist, list):
        hist = []
    hist.append(entry)
    save_json(HISTORY_FILE, hist[-200:])


def page_ctx(request: Request, error: str = "") -> dict:
    settings = get_settings()
    ab = get_ab()
    wrs = _LAST.get("wrs") or {}
    shifts = _LAST.get("shifts") or {}
    dv = _LAST.get("dvornik") or {}
    corr = get_corrections()
    managers = settings.get("managers") or []
    regions = settings.get("regions") or {}

    corr_items = []
    mgr_map = {r.get("manager"): r for r in (wrs.get("mgr_rows") or [])}
    for m in managers:
        c = corr.get(m) or {}
        row = mgr_map.get(m) or {}
        corr_items.append(
            {
                "manager": m,
                "region": region_for(m, regions),
                "total": row.get("total", 0),
                "sl_percent": row.get("sl_percent", "—"),
                "q": int(c.get("q", 0) or 0),
                "nr": int(c.get("nr", 0) or 0),
                "tech": int(c.get("tech", 0) or 0),
            }
        )

    schedule = get_dvornik()
    return {
        "request": request,
        "title": "WRS",
        "ver": APP_VER,
        "error": error,
        "ab_count": len(ab.get("rows") or []),
        "ab_updated": ab.get("updated_at") or "",
        "schedule_count": len(schedule),
        "has_results": bool(wrs or shifts or dv),
        "has_excel": bool(_LAST.get("shifts_full")),
        "report_date": (shifts.get("shift_ctx") or {}).get("report_date") or "",
        "generated_at": _LAST.get("generated_at") or "",
        "kpi": _LAST.get("kpi") or {},
        # 1. внутри магазина
        "wrs_report": _LAST.get("basic_report") or "",
        "wrs_report_sl": _LAST.get("basic_report_sl") or "",
        "wrs_mgr_rows": wrs.get("mgr_rows") or [],
        "wrs_region_rows": wrs.get("region_rows") or [],
        # 2. дворники
        "dv_day": dv.get("day_label") or shifts.get("day_label") or "",
        "dv_n_total": dv.get("n_total"),
        "dv_n_worked": dv.get("n_worked"),
        "dv_n_nonauth": dv.get("n_nonauth"),
        "dv_nonauth": dv.get("nonauth") or [],
        "dv_mgr_rows": dv.get("mgr_rows") or [],
        "dv_region_rows": dv.get("region_rows") or [],
        "dv_report": dv.get("report_text") or "",
        "dv_report_sl": dv.get("report_text_sl") or "",
        # 3. короткие смены
        "short_rows": shifts.get("report_rows") or [],
        "short_ctx": shifts.get("shift_ctx") or {},
        # 4. корректировки
        "corr_items": corr_items,
        "managers": managers,
        "regions": regions,
    }


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", page_ctx(request))


@app.post("/process")
async def process_all(
    request: Request,
    file_ab: UploadFile = File(...),
    file_wrs: UploadFile = File(...),
    file_shifts: UploadFile = File(...),
    file_dvornik: Optional[UploadFile] = File(None),
):
    """Один шаг: 3 файла (+ опционально график дворников) → все отчёты."""
    errors: List[str] = []

    # ── 1. Адресная книга ──────────────────────────────────────────────
    try:
        ab_bytes = await file_ab.read()
        ab_df_raw = pd.read_excel(io.BytesIO(ab_bytes))
        req = {"Адрес магазина", "Менеджер"}
        if not req.issubset(set(ab_df_raw.columns)):
            raise ValueError("Адресная книга: нужны колонки «Адрес магазина», «Менеджер»")
        clean = ab_df_raw[["Адрес магазина", "Менеджер"]].copy()
        clean["Адрес магазина"] = clean["Адрес магазина"].fillna("").astype(str).str.strip()
        clean["Менеджер"] = clean["Менеджер"].fillna("").astype(str).str.strip()
        clean = clean[clean["Адрес магазина"] != ""].copy()
        clean["_n"] = clean["Адрес магазина"].apply(normalize_address)
        rows = (
            clean.drop_duplicates(subset=["_n"], keep="last")[["Адрес магазина", "Менеджер"]]
            .to_dict("records")
        )
        ab_store = {
            "rows": rows,
            "source_file": file_ab.filename or "",
            "updated_at": datetime.now().strftime("%d.%m.%Y %H:%M:%S"),
        }
        save_ab(ab_store)
        ab_df = ab_to_df(ab_store)
    except Exception as e:
        return templates.TemplateResponse(
            "index.html",
            page_ctx(request, error=f"Адресная книга: {e}"),
            status_code=400,
        )

    settings = get_settings()
    corr = get_corrections()
    recon = get_recon()
    regions = settings.get("regions") or {}

    # ── опционально: график дворников ──────────────────────────────────
    if file_dvornik is not None and getattr(file_dvornik, "filename", None):
        try:
            dv_bytes = await file_dvornik.read()
            if dv_bytes:
                dv_df = pd.read_excel(io.BytesIO(dv_bytes))
                save_dvornik(parse_schedule_excel(dv_df))
        except Exception as e:
            errors.append(f"График дворников: {e}")

    schedule = get_dvornik()

    # ── 2. WRS — не авторизации внутри магазина ────────────────────────
    wrs_result: dict = {}
    basic = ""
    try:
        wrs_bytes = await file_wrs.read()
        wrs_df = pd.read_excel(io.BytesIO(wrs_bytes))
        wrs_result = process_wrs(
            wrs_df,
            ab_df,
            recon,
            regions,
            {},  # tracker
            corrections=corr,
        )
        basic = build_basic_report(wrs_result, now_str(), with_sl=False)
        basic_sl = build_basic_report(wrs_result, now_str(), with_sl=True)
        total_na = sum(len(v) for v in (wrs_result.get("non_auth") or {}).values())
        total_adj = sum(len(v) for v in (wrs_result.get("adj_only") or {}).values())
        shops_count = ab_df["Адрес магазина"].apply(normalize_address).nunique()
        _LAST["wrs"] = wrs_result
        _LAST["basic_report"] = basic
        _LAST["basic_report_sl"] = basic_sl
        _LAST["kpi"] = {
            "shops": int(shops_count),
            "nonauth": total_na,
            "adj": total_adj,
        }
    except Exception as e:
        errors.append(f"WRS: {e}")
        _LAST["wrs"] = {}
        _LAST["basic_report"] = ""

    # ── 3. Смены — короткие + данные для дворников ─────────────────────
    shifts_light: dict = {}
    try:
        sh_bytes = await file_shifts.read()
        sh_df = pd.read_excel(io.BytesIO(sh_bytes))
        sh_result = process_shifts(sh_df, ab_df, recon)
        adj_rows = extract_adj_from_dataframe(sh_df)
        sc = (sh_result.get("shift_ctx") or {}).get("sc")
        weekday, day_label = detect_weekday_from_shifts(sh_df, sc)
        shifts_light = {
            "report_rows": sh_result.get("report_rows") or [],
            "short_items": sh_result.get("short_items") or [],
            "shift_ctx": sh_result.get("shift_ctx") or {},
            "adj_rows": adj_rows,
            "weekday": weekday,
            "day_label": day_label,
        }
        _LAST["shifts"] = shifts_light
        _LAST["shifts_full"] = {
            "shift_details": sh_result.get("shift_details"),
            "shift_adj": sh_result.get("shift_adj"),
            "shift_ctx": sh_result.get("shift_ctx") or {},
        }
        _LAST["generated_at"] = now_str()
    except Exception as e:
        errors.append(f"Смены: {e}")
        _LAST["shifts"] = {}
        shifts_light = {}

    # ── 4. Дворники (если есть график) ─────────────────────────────────
    if schedule and shifts_light:
        try:
            weekday = shifts_light.get("weekday")
            if weekday is None:
                weekday = datetime.now().weekday()
            day_label = shifts_light.get("day_label") or datetime.now().strftime("%d.%m.%Y")
            dv_res = evaluate_dvornik(
                schedule,
                shifts_light.get("adj_rows") or [],
                int(weekday),
                ab_df,
                regions=regions,
            )
            dv_res["day_label"] = day_label
            dv_res["report_text"] = build_dvornik_report(dv_res, day_label, with_sl=False)
            dv_res["report_text_sl"] = build_dvornik_report(dv_res, day_label, with_sl=True)
            _LAST["dvornik"] = dv_res
        except Exception as e:
            errors.append(f"Дворники: {e}")
            _LAST["dvornik"] = {}
    elif not schedule:
        _LAST["dvornik"] = {}
        errors.append("График дворников не загружен — блок «Прилегающая» пропущен (можно добавить файл графика)")

    append_history(
        {
            "generated_at": now_str(),
            "file_ab": file_ab.filename or "",
            "file_wrs": file_wrs.filename or "",
            "file_shifts": file_shifts.filename or "",
            "errors": errors,
        }
    )

    err_msg = "; ".join(errors) if errors else ""
    return templates.TemplateResponse("index.html", page_ctx(request, error=err_msg))


@app.post("/corrections/save")
async def corrections_save(
    manager: str = Form(...),
    q: int = Form(0),
    nr: int = Form(0),
    tech: int = Form(0),
):
    data = get_corrections()
    data[manager] = {"q": max(0, int(q)), "nr": max(0, int(nr)), "tech": max(0, int(tech))}
    save_corrections(data)

    # пересчитать SL в последнем WRS, если есть
    last = _LAST.get("wrs")
    if last:
        settings = get_settings()
        regions = settings.get("regions") or {}
        total = last.get("total") or {}
        non_auth = last.get("non_auth") or {}
        adj_only = last.get("adj_only") or {}
        from app.services.wrs_processor import sl_manager, sl_region
        from app.utils.helpers import fmt_sl

        mgr_rows = []
        for m in settings.get("managers") or []:
            if m not in total and m not in non_auth and m not in adj_only:
                continue
            sl = sl_manager(m, total, non_auth, adj_only, data)
            mgr_rows.append(
                {
                    "manager": m,
                    "region": region_for(m, regions),
                    "total": int(total.get(m, 0) or 0),
                    "non_auth": len(non_auth.get(m, []) or []),
                    "adj_only": len(adj_only.get(m, []) or []),
                    "sl": sl,
                    "sl_percent": fmt_sl(sl),
                }
            )
        # keep managers that appear only in data
        seen = {r["manager"] for r in mgr_rows}
        for m in list(total.keys()) + list(non_auth.keys()) + list(adj_only.keys()):
            if m in seen:
                continue
            sl = sl_manager(m, total, non_auth, adj_only, data)
            mgr_rows.append(
                {
                    "manager": m,
                    "region": region_for(m, regions),
                    "total": int(total.get(m, 0) or 0),
                    "non_auth": len(non_auth.get(m, []) or []),
                    "adj_only": len(adj_only.get(m, []) or []),
                    "sl": sl,
                    "sl_percent": fmt_sl(sl),
                }
            )
        reg_rows = []
        for region, mgrs in regions.items():
            slr = sl_region(mgrs, total, non_auth, adj_only, data)
            reg_rows.append({"region": region, "sl": slr, "sl_percent": fmt_sl(slr)})
        last["mgr_rows"] = mgr_rows
        last["region_rows"] = reg_rows
        last["corrections"] = data
        _LAST["wrs"] = last
        _LAST["basic_report"] = build_basic_report(last, now_str(), with_sl=False)
        _LAST["basic_report_sl"] = build_basic_report(last, now_str(), with_sl=True)

    return RedirectResponse("/", status_code=303)




@app.get("/export/shifts")
async def export_shifts():
    full = _LAST.get("shifts_full")
    if not full or full.get("shift_details") is None:
        return JSONResponse({"ok": False, "error": "Сначала сформируйте отчёты (файл смен)"}, status_code=400)
    details = full.get("shift_details")
    if hasattr(details, "empty") and details.empty:
        return JSONResponse({"ok": False, "error": "Нет данных смен"}, status_code=400)
    ab_df = ab_to_df(get_ab())
    settings = get_settings()
    try:
        data = build_shifts_excel(
            details,
            full.get("shift_adj"),
            ab_df,
            full.get("shift_ctx") or {},
            manager_order=settings.get("managers") or [],
            regions=settings.get("regions") or {},
        )
    except Exception as e:
        return JSONResponse({"ok": False, "error": str(e)}, status_code=500)
    rd = (full.get("shift_ctx") or {}).get("report_date") or msk_now().strftime("%d.%m.%Y")
    fname = f"Smeny_{rd.replace('.', '-')}.xlsx"
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )

@app.get("/health")
async def health():
    return {"status": "ok", "version": APP_VER, "data_dir": DATA_DIR}


if __name__ == "__main__":
    import uvicorn

    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("app.main:app", host="0.0.0.0", port=port, reload=False)

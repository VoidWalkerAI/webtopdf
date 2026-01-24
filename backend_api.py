import os
import re
import uuid
import json
import datetime as dt
from typing import Optional, Dict, Any

import gspread
from google.oauth2.service_account import Credentials

from fastapi import FastAPI, Body, HTTPException, Request
from fastapi.responses import Response
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError


# =========================
# APP
# =========================
app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# =========================
# CONFIG
# =========================
SHEET_ID = os.getenv("VOYDS_FORMS_SHEET_ID", "").strip()
SERVICE_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
TOKENS_TAB = os.getenv("VOYDS_TOKENS_TAB", "WebToPDF_Tokens").strip()
ADMIN_KEY = os.getenv("ADMIN_KEY", "").strip()

TOKENS_HEADERS = [
    "token",
    "plan",
    "remaining",
    "created_at_utc",
    "last_used_at_utc",
    "note",
]


# =========================
# HELPERS
# =========================
def _utc_now_iso():
    return dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def _is_probably_url(s: str) -> bool:
    return bool(re.match(r"^https?://", (s or "").strip(), flags=re.I))


def _require_sheets_env():
    if not SHEET_ID:
        raise HTTPException(status_code=500, detail="Missing VOYDS_FORMS_SHEET_ID.")
    if not SERVICE_JSON:
        raise HTTPException(status_code=500, detail="Missing GOOGLE_SERVICE_ACCOUNT_JSON.")
    try:
        json.loads(SERVICE_JSON)
    except Exception:
        raise HTTPException(status_code=500, detail="GOOGLE_SERVICE_ACCOUNT_JSON is not valid JSON.")


def _get_gspread_client() -> gspread.Client:
    _require_sheets_env()
    sa_info = json.loads(SERVICE_JSON)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
    return gspread.authorize(creds)


def _get_tokens_ws():
    gc = _get_gspread_client()
    sh = gc.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet(TOKENS_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=TOKENS_TAB, rows=2000, cols=len(TOKENS_HEADERS) + 2)

    existing = ws.row_values(1)
    if existing != TOKENS_HEADERS:
        ws.update("A1", [TOKENS_HEADERS])

    return ws


def _mint_token() -> str:
    return uuid.uuid4().hex[:16].upper()


def _get_token_record(ws, token: str) -> Optional[Dict[str, Any]]:
    token = (token or "").strip()
    if not token:
        return None

    col = ws.col_values(1)
    try:
        row_idx = col.index(token) + 1
    except ValueError:
        return None

    if row_idx <= 1:
        return None

    row = ws.row_values(row_idx)
    rec = {TOKENS_HEADERS[i]: (row[i] if i < len(row) else "") for i in range(len(TOKENS_HEADERS))}
    rec["_row"] = row_idx
    return rec


def _parse_remaining(s: str) -> Optional[int]:
    try:
        return int((s or "").strip())
    except Exception:
        return None


def _allow_only(ws, token: str) -> Dict[str, Any]:
    rec = _get_token_record(ws, token)
    if not rec:
        raise HTTPException(status_code=402, detail="Invalid or missing token.")

    remaining = _parse_remaining(rec.get("remaining"))
    if remaining is None:
        raise HTTPException(status_code=402, detail="Token record invalid.")

    if remaining == -1:
        return {"token": token, "remaining": -1, "_row": rec["_row"], "plan": rec.get("plan", "")}

    if remaining <= 0:
        raise HTTPException(status_code=402, detail="Out of credits.")

    return {"token": token, "remaining": remaining, "_row": rec["_row"], "plan": rec.get("plan", "")}


def _burn_one_credit(ws, rec: Dict[str, Any]) -> Dict[str, Any]:
    if rec["remaining"] == -1:
        ws.update_cell(rec["_row"], 5, _utc_now_iso())
        return {"remaining": -1}

    new_remaining = rec["remaining"] - 1
    ws.update_cell(rec["_row"], 3, str(new_remaining))
    ws.update_cell(rec["_row"], 5, _utc_now_iso())
    return {"remaining": new_remaining}


# =========================
# ROUTES
# =========================
@app.get("/api/health")
def health():
    return {
        "ok": True,
        "service": "webtopdf",
        "time": _utc_now_iso(),
        "tokens_tab": TOKENS_TAB,
    }


@app.post("/api/admin/create-token")
def admin_create_token(request: Request, payload: dict = Body(...)):
    provided_key = (request.headers.get("x-admin-key") or "").strip()
    if not ADMIN_KEY or provided_key != ADMIN_KEY:
        raise HTTPException(status_code=401, detail="Unauthorized")

    plan = (payload.get("plan") or "10pack").strip()
    remaining = int(payload.get("remaining", 10))
    note = (payload.get("note") or "").strip()

    ws = _get_tokens_ws()
    token = _mint_token()

    ws.append_row([
        token,
        plan,
        str(remaining),
        _utc_now_iso(),
        "",
        note
    ], value_input_option="RAW")

    return {"ok": True, "token": token, "remaining": remaining}


@app.post("/api/web-to-pdf")
def web_to_pdf(payload: dict = Body(...)):
    url = (payload.get("url") or "").strip()
    token = (payload.get("token") or "").strip()

    if not _is_probably_url(url):
        raise HTTPException(status_code=400, detail="URL must start with http:// or https://")

    ws = _get_tokens_ws()

    # Gate ONLY
    allow = _allow_only(ws, token)

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"],
            )
            context = browser.new_context()
            page = context.new_page()

            page.goto(url, wait_until="networkidle", timeout=60000)
            pdf_bytes = page.pdf(
                format="Letter",
                print_background=True,
                margin={"top": "0.5in", "bottom": "0.5in", "left": "0.5in", "right": "0.5in"},
            )

            context.close()
            browser.close()

    except PWTimeoutError:
        raise HTTPException(status_code=408, detail="Page load timed out.")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF failed: {str(e)}")

    # SUCCESS → burn credit
    credit = _burn_one_credit(ws, allow)

    job_id = uuid.uuid4().hex[:10].upper()
    filename = f"webtopdf_{job_id}.pdf"

    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "X-Credits-Remaining": str(credit.get("remaining")),
        },
    )


# =========================
# STATIC UI (optional)
# =========================
if os.path.isdir("static"):
    app.mount("/", StaticFiles(directory="static", html=True), name="ui")

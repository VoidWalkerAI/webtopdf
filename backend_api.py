import os
import re
import uuid
import json
import datetime as dt
from typing import Optional, Dict, Any

import gspread
from google.oauth2.service_account import Credentials

from fastapi import FastAPI, Body, HTTPException, Request
from fastapi.responses import Response, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError


app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# -----------------------
# Config
# -----------------------
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


# -----------------------
# Helpers
# -----------------------
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
        raise HTTPException(status_code=500, detail="GOOGLE_SERVICE_ACCOUNT_JSON is not valid JSON text.")


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

    # Ensure headers
    existing = ws.row_values(1)
    if existing != TOKENS_HEADERS:
        ws.update("A1", [TOKENS_HEADERS])

    return ws


def _mint_token() -> str:
    # short but strong enough for MVP
    return uuid.uuid4().hex[:16].upper()


def _get_token_record(ws, token: str) -> Optional[Dict[str, Any]]:
    token = (token or "").strip()
    if not token:
        return None

    # Find token in column A
    try:
        cell = ws.find(token)
    except Exception:
        return None

    # Must be in column 1 and not header row
    if not cell or cell.row <= 1 or cell.col != 1:
        return None

    row = ws.row_values(cell.row)
    # Normalize to headers
    rec = {TOKENS_HEADERS[i]: (row[i] if i < len(row) else "") for i in range(len(TOKENS_HEADERS))}
    rec["_row"] = cell.row
    return rec


def _parse_remaining(s: str) -> Optional[int]:
    s = (s or "").strip()
    if s == "":
        return None
    try:
        return int(s)
    except Exception:
        return None


def _allow_only(ws, token: str) -> Dict[str, Any]:
    rec = _get_token_record(ws, token)
    if not rec:
        raise HTTPException(status_code=402, detail="Missing/invalid token. Buy credits to continue.")

    remaining_i = _parse_remaining(rec.get("remaining", ""))
    if remaining_i is None:
        raise HTTPException(status_code=402, detail="Token record invalid (remaining not a number).")

    # Unlimited
    if remaining_i == -1:
        return {"token": token, "plan": rec.get("plan", ""), "remaining": -1, "_row": rec["_row"]}

    if remaining_i <= 0:
        raise HTTPException(status_code=402, detail="Out of credits. Buy more to continue.")

    return {"token": token, "plan": rec.get("plan", ""), "remaining": remaining_i, "_row": rec["_row"]}

def _burn_one_credit(ws, rec: Dict[str, Any]) -> Dict[str, Any]:
    # rec must contain: _row, remaining, plan, token
    if rec["remaining"] == -1:
        ws.update_cell(rec["_row"], 5, _utc_now_iso())  # last_used_at_utc
        return {"token": rec["token"], "plan": rec.get("plan", ""), "remaining": -1}

    new_remaining = int(rec["remaining"]) - 1
    ws.update_cell(rec["_row"], 3, str(new_remaining))  # remaining
    ws.update_cell(rec["_row"], 5, _utc_now_iso())      # last_used_at_utc
    return {"token": rec["token"], "plan": rec.get("plan", ""), "remaining": new_remaining}
    

# -----------------------
# Routes
# -----------------------
@app.get("/api/health")
def health():
    return {
        "ok": True,
        "service": "webtopdf",
        "time": _utc_now_iso(),
        "sheets_configured": bool(SHEET_ID and SERVICE_JSON),
        "tokens_tab": TOKENS_TAB,
    }


@app.get("/api/credits")
def credits(token: str):
    ws = _get_tokens_ws()
    rec = _get_token_record(ws, token)
    if not rec:
        return {"ok": False, "token": token, "valid": False}
    remaining_i = _parse_remaining(rec.get("remaining", ""))
    return {
        "ok": True,
        "valid": True,
        "token": token,
        "plan": rec.get("plan", ""),
        "remaining": remaining_i,
        "last_used_at_utc": rec.get("last_used_at_utc", ""),
    }


@app.post("/api/admin/create-token")
def admin_create_token(payload: dict = Body(...)):
    # Protect with ADMIN_KEY header: X-Admin-Key
    # This lets YOU mint tokens without Stripe yet.
    # Later we can swap this for Stripe webhook automation.
    requested_plan = (payload.get("plan") or "10pack").strip()
    requested_remaining = payload.get("remaining", 10)
    note = (payload.get("note") or "").strip()

    # Simple auth
    # (We check header in request middleware style below — easiest is using Request, but keeping simple:)
    # We'll accept admin_key in payload too as fallback.
    provided_key = (payload.get("admin_key") or "").strip()
    if ADMIN_KEY and provided_key != ADMIN_KEY:
        raise HTTPException(status_code=401, detail="Unauthorized (bad admin_key).")
    if not ADMIN_KEY:
        raise HTTPException(status_code=500, detail="ADMIN_KEY not set on server.")

    try:
        remaining_i = int(requested_remaining)
    except Exception:
        raise HTTPException(status_code=400, detail="remaining must be an integer (use -1 for unlimited).")

    ws = _get_tokens_ws()
    token = _mint_token()

    row = [
        token,
        requested_plan,
        str(remaining_i),
        _utc_now_iso(),
        "",  # last_used_at_utc
        note,
    ]
    ws.append_row(row, value_input_option="RAW")

    return {"ok": True, "token": token, "plan": requested_plan, "remaining": remaining_i}


@app.post("/api/web-to-pdf")
def web_to_pdf(payload: dict = Body(...)):
    url = (payload.get("url") or "").strip()
    token = (payload.get("token") or "").strip()

    if not _is_probably_url(url):
        raise HTTPException(status_code=400, detail="URL must start with http:// or https://")
    if len(url) > 2000:
        raise HTTPException(status_code=400, detail="URL too long.")

    ws = _get_tokens_ws()

    # 1) Gate only (NO decrement yet)
    allow = _allow_only(ws, token)

    job_id = uuid.uuid4().hex[:10].upper()

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
        raise HTTPException(status_code=408, detail="Timed out loading page. Try again or use a simpler URL.")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF generation failed: {str(e)}")

    # 2) SUCCESS: burn one credit now
    credit_info = _burn_one_credit(ws, allow)

    filename = f"webtopdf_{job_id}.pdf"
    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "X-Job-Id": job_id,
            "X-Credits-Remaining": str(credit_info.get("remaining")),
        },
    )


# UI last (serve static/index.html at /) — only if folder exists
if os.path.isdir("static"):
    app.mount("/", StaticFiles(directory="static", html=True), name="ui")

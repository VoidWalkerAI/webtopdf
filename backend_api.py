# backend_api.py
# WebToPDF – single-file backend (FastAPI) with:
# - /api/web-to-pdf (Playwright PDF)
# - token/credits gate + burn-on-success
# - Stripe Checkout session creation + webhook token mint
# - Stripe log sheet + Jobs log sheet
#
# ENV VARS REQUIRED:
#   VOYDS_FORMS_SHEET_ID
#   GOOGLE_SERVICE_ACCOUNT_JSON
#
# OPTIONAL TAB NAMES (defaults shown):
#   VOYDS_TOKENS_TAB          = WebToPDF_Tokens
#   VOYDS_STRIPE_LOG_TAB      = WebToPDF_StripeLog
#   VOYDS_JOBS_TAB            = WebToPDF_Jobs
#
# ADMIN:
#   ADMIN_KEY
#
# STRIPE:
#   STRIPE_SECRET_KEY
#   STRIPE_WEBHOOK_SECRET
#   STRIPE_SUCCESS_URL
#   STRIPE_CANCEL_URL
#   STRIPE_PRICE_ECONOMY      (10 credits)
#   STRIPE_PRICE_PRO          (50 credits)
#   STRIPE_PRICE_PLATINUM     (200 credits)

import os
import json
import uuid
import secrets
from datetime import datetime, timezone
from typing import Dict, Any, Optional, List

from fastapi import FastAPI, Body, HTTPException, Request, Response
from fastapi.middleware.cors import CORSMiddleware

import stripe
import gspread
from google.oauth2.service_account import Credentials

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError


# -----------------------
# App
# -----------------------
app = FastAPI(title="WebToPDF API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # tighten later if you want
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# -----------------------
# ENV
# -----------------------
SHEET_ID = (os.getenv("VOYDS_FORMS_SHEET_ID") or "").strip()
SERVICE_JSON = (os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON") or "").strip()

TOKENS_TAB = (os.getenv("VOYDS_TOKENS_TAB") or "WebToPDF_Tokens").strip()
STRIPE_LOG_TAB = (os.getenv("VOYDS_STRIPE_LOG_TAB") or "WebToPDF_StripeLog").strip()
JOBS_TAB = (os.getenv("VOYDS_JOBS_TAB") or "WebToPDF_Jobs").strip()

ADMIN_KEY = (os.getenv("ADMIN_KEY") or "").strip()

STRIPE_SECRET_KEY = (os.getenv("STRIPE_SECRET_KEY") or "").strip()
STRIPE_WEBHOOK_SECRET = (os.getenv("STRIPE_WEBHOOK_SECRET") or "").strip()
SUCCESS_URL = (os.getenv("STRIPE_SUCCESS_URL") or "").strip()
CANCEL_URL = (os.getenv("STRIPE_CANCEL_URL") or "").strip()

PRICE_ECONOMY = (os.getenv("STRIPE_PRICE_ECONOMY") or "").strip()     # 10 credits
PRICE_PRO = (os.getenv("STRIPE_PRICE_PRO") or "").strip()             # 50 credits
PRICE_PLATINUM = (os.getenv("STRIPE_PRICE_PLATINUM") or "").strip()   # 200 credits

if STRIPE_SECRET_KEY:
    stripe.api_key = STRIPE_SECRET_KEY


# -----------------------
# Headers (Google Sheets)
# -----------------------
TOKENS_HEADERS = ["token", "plan", "remaining", "created_at_utc", "last_used_at_utc", "note"]

STRIPE_LOG_HEADERS = ["session_id", "paid_at_utc", "tier", "credits", "token", "customer_email"]

JOBS_HEADERS = [
    "job_id",
    "created_at_utc",
    "finished_at_utc",
    "url",
    "status",
    "pdf_filename",
    "error",
    "token",
    "plan",
    "credits_before",
    "credits_after",
]


# -----------------------
# Time / Utils
# -----------------------
def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _is_probably_url(u: str) -> bool:
    u = (u or "").strip().lower()
    return u.startswith("http://") or u.startswith("https://")


def _mint_token(prefix: str = "") -> str:
    core = secrets.token_hex(8).upper()  # 16 chars
    return f"{prefix}{core}" if prefix else core


def _parse_int_strict(s: Any) -> Optional[int]:
    s = str(s).strip()
    if s == "":
        return None
    try:
        return int(s)
    except Exception:
        return None


def _stripe_prices_ok() -> bool:
    return bool(
        STRIPE_SECRET_KEY
        and STRIPE_WEBHOOK_SECRET
        and PRICE_ECONOMY
        and PRICE_PRO
        and PRICE_PLATINUM
        and SUCCESS_URL
        and CANCEL_URL
    )


# -----------------------
# Google Sheets Client (cached)
# -----------------------
_gc_client: Optional[gspread.Client] = None
_sheet_cache: Optional[gspread.Spreadsheet] = None


def _get_gspread_client() -> gspread.Client:
    global _gc_client
    if _gc_client:
        return _gc_client

    if not SERVICE_JSON:
        raise HTTPException(status_code=500, detail="GOOGLE_SERVICE_ACCOUNT_JSON missing.")
    try:
        info = json.loads(SERVICE_JSON)
    except Exception:
        raise HTTPException(status_code=500, detail="GOOGLE_SERVICE_ACCOUNT_JSON is not valid JSON.")

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    _gc_client = gspread.authorize(creds)
    return _gc_client


def _get_sheet() -> gspread.Spreadsheet:
    global _sheet_cache
    if _sheet_cache:
        return _sheet_cache
    if not SHEET_ID:
        raise HTTPException(status_code=500, detail="VOYDS_FORMS_SHEET_ID missing.")
    sh = _get_gspread_client().open_by_key(SHEET_ID)
    _sheet_cache = sh
    return sh


def _ws(tab_name: str, headers: Optional[List[str]] = None) -> gspread.Worksheet:
    sh = _get_sheet()

    try:
        ws = sh.worksheet(tab_name)
    except Exception:
        cols = max(1, len(headers or []))
        ws = sh.add_worksheet(title=tab_name, rows=2000, cols=cols)

    if headers:
        existing = ws.row_values(1)
        if existing != headers:
            ws.update("A1", [headers])

    return ws


def _get_tokens_ws() -> gspread.Worksheet:
    return _ws(TOKENS_TAB, TOKENS_HEADERS)


def _get_stripe_log_ws() -> gspread.Worksheet:
    return _ws(STRIPE_LOG_TAB, STRIPE_LOG_HEADERS)


def _get_jobs_ws() -> gspread.Worksheet:
    return _ws(JOBS_TAB, JOBS_HEADERS)


# -----------------------
# Tokens logic
# -----------------------
def _get_all_records(ws: gspread.Worksheet) -> List[Dict[str, Any]]:
    try:
        return ws.get_all_records()
    except Exception:
        return []


def _find_row_index_by_token(ws: gspread.Worksheet, token: str) -> Optional[int]:
    token = (token or "").strip()
    if not token:
        return None
    try:
        cell = ws.find(token)
    except Exception:
        return None

    # must be col A, not header
    if not cell or cell.col != 1 or cell.row <= 1:
        return None
    return cell.row


def _get_token_record(ws: gspread.Worksheet, token: str) -> Optional[Dict[str, Any]]:
    token = (token or "").strip()
    if not token:
        return None
    rows = _get_all_records(ws)
    for r in rows:
        if str(r.get("token", "")).strip() == token:
            return r
    return None


def _allow_only(ws: gspread.Worksheet, token: str) -> Dict[str, Any]:
    token = (token or "").strip()
    if not token:
        raise HTTPException(status_code=401, detail="Missing token.")

    rec = _get_token_record(ws, token)
    if not rec:
        raise HTTPException(status_code=401, detail="Invalid token.")

    remaining_i = _parse_int_strict(rec.get("remaining", ""))
    if remaining_i is None:
        raise HTTPException(status_code=402, detail="Token record invalid (remaining not a number).")

    row_index = _find_row_index_by_token(ws, token)
    if not row_index:
        raise HTTPException(status_code=500, detail="Token row not found for update.")

    # Unlimited
    if remaining_i == -1:
        return {"token": token, "row_index": row_index, "plan": rec.get("plan", ""), "remaining_before": -1}

    if remaining_i <= 0:
        raise HTTPException(status_code=402, detail="No credits remaining.")

    return {"token": token, "row_index": row_index, "plan": rec.get("plan", ""), "remaining_before": remaining_i}


def _burn_one_credit(ws: gspread.Worksheet, allow: Dict[str, Any]) -> Dict[str, Any]:
    row = int(allow["row_index"])
    remaining_before = int(allow["remaining_before"])
    token = allow["token"]

    # unlimited
    if remaining_before == -1:
        ws.update_cell(row, 5, _utc_now_iso())  # last_used_at_utc
        return {"token": token, "remaining_before": -1, "remaining_after": -1}

    remaining_after = max(0, remaining_before - 1)

    ws.update_cell(row, 3, str(remaining_after))  # remaining (col C)
    ws.update_cell(row, 5, _utc_now_iso())        # last_used_at_utc (col E)
    return {"token": token, "remaining_before": remaining_before, "remaining_after": remaining_after}


def _create_token_in_sheet(credits: int, plan_name: str, note: str, email: str = "") -> str:
    ws = _get_tokens_ws()
    token = _mint_token()
    note_full = note
    if email:
        note_full = f"{note} | email:{email}".strip()

    row = [token, plan_name, str(int(credits)), _utc_now_iso(), "", note_full]
    ws.append_row(row, value_input_option="RAW")
    return token


# -----------------------
# Jobs log logic
# -----------------------
def _jobs_find_row(ws: gspread.Worksheet, job_id: str) -> Optional[int]:
    if not job_id:
        return None
    try:
        cell = ws.find(job_id)
    except Exception:
        return None
    if not cell or cell.col != 1 or cell.row <= 1:
        return None
    return cell.row


def _jobs_append_start(ws: gspread.Worksheet, job_id: str, url: str, token: str, plan: str, credits_before: int) -> None:
    row = [
        job_id,
        _utc_now_iso(),   # created_at_utc
        "",               # finished_at_utc
        url,
        "started",
        "",               # pdf_filename
        "",               # error
        token,
        plan,
        str(credits_before),
        "",               # credits_after
    ]
    ws.append_row(row, value_input_option="RAW")


def _jobs_update_finish(
    ws: gspread.Worksheet,
    job_id: str,
    status: str,
    pdf_filename: str = "",
    error: str = "",
    credits_after: Optional[int] = None,
) -> None:
    row = _jobs_find_row(ws, job_id)
    if not row:
        return

    ws.update_cell(row, 3, _utc_now_iso())  # finished_at_utc
    ws.update_cell(row, 5, status)          # status
    ws.update_cell(row, 6, pdf_filename or "")
    ws.update_cell(row, 7, error or "")
    if credits_after is not None:
        ws.update_cell(row, 11, str(credits_after))


# -----------------------
# Stripe log helpers
# -----------------------
def _log_has_session(ws: gspread.Worksheet, session_id: str) -> bool:
    if not session_id:
        return False
    try:
        cell = ws.find(session_id)
        return bool(cell and cell.row > 1 and cell.col == 1)
    except Exception:
        return False


def _log_get_by_session(ws: gspread.Worksheet, session_id: str) -> Optional[Dict[str, Any]]:
    recs = _get_all_records(ws)
    for r in recs:
        if str(r.get("session_id", "")).strip() == session_id:
            return r
    return None


# -----------------------
# Pricing tier mapping (accepts both labels and numbers)
# -----------------------
def _tier_to_price_and_credits(tier: str) -> Dict[str, Any]:
    t = (tier or "").strip().lower()

    if t in ("economy", "eco", "10", "10pack", "10credits", "10_credits"):
        return {"tier": "Economy", "price_id": PRICE_ECONOMY, "credits": 10}

    if t in ("pro", "professional", "50", "50pack", "50credits", "50_credits"):
        return {"tier": "Pro", "price_id": PRICE_PRO, "credits": 50}

    if t in ("platinum", "plat", "200", "200pack", "200credits", "200_credits"):
        return {"tier": "Platinum", "price_id": PRICE_PLATINUM, "credits": 200}

    raise HTTPException(status_code=400, detail="Bad tier. Use economy | pro | platinum (or 10 | 50 | 200).")


# -----------------------
# API
# -----------------------
@app.get("/api/health")
def health():
    return {
        "ok": True,
        "service": "webtopdf",
        "time": _utc_now_iso(),
        "sheets_configured": bool(SHEET_ID and SERVICE_JSON),
        "tokens_tab": TOKENS_TAB,
        "stripe_log_tab": STRIPE_LOG_TAB,
        "jobs_tab": JOBS_TAB,
        "stripe_ready": _stripe_prices_ok(),
    }


@app.get("/api/credits")
def credits(token: str):
    ws = _get_tokens_ws()
    rec = _get_token_record(ws, token)
    if not rec:
        return {"ok": False, "token": token, "valid": False}
    remaining_i = _parse_int_strict(rec.get("remaining", ""))
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
    requested_plan = (payload.get("plan") or "manual").strip()
    requested_remaining = payload.get("remaining", 10)
    note = (payload.get("note") or "").strip()

    provided_key = (payload.get("admin_key") or "").strip()
    if ADMIN_KEY and provided_key != ADMIN_KEY:
        raise HTTPException(status_code=401, detail="Unauthorized (bad admin_key).")
    if not ADMIN_KEY:
        raise HTTPException(status_code=500, detail="ADMIN_KEY not set on server.")

    remaining_i = _parse_int_strict(requested_remaining)
    if remaining_i is None:
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

    tokens_ws = _get_tokens_ws()
    jobs_ws = _get_jobs_ws()

    # Gate only (no decrement until success)
    allow = _allow_only(tokens_ws, token)

    job_id = uuid.uuid4().hex[:10].upper()
    filename = f"webtopdf_{job_id}.pdf"

    # Log START
    _jobs_append_start(
        jobs_ws,
        job_id=job_id,
        url=url,
        token=allow["token"],
        plan=allow.get("plan", ""),
        credits_before=int(allow.get("remaining_before", 0)),
    )

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"],
            )
            context = browser.new_context()
            page = context.new_page()
            try:
                page.goto(url, wait_until="networkidle", timeout=60000)
                pdf_bytes = page.pdf(
                    format="Letter",
                    print_background=True,
                    margin={"top": "0.5in", "bottom": "0.5in", "left": "0.5in", "right": "0.5in"},
                )
            finally:
                context.close()
                browser.close()

    except PWTimeoutError:
        _jobs_update_finish(jobs_ws, job_id=job_id, status="timeout", error="Timed out loading page.")
        raise HTTPException(status_code=408, detail="Timed out loading page. Try again or use a simpler URL.")
    except Exception as e:
        _jobs_update_finish(jobs_ws, job_id=job_id, status="error", error=f"PDF generation failed: {str(e)}")
        raise HTTPException(status_code=500, detail=f"PDF generation failed: {str(e)}")

    # Burn credit ONLY on success
    credit_info = _burn_one_credit(tokens_ws, allow)

    # Log SUCCESS
    _jobs_update_finish(
        jobs_ws,
        job_id=job_id,
        status="success",
        pdf_filename=filename,
        error="",
        credits_after=credit_info.get("remaining_after"),
    )

    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "X-Job-Id": job_id,
            "X-Credits-Remaining": str(credit_info.get("remaining_after")),
        },
    )


# -----------------------
# Stripe Checkout: create session
# -----------------------
@app.post("/api/stripe/create-checkout-session")
def stripe_create_checkout_session(payload: dict = Body(...)):
    if not _stripe_prices_ok():
        raise HTTPException(status_code=500, detail="Stripe not configured. Missing keys/price IDs/success/cancel URLs.")

    tier = (payload.get("tier") or "").strip()
    info = _tier_to_price_and_credits(tier)

    try:
        session = stripe.checkout.Session.create(
            mode="payment",
            line_items=[{"price": info["price_id"], "quantity": 1}],
            success_url=SUCCESS_URL,
            cancel_url=CANCEL_URL,
            metadata={
                "tier": info["tier"],
                "credits": str(info["credits"]),
                "product": "WebToPDF",
            },
        )
        return {"ok": True, "url": session.url}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Stripe session create failed: {str(e)}")


# -----------------------
# Stripe: webhook (mints token + logs session)
# -----------------------
@app.post("/api/stripe/webhook")
async def stripe_webhook(request: Request):
    if not STRIPE_WEBHOOK_SECRET:
        raise HTTPException(status_code=500, detail="STRIPE_WEBHOOK_SECRET not set on server.")

    payload_bytes = await request.body()
    sig = request.headers.get("stripe-signature", "")

    try:
        event = stripe.Webhook.construct_event(
            payload=payload_bytes,
            sig_header=sig,
            secret=STRIPE_WEBHOOK_SECRET,
        )
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid Stripe signature.")

    # Only care about successful payments
    if event.get("type") != "checkout.session.completed":
        return {"ok": True, "ignored": True, "type": event.get("type")}

    session = event["data"]["object"]
    session_id = session.get("id", "")

    customer_email = (session.get("customer_details", {}) or {}).get("email", "") or ""

    meta = session.get("metadata", {}) or {}
    tier = meta.get("tier", "Unknown")
    credits_s = meta.get("credits", "0")
    credits_i = _parse_int_strict(credits_s) or 0

    log_ws = _get_stripe_log_ws()

    # Dedupe
    if _log_has_session(log_ws, session_id):
        return {"ok": True, "deduped": True, "session_id": session_id}

    # Mint token (credits)
    token = _create_token_in_sheet(
        credits=credits_i,
        plan_name=f"Stripe {tier}",
        note=f"stripe_session:{session_id}",
        email=customer_email,
    )

    # Log session
    log_ws.append_row(
        [session_id, _utc_now_iso(), tier, str(credits_i), token, customer_email],
        value_input_option="RAW",
    )

    return {"ok": True, "session_id": session_id, "token_minted": True}


# -----------------------
# Stripe: success page lookup
# -----------------------
@app.get("/api/stripe/session-result")
def stripe_session_result(session_id: str):
    log_ws = _get_stripe_log_ws()
    rec = _log_get_by_session(log_ws, session_id)
    if not rec:
        return {"ok": False, "ready": False}
    return {
        "ok": True,
        "ready": True,
        "tier": rec.get("tier", ""),
        "credits": rec.get("credits", ""),
        "token": rec.get("token", ""),
    }

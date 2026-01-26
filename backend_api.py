# backend_api.py
# WebToPDF – single-file backend (FastAPI) with:
# - /api/web-to-pdf (Playwright PDF)
# - token/credits gate + burn-on-success
# - Stripe Checkout session creation + webhook mint
# - Jobs log sheet (start + finish + error)
#
# ENV VARS REQUIRED:
#   VOYDS_FORMS_SHEET_ID
#   GOOGLE_SERVICE_ACCOUNT_JSON
#   VOYDS_TOKENS_TAB            (default: WebToPDF_Tokens)
#   VOYDS_STRIPE_LOG_TAB        (default: WebToPDF_StripeLog)
#   VOYDS_JOBS_TAB              (default: WebToPDF_Jobs)   <-- add this
#   ADMIN_KEY
#
# STRIPE:
#   STRIPE_SECRET_KEY
#   STRIPE_WEBHOOK_SECRET
#   STRIPE_SUCCESS_URL
#   STRIPE_CANCEL_URL
#   STRIPE_PRICE_ECONOMY
#   STRIPE_PRICE_PRO
#   STRIPE_PRICE_PLATINUM

import os
import json
import uuid
import secrets
from datetime import datetime, timezone
from typing import Dict, Any, Optional, Tuple

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
    allow_origins=["*"],   # tighten later if you want
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

PRICE_ECONOMY = (os.getenv("STRIPE_PRICE_ECONOMY") or "").strip()      # 10 credits
PRICE_PRO = (os.getenv("STRIPE_PRICE_PRO") or "").strip()              # 50 credits
PRICE_PLATINUM = (os.getenv("STRIPE_PRICE_PLATINUM") or "").strip()    # 200 credits

if STRIPE_SECRET_KEY:
    stripe.api_key = STRIPE_SECRET_KEY


# -----------------------
# Time / Utils
# -----------------------
def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _is_probably_url(u: str) -> bool:
    u = (u or "").strip().lower()
    return u.startswith("http://") or u.startswith("https://")


def _mint_token(prefix: str = "") -> str:
    # stable, short-ish token; you can change prefix later if you want
    core = secrets.token_hex(8).upper()
    return f"{prefix}{core}" if prefix else core


def _parse_remaining(s: str) -> int:
    try:
        return int(str(s).strip())
    except Exception:
        return 0


def _stripe_prices_ok() -> bool:
    return bool(
        STRIPE_SECRET_KEY
        and PRICE_ECONOMY
        and PRICE_PRO
        and PRICE_PLATINUM
        and SUCCESS_URL
        and CANCEL_URL
    )


# -----------------------
# Google Sheets Client
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


def _ws(tab_name: str) -> gspread.Worksheet:
    sh = _get_sheet()
    try:
        return sh.worksheet(tab_name)
    except Exception:
        raise HTTPException(status_code=500, detail=f"Missing worksheet tab: {tab_name}")


def _get_tokens_ws() -> gspread.Worksheet:
    return _ws(TOKENS_TAB)


def _get_stripe_log_ws() -> gspread.Worksheet:
    return _ws(STRIPE_LOG_TAB)


def _get_jobs_ws() -> gspread.Worksheet:
    return _ws(JOBS_TAB)


# -----------------------
# Tokens Table Logic
# Expected headers:
# token | plan | remaining | created_at_utc | last_used_at_utc | note
# -----------------------
def _get_all_records(ws: gspread.Worksheet) -> list[dict]:
    # gspread get_all_records assumes header row exists
    try:
        return ws.get_all_records()
    except Exception:
        return []


def _find_row_index_by_token(ws: gspread.Worksheet, token: str) -> Optional[int]:
    # token is in col A (1)
    if not token:
        return None
    try:
        cell = ws.find(token)
        return cell.row if cell else None
    except Exception:
        return None


def _get_token_record(ws: gspread.Worksheet, token: str) -> Optional[dict]:
    token = (token or "").strip()
    if not token:
        return None
    rows = _get_all_records(ws)
    for r in rows:
        if str(r.get("token", "")).strip() == token:
            return r
    return None


def _allow_only(ws: gspread.Worksheet, token: str) -> dict:
    """
    Gate only. Returns a struct we can later burn.
    - remaining = -1 means unlimited
    """
    token = (token or "").strip()
    if not token:
        raise HTTPException(status_code=401, detail="Missing token.")

    rec = _get_token_record(ws, token)
    if not rec:
        raise HTTPException(status_code=401, detail="Invalid token.")

    remaining_i = _parse_remaining(rec.get("remaining", "0"))
    if remaining_i == 0:
        raise HTTPException(status_code=402, detail="No credits remaining.")

    row_index = _find_row_index_by_token(ws, token)
    if not row_index:
        raise HTTPException(status_code=500, detail="Token row not found for update.")

    return {
        "token": token,
        "row_index": row_index,
        "plan": rec.get("plan", ""),
        "remaining_before": remaining_i,
    }


def _burn_one_credit(ws: gspread.Worksheet, allow: dict) -> dict:
    """
    Decrement remaining by 1 (unless unlimited), and set last_used_at_utc.
    Returns remaining_after.
    """
    row = int(allow["row_index"])
    remaining_before = int(allow["remaining_before"])
    token = allow["token"]

    if remaining_before == -1:
        # unlimited
        ws.update_cell(row, 5, _utc_now_iso())  # last_used_at_utc
        return {"token": token, "remaining_before": -1, "remaining_after": -1}

    remaining_after = max(0, remaining_before - 1)

    # Column mapping:
    # A token (1)
    # B plan (2)
    # C remaining (3)
    # D created_at_utc (4)
    # E last_used_at_utc (5)
    # F note (6)
    ws.update_cell(row, 3, str(remaining_after))
    ws.update_cell(row, 5, _utc_now_iso())
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
# Jobs Log Logic
# Expected headers (what you already created):
# job_id | created_at_utc | finished_at_utc | url | status | pdf_filename | error | token | plan | credits_before | credits_after
# -----------------------
def _jobs_find_row(ws: gspread.Worksheet, job_id: str) -> Optional[int]:
    try:
        cell = ws.find(job_id)
        return cell.row if cell else None
    except Exception:
        return None


def _jobs_append_start(ws: gspread.Worksheet, job_id: str, url: str, token: str, plan: str, credits_before: int) -> int:
    # Write a "started" row immediately
    row = [
        job_id,
        _utc_now_iso(),     # created_at_utc
        "",                 # finished_at_utc
        url,
        "started",
        "",                 # pdf_filename
        "",                 # error
        token,
        plan,
        str(credits_before),
        "",                 # credits_after
    ]
    ws.append_row(row, value_input_option="RAW")
    # return row index
    idx = _jobs_find_row(ws, job_id)
    return idx or 0


def _jobs_update_finish(
    ws: gspread.Worksheet,
    job_id: str,
    status: str,
    pdf_filename: str,
    error: str,
    credits_after: Optional[int],
) -> None:
    row = _jobs_find_row(ws, job_id)
    if not row:
        # If we cannot find it, don't crash the request; best-effort logging.
        return

    finished = _utc_now_iso()
    # Columns:
    # A job_id (1)
    # B created_at_utc (2)
    # C finished_at_utc (3)
    # D url (4)
    # E status (5)
    # F pdf_filename (6)
    # G error (7)
    # H token (8)
    # I plan (9)
    # J credits_before (10)
    # K credits_after (11)
    ws.update_cell(row, 3, finished)
    ws.update_cell(row, 5, status)
    ws.update_cell(row, 6, pdf_filename or "")
    ws.update_cell(row, 7, error or "")
    if credits_after is not None:
        ws.update_cell(row, 11, str(credits_after))


# -----------------------
# Pricing tier mapping (supports BOTH label names and numeric inputs)
# -----------------------
def _tier_to_price_and_credits(tier: str) -> Dict[str, Any]:
    t = (tier or "").strip().lower()

    # Economy / 10
    if t in ("economy", "eco", "10", "10pack", "10_credits", "10credits"):
        return {"tier": "Economy", "price_id": PRICE_ECONOMY, "credits": 10}

    # Pro / 50
    if t in ("pro", "professional", "50", "50pack", "50_credits", "50credits"):
        return {"tier": "Pro", "price_id": PRICE_PRO, "credits": 50}

    # Platinum / 200
    if t in ("platinum", "plat", "200", "200pack", "200_credits", "200credits"):
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
        "jobs_tab": JOBS_TAB,
        "stripe_ready": _stripe_prices_ok(),
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
    requested_plan = (payload.get("plan") or "manual").strip()
    requested_remaining = payload.get("remaining", 10)
    note = (payload.get("note") or "").strip()

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

    tokens_ws = _get_tokens_ws()
    jobs_ws = _get_jobs_ws()

    # Gate only (no decrement until success)
    allow = _allow_only(tokens_ws, token)

    job_id = uuid.uuid4().hex[:10].upper()
    filename = f"webtopdf_{job_id}.pdf"

    # Write job start immediately
    _jobs_append_start(
        jobs_ws,
        job_id=job_id,
        url=url,
        token=allow["token"],
        plan=allow.get("plan", ""),
        credits_before=int(allow.get("remaining_before", 0)),
    )

    pdf_bytes = b""
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
        _jobs_update_finish(
            jobs_ws,
            job_id=job_id,
            status="timeout",
            pdf_filename="",
            error="Timed out loading page.",
            credits_after=None,
        )
        raise HTTPException(status_code=408, detail="Timed out loading page. Try again or use a simpler URL.")
    except Exception as e:
        _jobs_update_finish(
            jobs_ws,
            job_id=job_id,
            status="error",
            pdf_filename="",
            error=f"PDF generation failed: {str(e)}",
            credits_after=None,
        )
        raise HTTPException(status_code=500, detail=f"PDF generation failed: {str(e)}")

    # Success: burn one credit now
    credit_info = _burn_one_credit(tokens_ws, allow)

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
        raise HTTPException(status_code=500, detail="Stripe not configured. Missing keys/price IDs/success URLs.")

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

    # We only care about successful payments
    if event["type"] != "checkout.session.completed":
        return {"ok": True, "ignored": True, "type": event["type"]}

    session = event["data"]["object"]
    session_id = session.get("id", "")
    customer_email = (session.get("customer_details", {}) or {}).get("email", "") or ""

    meta = session.get("metadata", {}) or {}
    tier = meta.get("tier", "Unknown")
    credits_s = meta.get("credits", "0")

    try:
        credits_i = int(credits_s)
    except Exception:
        credits_i = 0

    # Dedupe
    log_ws = _get_stripe_log_ws()
    existing = _log_has_session(log_ws, session_id)
    if existing:
        return {"ok": True, "deduped": True, "session_id": session_id}

    # Mint token in tokens sheet
    token = _create_token_in_sheet(
        credits=credits_i,
        plan_name=f"Stripe {tier}",
        note=f"stripe_session:{session_id}",
        email=customer_email,
    )

    # Log for success-page lookup
    log_ws.append_row(
        [session_id, _utc_now_iso(), tier, str(credits_i), token, customer_email],
        value_input_option="RAW",
    )

    return {"ok": True, "session_id": session_id, "token_minted": True}


def _log_has_session(ws: gspread.Worksheet, session_id: str) -> bool:
    if not session_id:
        return False
    try:
        ws.find(session_id)
        return True
    except Exception:
        return False


def _log_get_by_session(ws: gspread.Worksheet, session_id: str) -> Optional[dict]:
    # expected headers:
    # session_id | paid_at_utc | tier | credits | token | customer_email
    recs = _get_all_records(ws)
    for r in recs:
        if str(r.get("session_id", "")).strip() == session_id:
            return r
    return None


# -----------------------
# Stripe: success page lookup (front-end calls this)
# -----------------------
@app.get("/api/stripe/session-result")
def stripe_session_result(session_id: str):
    log_ws = _get_stripe_log_ws()
    rec = _log_get_by_session(log_ws, session_id)
    if not rec:
        # webhook might not have landed yet; front-end can retry
        return {"ok": False, "ready": False}
    return {
        "ok": True,
        "ready": True,
        "tier": rec.get("tier", ""),
        "credits": rec.get("credits", ""),
        "token": rec.get("token", ""),
    }

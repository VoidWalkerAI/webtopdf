import os
import re
import uuid
import json
import datetime as dt
from typing import Optional, Dict, Any

import gspread
from google.oauth2.service_account import Credentials

import stripe
from fastapi import FastAPI, Body, HTTPException, Request
from fastapi.responses import Response
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError


# -----------------------
# App
# -----------------------
app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# -----------------------
# Config (Sheets)
# -----------------------
SHEET_ID = os.getenv("VOYDS_FORMS_SHEET_ID", "").strip()
SERVICE_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()

TOKENS_TAB = os.getenv("VOYDS_TOKENS_TAB", "WebToPDF_Tokens").strip()
STRIPE_LOG_TAB = os.getenv("VOYDS_STRIPE_LOG_TAB", "WebToPDF_StripeLog").strip()

ADMIN_KEY = os.getenv("ADMIN_KEY", "").strip()

TOKENS_HEADERS = [
    "token",
    "plan",
    "remaining",
    "created_at_utc",
    "last_used_at_utc",
    "note",
]

STRIPE_LOG_HEADERS = [
    "session_id",
    "paid_at_utc",
    "tier",
    "credits",
    "token",
    "customer_email",
]

# -----------------------
# Config (Stripe)
# -----------------------
STRIPE_SECRET_KEY = os.getenv("STRIPE_SECRET_KEY", "").strip()
STRIPE_WEBHOOK_SECRET = os.getenv("STRIPE_WEBHOOK_SECRET", "").strip()

PRICE_ECONOMY = os.getenv("STRIPE_PRICE_ECONOMY", "").strip()      # 10 credits
PRICE_PRO = os.getenv("STRIPE_PRICE_PRO", "").strip()              # 50 credits
PRICE_PLATINUM = os.getenv("STRIPE_PRICE_PLATINUM", "").strip()    # 200 credits

SUCCESS_URL = os.getenv("STRIPE_SUCCESS_URL", "").strip()  # ex: https://webtopdf-1.onrender.com/?success=1&session_id={CHECKOUT_SESSION_ID}
CANCEL_URL = os.getenv("STRIPE_CANCEL_URL", "").strip()    # ex: https://webtopdf-1.onrender.com/?cancel=1

if STRIPE_SECRET_KEY:
    stripe.api_key = STRIPE_SECRET_KEY


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


def _get_ws(tab_name: str, headers: list, rows: int = 2000):
    gc = _get_gspread_client()
    sh = gc.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=tab_name, rows=rows, cols=len(headers) + 4)

    existing = ws.row_values(1)
    if existing != headers:
        ws.update("A1", [headers])
    return ws


def _get_tokens_ws():
    return _get_ws(TOKENS_TAB, TOKENS_HEADERS)


def _get_stripe_log_ws():
    return _get_ws(STRIPE_LOG_TAB, STRIPE_LOG_HEADERS)


def _mint_token() -> str:
    return uuid.uuid4().hex[:16].upper()


def _get_token_record(ws, token: str) -> Optional[Dict[str, Any]]:
    token = (token or "").strip()
    if not token:
        return None
    try:
        cell = ws.find(token)
    except Exception:
        return None
    if not cell or cell.row <= 1 or cell.col != 1:
        return None

    row = ws.row_values(cell.row)
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
    if rec["remaining"] == -1:
        ws.update_cell(rec["_row"], 5, _utc_now_iso())  # last_used_at_utc
        return {"token": rec["token"], "plan": rec.get("plan", ""), "remaining": -1}

    new_remaining = int(rec["remaining"]) - 1
    ws.update_cell(rec["_row"], 3, str(new_remaining))  # remaining
    ws.update_cell(rec["_row"], 5, _utc_now_iso())      # last_used_at_utc
    return {"token": rec["token"], "plan": rec.get("plan", ""), "remaining": new_remaining}


def _stripe_prices_ok() -> bool:
    return bool(STRIPE_SECRET_KEY and PRICE_ECONOMY and PRICE_PRO and PRICE_PLATINUM and SUCCESS_URL and CANCEL_URL)


def _tier_to_price_and_credits(tier: str) -> Dict[str, Any]:
    t = (tier or "").strip().lower()

    if t in ("economy", "eco", "10", "10pack"):
        if not PRICE_ECONOMY:
            raise HTTPException(status_code=500, detail="Missing STRIPE_PRICE_ECONOMY (PRICE_ECONOMY is blank).")
        return {"tier": "Economy", "price_id": PRICE_ECONOMY, "credits": 10}

    if t in ("pro", "professional", "50", "50pack"):
        if not PRICE_PRO:
            raise HTTPException(status_code=500, detail="Missing STRIPE_PRICE_PRO (PRICE_PRO is blank).")
        return {"tier": "Pro", "price_id": PRICE_PRO, "credits": 50}

    if t in ("platinum", "plat", "200", "200pack"):
        if not PRICE_PLATINUM:
            raise HTTPException(status_code=500, detail="Missing STRIPE_PRICE_PLATINUM (PRICE_PLATINUM is blank).")
        return {"tier": "Platinum", "price_id": PRICE_PLATINUM, "credits": 200}

    raise HTTPException(status_code=400, detail="Bad tier. Use economy | pro | platinum (or 10/50/200).")


def _log_has_session(ws, session_id: str) -> bool:
    session_id = (session_id or "").strip()
    if not session_id:
        return False
    try:
        cell = ws.find(session_id)
        return bool(cell and cell.row > 1 and cell.col == 1)
    except Exception:
        return False


def _log_get_by_session(ws, session_id: str) -> Optional[Dict[str, Any]]:
    session_id = (session_id or "").strip()
    if not session_id:
        return None
    try:
        cell = ws.find(session_id)
    except Exception:
        return None
    if not cell or cell.row <= 1 or cell.col != 1:
        return None
    row = ws.row_values(cell.row)
    rec = {STRIPE_LOG_HEADERS[i]: (row[i] if i < len(row) else "") for i in range(len(STRIPE_LOG_HEADERS))}
    return rec


def _create_token_in_sheet(credits: int, plan_name: str, note: str = "", email: str = "") -> str:
    ws = _get_tokens_ws()
    token = _mint_token()
    row = [
        token,
        plan_name,
        str(int(credits)),
        _utc_now_iso(),
        "",  # last_used_at_utc
        (note or "").strip() + (f" | email:{email}" if email else ""),
    ]
    ws.append_row(row, value_input_option="RAW")
    return token


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

    ws = _get_tokens_ws()

    # Gate only (no decrement until success)
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
        raise HTTPException(status_code=408, detail="Timed out loading page. Try again or use a simpler URL.")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF generation failed: {str(e)}")

    # Success: burn one credit now
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
        "tier": info["tier"],          # Economy / Pro / Platinum
        "credits": str(info["credits"]),  # 10 / 50 / 200
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
    if _log_has_session(log_ws, session_id):
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


# UI last (serve static/index.html at /)
if os.path.isdir("static"):
    app.mount("/", StaticFiles(directory="static", html=True), name="ui")

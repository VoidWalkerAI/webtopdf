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

import stripe


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

# Stripe
STRIPE_SECRET_KEY = os.getenv("STRIPE_SECRET_KEY", "").strip()
STRIPE_WEBHOOK_SECRET = os.getenv("STRIPE_WEBHOOK_SECRET", "").strip()
PUBLIC_BASE_URL = os.getenv("PUBLIC_BASE_URL", "").strip().rstrip("/")

# Price IDs (from Stripe)
STRIPE_PRICE_10 = os.getenv("STRIPE_PRICE_10", "").strip()
STRIPE_PRICE_50 = os.getenv("STRIPE_PRICE_50", "").strip()
STRIPE_PRICE_200 = os.getenv("STRIPE_PRICE_200", "").strip()

STRIPE_EVENTS_TAB = os.getenv("VOYDS_STRIPE_EVENTS_TAB", "WebToPDF_StripeEvents").strip()

TOKENS_HEADERS = [
    "token",
    "plan",
    "remaining",
    "created_at_utc",
    "last_used_at_utc",
    "note",
]

EVENTS_HEADERS = [
    "event_id",
    "created_at_utc",
    "session_id",
    "payment_intent",
    "status",
    "token",
    "plan",
    "credits_added",
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


def _require_stripe_env():
    if not STRIPE_SECRET_KEY:
        raise HTTPException(status_code=500, detail="Missing STRIPE_SECRET_KEY.")
    if not STRIPE_WEBHOOK_SECRET:
        raise HTTPException(status_code=500, detail="Missing STRIPE_WEBHOOK_SECRET.")
    if not PUBLIC_BASE_URL:
        raise HTTPException(status_code=500, detail="Missing PUBLIC_BASE_URL (e.g., https://webtopdf-1.onrender.com).")
    if not (STRIPE_PRICE_10 and STRIPE_PRICE_50 and STRIPE_PRICE_200):
        raise HTTPException(
            status_code=500,
            detail="Missing Stripe price IDs. Set STRIPE_PRICE_10 / STRIPE_PRICE_50 / STRIPE_PRICE_200.",
        )


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


def _get_events_ws():
    gc = _get_gspread_client()
    sh = gc.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet(STRIPE_EVENTS_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=STRIPE_EVENTS_TAB, rows=5000, cols=len(EVENTS_HEADERS) + 2)

    existing = ws.row_values(1)
    if existing != EVENTS_HEADERS:
        ws.update("A1", [EVENTS_HEADERS])

    return ws


def _mint_token() -> str:
    # you can change the prefix if you want
    return f"wtp_{uuid.uuid4().hex[:8]}"


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

    if remaining_i == -1:
        return {"token": token, "plan": rec.get("plan", ""), "remaining": -1, "_row": rec["_row"]}

    if remaining_i <= 0:
        raise HTTPException(status_code=402, detail="Out of credits. Buy more to continue.")

    return {"token": token, "plan": rec.get("plan", ""), "remaining": remaining_i, "_row": rec["_row"]}


def _burn_one_credit(ws, rec: Dict[str, Any]) -> Dict[str, Any]:
    if rec["remaining"] == -1:
        ws.update_cell(rec["_row"], 5, _utc_now_iso())
        return {"token": rec["token"], "plan": rec.get("plan", ""), "remaining": -1}

    new_remaining = int(rec["remaining"]) - 1
    ws.update_cell(rec["_row"], 3, str(new_remaining))
    ws.update_cell(rec["_row"], 5, _utc_now_iso())
    return {"token": rec["token"], "plan": rec.get("plan", ""), "remaining": new_remaining}


def _price_for_plan(plan: str) -> str:
    plan = (plan or "").strip().lower()
    if plan in ["10", "10pack", "pack10"]:
        return STRIPE_PRICE_10
    if plan in ["50", "50pack", "pack50"]:
        return STRIPE_PRICE_50
    if plan in ["200", "200pack", "pack200"]:
        return STRIPE_PRICE_200
    # default
    return STRIPE_PRICE_10


def _credits_for_plan(plan: str) -> int:
    plan = (plan or "").strip().lower()
    if plan in ["10", "10pack", "pack10"]:
        return 10
    if plan in ["50", "50pack", "pack50"]:
        return 50
    if plan in ["200", "200pack", "pack200"]:
        return 200
    return 10


def _events_has_event_id(ws_events, event_id: str) -> bool:
    try:
        cell = ws_events.find(event_id)
        return bool(cell and cell.row > 1)
    except Exception:
        return False


def _append_event(ws_events, row: Dict[str, Any]):
    ws_events.append_row(
        [
            row.get("event_id", ""),
            row.get("created_at_utc", _utc_now_iso()),
            row.get("session_id", ""),
            row.get("payment_intent", ""),
            row.get("status", ""),
            row.get("token", ""),
            row.get("plan", ""),
            str(row.get("credits_added", "")),
            row.get("note", ""),
        ],
        value_input_option="RAW",
    )


def _find_event_by_session_id(ws_events, session_id: str) -> Optional[Dict[str, Any]]:
    session_id = (session_id or "").strip()
    if not session_id:
        return None

    # Find session_id in column C
    try:
        cell = ws_events.find(session_id)
    except Exception:
        return None

    if not cell or cell.row <= 1:
        return None

    row = ws_events.row_values(cell.row)
    rec = {EVENTS_HEADERS[i]: (row[i] if i < len(row) else "") for i in range(len(EVENTS_HEADERS))}
    return rec


def _create_token_row(ws_tokens, token: str, plan: str, remaining_i: int, note: str = ""):
    ws_tokens.append_row(
        [
            token,
            plan,
            str(remaining_i),
            _utc_now_iso(),
            "",  # last_used_at_utc
            note,
        ],
        value_input_option="RAW",
    )


def _add_credits_to_token(ws_tokens, token: str, plan: str, credits: int) -> Dict[str, Any]:
    rec = _get_token_record(ws_tokens, token)
    if not rec:
        # token does not exist: create it
        _create_token_row(ws_tokens, token, plan, credits, note="created via stripe")
        return {"token": token, "plan": plan, "remaining": credits}

    remaining_i = _parse_remaining(rec.get("remaining", ""))
    if remaining_i is None:
        raise HTTPException(status_code=500, detail="Token record invalid (remaining not a number).")

    if remaining_i == -1:
        # unlimited stays unlimited
        ws_tokens.update_cell(rec["_row"], 5, _utc_now_iso())
        return {"token": token, "plan": rec.get("plan", plan), "remaining": -1}

    new_remaining = remaining_i + int(credits)
    ws_tokens.update_cell(rec["_row"], 2, plan)  # plan
    ws_tokens.update_cell(rec["_row"], 3, str(new_remaining))  # remaining
    ws_tokens.update_cell(rec["_row"], 5, _utc_now_iso())  # last_used_at_utc
    return {"token": token, "plan": plan, "remaining": new_remaining}


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
        "stripe_configured": bool(STRIPE_SECRET_KEY and STRIPE_WEBHOOK_SECRET and PUBLIC_BASE_URL),
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
    requested_plan = (payload.get("plan") or "10pack").strip()
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

    _create_token_row(ws, token, requested_plan, remaining_i, note=note)

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


# -----------------------
# Stripe: Checkout + Webhook
# -----------------------
@app.post("/api/stripe/create-checkout-session")
def stripe_create_checkout_session(payload: dict = Body(...)):
    """
    payload:
      plan: "10pack" | "50pack" | "200pack"
      token: optional; if empty, we mint a new one
    """
    _require_stripe_env()
    stripe.api_key = STRIPE_SECRET_KEY

    plan = (payload.get("plan") or "10pack").strip()
    token = (payload.get("token") or "").strip()

    if not token:
        token = _mint_token()
        token_mode = "new"
    else:
        token_mode = "existing"

    price_id = _price_for_plan(plan)

    # Stripe success returns session_id so UI can fetch the minted token/credits result
    success_url = f"{PUBLIC_BASE_URL}/?success=1&session_id={{CHECKOUT_SESSION_ID}}"
    cancel_url = f"{PUBLIC_BASE_URL}/?canceled=1"

    try:
        session = stripe.checkout.Session.create(
            mode="payment",
            line_items=[{"price": price_id, "quantity": 1}],
            success_url=success_url,
            cancel_url=cancel_url,
            metadata={
                "plan": plan,
                "token": token,
                "token_mode": token_mode,
            },
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Stripe session create failed: {str(e)}")

    return {"ok": True, "checkout_url": session.url, "token": token, "plan": plan}


@app.post("/api/stripe/webhook")
async def stripe_webhook(request: Request):
    """
    Stripe will POST events here.
    We verify signature, then on checkout.session.completed we grant credits.
    """
    _require_stripe_env()
    stripe.api_key = STRIPE_SECRET_KEY

    payload = await request.body()
    sig_header = request.headers.get("stripe-signature", "")

    try:
        event = stripe.Webhook.construct_event(payload, sig_header, STRIPE_WEBHOOK_SECRET)
    except Exception:
        # Do NOT leak details here
        return JSONResponse({"ok": False, "error": "Invalid signature"}, status_code=400)

    ws_events = _get_events_ws()

    event_id = event.get("id", "")
    if event_id and _events_has_event_id(ws_events, event_id):
        # Idempotent: already processed
        return {"ok": True, "received": True, "deduped": True}

    event_type = event.get("type", "")
    created_at = _utc_now_iso()

    # We only fulfill on successful checkout completion
    if event_type == "checkout.session.completed":
        session = event["data"]["object"]

        session_id = session.get("id", "")
        payment_intent = session.get("payment_intent", "")
        meta = session.get("metadata") or {}

        plan = (meta.get("plan") or "10pack").strip()
        token = (meta.get("token") or "").strip()

        credits_to_add = _credits_for_plan(plan)

        ws_tokens = _get_tokens_ws()
        try:
            token_result = _add_credits_to_token(ws_tokens, token, plan, credits_to_add)
            status = "fulfilled"
            note = "credits granted"
        except Exception as e:
            token_result = {"token": token, "plan": plan, "remaining": ""}
            status = "fulfillment_error"
            note = str(e)

        _append_event(
            ws_events,
            {
                "event_id": event_id,
                "created_at_utc": created_at,
                "session_id": session_id,
                "payment_intent": payment_intent,
                "status": status,
                "token": token_result.get("token", token),
                "plan": plan,
                "credits_added": credits_to_add if status == "fulfilled" else "",
                "note": note,
            },
        )

        return {"ok": True, "received": True}

    # Log other event types as received (optional but helpful)
    _append_event(
        ws_events,
        {
            "event_id": event_id,
            "created_at_utc": created_at,
            "session_id": "",
            "payment_intent": "",
            "status": f"ignored:{event_type}",
            "token": "",
            "plan": "",
            "credits_added": "",
            "note": "",
        },
    )

    return {"ok": True, "received": True}


@app.get("/api/stripe/session-result")
def stripe_session_result(session_id: str):
    """
    UI calls this after redirect from Stripe success page.
    It looks up the fulfillment record in Google Sheets by session_id
    and returns the token + latest credits.
    """
    ws_events = _get_events_ws()
    rec = _find_event_by_session_id(ws_events, session_id)
    if not rec:
        return {"ok": False, "found": False, "session_id": session_id}

    token = rec.get("token", "")
    plan = rec.get("plan", "")
    status = rec.get("status", "")

    remaining = None
    if token:
        ws_tokens = _get_tokens_ws()
        t = _get_token_record(ws_tokens, token)
        if t:
            remaining = _parse_remaining(t.get("remaining", ""))

    return {
        "ok": True,
        "found": True,
        "session_id": session_id,
        "status": status,
        "token": token,
        "plan": plan,
        "remaining": remaining,
    }


# UI last (serve static/index.html at /) — only if folder exists
if os.path.isdir("static"):
    app.mount("/", StaticFiles(directory="static", html=True), name="ui")

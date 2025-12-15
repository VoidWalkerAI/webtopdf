import os
import re
import uuid
import datetime as dt

from fastapi import FastAPI, Body, HTTPException
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

def _utc_now_iso():
    return dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"

def _is_probably_url(s: str) -> bool:
    return bool(re.match(r"^https?://", (s or "").strip(), flags=re.I))

@app.get("/api/health")
def health():
    return {"ok": True, "service": "webtopdf", "time": _utc_now_iso()}

@app.post("/api/web-to-pdf")
def web_to_pdf(payload: dict = Body(...)):
    url = (payload.get("url") or "").strip()
    if not _is_probably_url(url):
        raise HTTPException(status_code=400, detail="URL must start with http:// or https://")

    # Safety: basic size guard (you can expand later)
    if len(url) > 2000:
        raise HTTPException(status_code=400, detail="URL too long.")

    job_id = uuid.uuid4().hex[:10].upper()

    # Playwright / Chromium PDF render
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=[
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-gpu",
                ],
            )
            context = browser.new_context()
            page = context.new_page()

            # Load page (timeout is the big lever)
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

    filename = f"webtopdf_{job_id}.pdf"
    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "X-Job-Id": job_id,
        },
    )

# UI last (serve static/index.html at /) — only if folder exists
if os.path.isdir("static"):
    app.mount("/", StaticFiles(directory="static", html=True), name="ui")

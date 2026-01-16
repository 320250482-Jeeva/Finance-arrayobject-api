import os
import json
import logging
import mimetypes
import requests
from io import BytesIO
from datetime import datetime
from typing import List, Dict, Any, Optional

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel, EmailStr
from pptx import Presentation

# ============================================================================
# Router
# ============================================================================

router = APIRouter(
    prefix="/api/v1",
    tags=["PPT Email"]
)

# ============================================================================
# Config
# ============================================================================

TOKEN_URL = "https://login.microsoftonline.com/1a407a2d-7675-4d17-8692-b3ac285306e4/oauth2/v2.0/token"
EMAIL_URL = "https://dev.apps.api.it.philips.com/api/email"

CLIENT_ID = os.getenv("CLIENT_ID", "826bc22b-bb13-471b-a9c3-10cfb0b11a83")
CLIENT_SECRET = os.getenv("CLIENT_SECRET", "Yy48Q~-5lglgh.GXF13sOp.qwYIFJq0-XPjPdc7O")
SCOPE = os.getenv("SCOPE", "api://itaap-common-email-service-non-prod/.default")

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ============================================================================
# Models
# ============================================================================

class PPTEmailRequest(BaseModel):
    business_name: str
    summary: str
    data: List[Dict[str, Any]]
    email: EmailStr
    cc_emails: Optional[List[EmailStr]] = None
    bcc_emails: Optional[List[EmailStr]] = None
    subject: Optional[str] = None
    body: Optional[str] = None


class APIResponse(BaseModel):
    success: bool
    message: str
    request_id: str
    timestamp: str
    pptx_filename: str
    email_status_code: Optional[int] = None

# ============================================================================
# Helpers (kept local to router)
# ============================================================================

def get_bearer_token() -> str:
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": SCOPE,
    }
    resp = requests.post(TOKEN_URL, data=data, timeout=60)
    resp.raise_for_status()
    return resp.json()["access_token"]


def send_email(
    token: str,
    to_emails: List[str],
    subject: str,
    body: str,
    attachment: BytesIO,
    filename: str,
    cc: List[str],
    bcc: List[str],
):
    headers = {"Authorization": f"Bearer {token}"}

    payload = {
        "to": to_emails,
        "cc": cc,
        "bcc": bcc,
        "subject": subject,
        "body": body,
        "platformName": "ITAAP",
        "projectName": "email-service",
        "priority": "1",
        "url": "email-attachments/",
    }

    files = {
        "email-data": ("email.json", json.dumps(payload), "application/json"),
        "attachment": (filename, attachment, "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
    }

    return requests.post(EMAIL_URL, headers=headers, files=files, timeout=60)

# ============================================================================
# Routes
# ============================================================================

@router.get("/health")
def health():
    return {"status": "healthy", "timestamp": datetime.now().isoformat()}


@router.post("/generate-and-send", response_model=APIResponse)
async def generate_and_send(request: PPTEmailRequest):
    request_id = datetime.now().strftime("%Y%m%d%H%M%S")

    # Create PPT
    ppt_buffer = BytesIO()
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[0])
    prs.save(ppt_buffer)
    ppt_buffer.seek(0)

    try:
        token = get_bearer_token()
        response = send_email(
            token=token,
            to_emails=[request.email],
            subject=request.subject or "Analysis Report",
            body=request.body or "<h3>Please find attached</h3>",
            attachment=ppt_buffer,
            filename=f"{request.business_name}.pptx",
            cc=request.cc_emails or [],
            bcc=request.bcc_emails or [],
        )

        response.raise_for_status()

        return APIResponse(
            success=True,
            message="PPT generated and email sent",
            request_id=request_id,
            timestamp=datetime.now().isoformat(),
            pptx_filename=f"{request.business_name}.pptx",
            email_status_code=response.status_code,
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

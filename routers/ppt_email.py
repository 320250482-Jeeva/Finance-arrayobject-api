"""
PPT Email Service API - Router
Integrated with Existing Email Functions
"""
import os
import json
import logging
from io import BytesIO
from typing import List, Dict, Any, Optional
from datetime import datetime
import mimetypes
import re

import requests
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel, EmailStr, Field
from fastapi.responses import JSONResponse

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# ============================================================================
# Router
# ============================================================================

router = APIRouter(
    prefix="/api/v1",
    tags=["PPT Email"]
)

# ============================================================================
# Logging
# ============================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)

# ============================================================================
# Configuration
# ============================================================================

TOKEN_URL = "https://login.microsoftonline.com/1a407a2d-7675-4d17-8692-b3ac285306e4/oauth2/v2.0/token"
EMAIL_URL = "https://dev.apps.api.it.philips.com/api/email"

CLIENT_ID = os.getenv("CLIENT_ID", "826bc22b-bb13-471b-a9c3-10cfb0b11a83")
CLIENT_SECRET = os.getenv("CLIENT_SECRET", "Yy48Q~-5lglgh.GXF13sOp.qwYIFJq0-XPjPdc7O")
SCOPE = os.getenv("SCOPE", "api://itaap-common-email-service-non-prod/.default")

MIME_OVERRIDES = {
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}

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
# Helpers
# ============================================================================

def guess_mime(filename: str) -> str:
    ext = os.path.splitext(filename.lower())[1]
    if ext in MIME_OVERRIDES:
        return MIME_OVERRIDES[ext]
    mt, _ = mimetypes.guess_type(filename)
    return mt or "application/octet-stream"


def get_bearer_token(
    client_id: Optional[str] = None,
    client_secret: Optional[str] = None,
    scope: Optional[str] = None,
) -> str:
    client_id = client_id or CLIENT_ID
    client_secret = client_secret or CLIENT_SECRET
    scope = scope or SCOPE

    if not client_secret:
        raise RuntimeError("CLIENT_SECRET not set")

    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "scope": scope,
        "client_secret": client_secret,
    }

    resp = requests.post(TOKEN_URL, data=data, timeout=60)
    resp.raise_for_status()

    token = resp.json().get("access_token")
    if not token:
        raise RuntimeError("No access token returned")

    return token


def send_email(
    bearer_token: str,
    to_emails: List[str],
    cc_emails: List[str],
    bcc_emails: List[str],
    subject: str,
    body: str,
    attachment_buffer: BytesIO,
    attachment_name: str,
):
    payload = {
        "to": to_emails,
        "cc": cc_emails,
        "bcc": bcc_emails,
        "subject": subject,
        "body": body,
        "platformName": "ITAAP",
        "projectName": "email-service",
        "priority": "1",
        "url": "email-attachments/",
    }

    headers = {"Authorization": f"Bearer {bearer_token}"}

    files = {
        "email-data": ("email-data.json", json.dumps(payload), "application/json"),
        "attachment": (attachment_name, attachment_buffer, guess_mime(attachment_name)),
    }

    return requests.post(EMAIL_URL, headers=headers, files=files, timeout=60)

# ============================================================================
# PPT Generation (unchanged)
# ============================================================================

def _parse_summary(summary: str) -> List[str]:
    clean = (summary or "").strip()
    if not clean:
        return ["No summary provided"]

    numbered = re.findall(r"(?:^|\s)(\d+)\.\s", clean)
    bullets = []

    if numbered:
        parts = re.split(r"(?:^|\s)(?=\d+\.\s)", clean)
        for p in parts:
            p = re.sub(r"^\d+\.\s*", "", p).strip()
            if p:
                bullets.append(p)
    else:
        bullets = [b.strip("-• ").strip() for b in clean.splitlines() if b.strip()]

    return bullets or ["No summary provided"]


def create_pptx_buffer(business_name: str, summary: str, data: List[Dict[str, Any]]) -> BytesIO:
    """
    Create PPTX with v2.2 formatting:
    - Slide 1: Title (template background)
    - Slide 2: Summary (white background, centered)
    - Slide 3: Data table (white background, centered, negatives in red)
    - Slide 4: Thank you (template background)
    """
    logger.info(f"Generating PPTX for: {business_name}")
    
    prs = Presentation()
    
    # =====================================================================
    # SLIDE 1: Title (Template Background)
    # =====================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Remove placeholders
    for shape in list(slide.shapes):
        if shape.is_placeholder:
            sp = shape.element
            sp.getparent().remove(sp)
    
    # Add title
    title_text = f"{business_name} - Analysis"
    left = Inches(0.5)
    top = Inches(2.5)
    width = Inches(8)
    height = Inches(2)
    
    title_box = slide.shapes.add_textbox(left, top, width, height)
    text_frame = title_box.text_frame
    text_frame.word_wrap = True
    
    p = text_frame.paragraphs[0]
    p.text = title_text
    p.font.size = Pt(60)
    p.font.bold = True
    from pptx.dml.color import RGBColor
    p.font.color.rgb = RGBColor(79, 129, 189)  # Dark Blue Accent 1 Light 60%
    p.alignment = PP_ALIGN.CENTER
    
    # =====================================================================
    # SLIDE 2: Summary (White Background)
    # =====================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Set white background
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    # Remove placeholders
    for shape in list(slide.shapes):
        if shape.is_placeholder:
            sp = shape.element
            sp.getparent().remove(sp)
    
    # Add Summary title
    title_left = Inches(0.75)
    title_top = Inches(0.5)
    title_width = Inches(8.5)
    title_height = Inches(0.8)
    
    title_box = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_frame = title_box.text_frame
    title_p = title_frame.paragraphs[0]
    title_p.text = "SUMMARY"
    title_p.font.size = Pt(32)
    title_p.font.bold = True
    title_p.alignment = PP_ALIGN.LEFT
    
    # Parse and add summary
    bullets = _parse_summary(summary)
    
    content_left = Inches(0.75)
    content_top = Inches(1.5)
    content_width = Inches(7.5)
    content_height = Inches(4.5)
    
    content_box = slide.shapes.add_textbox(content_left, content_top, content_width, content_height)
    text_frame = content_box.text_frame
    text_frame.word_wrap = True
    text_frame.clear()
    
    for i, bullet_text in enumerate(bullets):
        if i == 0:
            p = text_frame.paragraphs[0]
        else:
            p = text_frame.add_paragraph()
        
        p.text = bullet_text
        p.font.size = Pt(14)
        p.level = 0
        p.space_before = Pt(8)
        p.space_after = Pt(8)
        p.alignment = PP_ALIGN.LEFT
    
    # =====================================================================
    # SLIDE 3: Data Table (White Background)
    # =====================================================================
    if data:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # Set white background
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(255, 255, 255)
        
        # Remove placeholders
        for shape in list(slide.shapes):
            if shape.is_placeholder:
                sp = shape.element
                sp.getparent().remove(sp)
        
        # Add Data title
        title_left = Inches(0.75)
        title_top = Inches(0.5)
        title_width = Inches(8.5)
        title_height = Inches(0.8)
        
        title_box = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
        title_frame = title_box.text_frame
        title_p = title_frame.paragraphs[0]
        title_p.text = "DATA"
        title_p.font.size = Pt(32)
        title_p.font.bold = True
        title_p.alignment = PP_ALIGN.LEFT
        
        # Create table
        keys = list(data[0].keys())
        num_rows = len(data) + 1
        num_cols = len(keys)
        
        table_width = Inches(8)
        left = Inches(0.5)
        top = Inches(1.5)
        height = Inches(4.5)
        
        table_shape = slide.shapes.add_table(num_rows, num_cols, left, top, table_width, height)
        table = table_shape.table
        
        # Set column widths
        col_width = table_width / num_cols
        for col in table.columns:
            col.width = int(col_width)
        
        # Header row
        for col_idx, key in enumerate(keys):
            cell = table.cell(0, col_idx)
            cell_text_frame = cell.text_frame
            cell_text_frame.clear()
            p = cell_text_frame.paragraphs[0]
            p.text = str(key)
            p.font.bold = True
            p.font.size = Pt(12)
            p.font.color.rgb = RGBColor(255, 255, 255)
            p.alignment = PP_ALIGN.CENTER
            
            fill = cell.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(0, 51, 102)
        
        # Data rows
        for row_idx, row_data in enumerate(data, start=1):
            for col_idx, key in enumerate(keys):
                cell = table.cell(row_idx, col_idx)
                cell_value = str(row_data.get(key, ""))
                
                cell_text_frame = cell.text_frame
                cell_text_frame.clear()
                p = cell_text_frame.paragraphs[0]
                p.text = cell_value
                p.font.size = Pt(12)
                p.alignment = PP_ALIGN.CENTER
                
                # Red color for negative values
                if cell_value.strip().startswith('-'):
                    p.font.color.rgb = RGBColor(255, 0, 0)
                
                # Alternate row colors
                if row_idx % 2 == 0:
                    fill = cell.fill
                    fill.solid()
                    fill.fore_color.rgb = RGBColor(242, 242, 242)
    
    # =====================================================================
    # SLIDE 4: Thank You - Template Background
    # =====================================================================
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Remove placeholders
    for shape in list(slide.shapes):
        if shape.is_placeholder:
            sp = shape.element
            sp.getparent().remove(sp)
    
    # Add thank you
    thank_you_left = Inches(0.5)
    thank_you_top = Inches(1.5)
    thank_you_width = Inches(8)
    thank_you_height = Inches(1.5)
    
    thank_you_box = slide.shapes.add_textbox(thank_you_left, thank_you_top, thank_you_width, thank_you_height)
    thank_you_frame = thank_you_box.text_frame
    thank_you_p = thank_you_frame.paragraphs[0]
    thank_you_p.text = "THANK YOU"
    thank_you_p.font.size = Pt(60)
    thank_you_p.font.bold = True
    thank_you_p.alignment = PP_ALIGN.CENTER
    
    # Save to BytesIO
    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    
    logger.info("PPTX created successfully")
    return buf

# ============================================================================
# Routes
# ============================================================================

@router.get("/health")
async def health_check():
    return {
        "status": "healthy",
        "service": "PPT Email Service",
        "timestamp": datetime.now().isoformat(),
    }


@router.post("/generate-and-send", response_model=APIResponse)
async def generate_and_send(request: PPTEmailRequest):
    request_id = f"{datetime.now().strftime('%Y%m%d%H%M%S')}-{hash(request.email) % 10000}"
    timestamp = datetime.now().isoformat()

    pptx_buffer = create_pptx_buffer(
        request.business_name, request.summary, request.data
    )

    token = get_bearer_token()

    response = send_email(
        bearer_token=token,
        to_emails=[request.email],
        cc_emails=request.cc_emails or [],
        bcc_emails=request.bcc_emails or [],
        subject=request.subject or f"{request.business_name} - Analysis Report",
        body=request.body or "<p>Please find the report attached.</p>",
        attachment_buffer=pptx_buffer,
        attachment_name=f"{request.business_name}.pptx",
    )

    response.raise_for_status()

    return APIResponse(
        success=True,
        message="PPTX generated and email sent successfully",
        request_id=request_id,
        timestamp=timestamp,
        pptx_filename=f"{request.business_name}.pptx",
        email_status_code=response.status_code,
    )

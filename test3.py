# app.py
"""
Profile & Resume Analyzer — LeetCode | GitHub | ATS
Reads a Google Sheet (CSV), fetches LeetCode/GitHub metrics, downloads resumes (Drive or direct PDF),
computes ATS score (PyPDF2 with optional OCR fallback), builds Excel/PDF reports, and can email top candidates.
"""

import os
import time
import io
import re
import json
import requests
import smtplib
from email.message import EmailMessage
from typing import List, Dict, Any, Optional, Tuple
import pandas as pd
import streamlit as st
import plotly.express as px
import matplotlib.pyplot as plt
from dotenv import load_dotenv
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import utils

# Optional imports for Google Sheets write-back and OCR
try:
    import gspread
    from google.oauth2.service_account import Credentials
except Exception:
    gspread = None
    Credentials = None

try:
    from PyPDF2 import PdfReader
except Exception:
    PdfReader = None

# pytesseract + PIL + pdf2image for OCR fallback (optional)
try:
    import pytesseract
    from PIL import Image
    from pdf2image import convert_from_bytes
    OCR_AVAILABLE = True
except Exception:
    pytesseract = None
    Image = None
    convert_from_bytes = None
    OCR_AVAILABLE = False

# Try to import user's mail helper if present (akshi.py). If not present, we'll fall back to an internal sender.
try:
    from akshi import send_bulk_emails as send_bulk_emails_external  # type: ignore
except Exception:
    send_bulk_emails_external = None  # fallback later

# External services (your implementations in services/)
try:
    from services.leetcode_service import get_leetcode_data
except Exception:
    def get_leetcode_data(username: str) -> Dict[str, Any]:
        return {"error": "LeetCode service not available."}

try:
    from services.github_service import get_github_data
except Exception:
    def get_github_data(username: str) -> Dict[str, Any]:
        return {"error": "GitHub service not available."}

# Attempt to use project's ats_service if present
try:
    from services import ats_service  # should have get_ats_score_from_pdf, get_ats_score
except Exception:
    ats_service = None

# Excel exporter util (your existing module)
try:
    from utils.excel_generator import create_excel_report_bytes
except Exception:
    def create_excel_report_bytes(df: pd.DataFrame) -> bytes:
        buf = io.BytesIO()
        df.to_csv(buf, index=False)
        buf.seek(0)
        return buf.read()

# ---------------------------------------------
# SETUP
# ---------------------------------------------
load_dotenv()
st.set_page_config(page_title="Profile & Resume Analyzer", layout="wide")
st.title("Profile & Resume Analyzer — LeetCode | GitHub | ATS")
st.write("Paste a Google Sheet link (regular sheet URL or published CSV), or enter names/emails/usernames or upload PDFs.")

# ---------------------------------------------
# Utilities: Google Sheet URL normalization & fetch
# ---------------------------------------------
def normalize_sheet_to_csv_url(url: str) -> Optional[str]:
    if not url:
        return None
    url = url.strip()
    # Already a csv/export url
    if "output=csv" in url or "export?format=csv" in url:
        return url
    # Standard sheets URL with /d/<id>/
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if m:
        file_id = m.group(1)
        return f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=csv"
    # Older published form /d/e/<id>/
    m2 = re.search(r"/spreadsheets/d/e/([a-zA-Z0-9-_]+)", url)
    if m2:
        file_id = m2.group(1)
        return f"https://docs.google.com/spreadsheets/d/e/{file_id}/pub?output=csv"
    # If it's a Drive link return None (we handle drive downloads separately)
    if "drive.google.com" in url:
        return None
    # otherwise assume direct CSV url
    if url.startswith("http://") or url.startswith("https://"):
        return url
    return None

@st.cache_data(ttl=300)
def fetch_csv_from_url(url: str, timeout: int = 10) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    try:
        headers = {"User-Agent": "Profile-Analyzer/1.0"}
        r = requests.get(url, headers=headers, timeout=timeout)
        if r.status_code != 200:
            return None, f"HTTP {r.status_code} when requesting the sheet (URL: {url})."
        try:
            df = pd.read_csv(io.StringIO(r.text))
            return df, None
        except Exception:
            try:
                df = pd.read_csv(io.BytesIO(r.content))
                return df, None
            except Exception as e:
                snippet = r.content[:500].decode("utf-8", errors="replace")
                return None, f"Failed to parse CSV. Response snippet:\n{snippet}\nError: {e}"
    except requests.exceptions.RequestException as e:
        return None, f"Request error when fetching sheet: {e}"
    except Exception as e:
        return None, f"Unexpected error when fetching sheet: {e}"

def normalize_sheet_columns(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {}
    for c in df.columns:
        lc = c.strip().lower()
        if lc in ("name", "candidate_name", "full_name", "applicant_name"):
            mapping[c] = "candidate_name"
            continue
        if "email" in lc or "e-mail" in lc:
            mapping[c] = "candidate_email"
            continue
        if "leetcode" in lc and ("user" in lc or "name" in lc or "username" in lc):
            mapping[c] = "leetcode_username"
            continue
        if lc in ("leetcode", "leetcode_username"):
            mapping[c] = "leetcode_username"
            continue
        if "github" in lc and ("user" in lc or "name" in lc or "username" in lc):
            mapping[c] = "github_username"
            continue
        if lc in ("github", "github_username"):
            mapping[c] = "github_username"
            continue
        if "pdf" in lc or "resume" in lc or "cv" in lc or "upload" in lc or "link" in lc:
            mapping[c] = "resume_link"
            continue
    df = df.rename(columns=mapping)
    return df

@st.cache_data(ttl=300)
def fetch_candidates_from_sheet_url(sheet_url: str) -> pd.DataFrame:
    if not sheet_url:
        raise ValueError("No Google Sheet URL provided.")
    normalized = normalize_sheet_to_csv_url(sheet_url)
    if not normalized:
        raise ValueError(
            "Unable to convert the provided URL into a CSV-export URL. "
            "If your sheet is a Google Sheet, use the full sheet URL (https://docs.google.com/spreadsheets/d/<id>/edit) "
            "or the published CSV link (File → Publish to the web → CSV). "
            "If you pasted a Google Drive file link (drive.google.com), convert it to a Google Sheet or publish as CSV."
        )
    df, err = fetch_csv_from_url(normalized)
    if df is None:
        raise ValueError(f"Failed to fetch or parse CSV: {err}")
    df = normalize_sheet_columns(df)
    for col in ("candidate_name", "candidate_email", "leetcode_username", "github_username", "resume_link"):
        if col not in df.columns:
            df[col] = ""
    return df

# ---------------------------------------------
# Helpers for resume downloads and external services
# ---------------------------------------------
def download_drive_pdf(url: str) -> Optional[bytes]:
    """
    Improved Drive downloader:
    - Extract file id from many common patterns
    - Try multiple uc/export/download endpoints
    - Validate response looks like a PDF (starts with %PDF or content-type)
    - Return None if can't download (file not public / requires auth / scanned/unreadable)
    """
    if not url:
        return None
    try:
        url = str(url).strip()
        m = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
        if not m:
            m = re.search(r"id=([a-zA-Z0-9_-]+)", url)
        if not m:
            return None
        file_id = m.group(1)
        headers = {"User-Agent": "Profile-Analyzer/1.0"}
        candidates = [
            f"https://drive.google.com/uc?export=download&id={file_id}",
            f"https://docs.google.com/uc?export=download&id={file_id}",
            f"https://drive.google.com/uc?export=open&id={file_id}",
            f"https://drive.google.com/file/d/{file_id}/view"
        ]
        for dl_url in candidates:
            try:
                r = requests.get(dl_url, headers=headers, timeout=20, allow_redirects=True)
            except Exception:
                continue
            if r.status_code != 200:
                continue
            content = r.content or b""
            content_type = r.headers.get("Content-Type", "").lower()
            if content.startswith(b"%PDF"):
                return content
            if "application/pdf" in content_type:
                return content
            idx = content.find(b"%PDF")
            if idx != -1:
                return content[idx:]
        return None
    except Exception:
        return None

# ---------------------------------------------
# ATS: fallback implementations (if services.ats_service not present)
# ---------------------------------------------
ATS_KEYWORDS = [
    "problem solving", "communication", "teamwork", "leadership",
    "conflict resolution", "strategic thinking", "time management",
    "decision making", "innovation", "adaptability", "collaboration",
    "analytical thinking", "critical thinking", "creativity", "initiative",
    "data analysis", "machine learning", "deep learning", "natural language processing",
    "python", "java", "c++", "c", "sql", "html", "css", "javascript",
    "react", "node.js", "flask", "django", "git", "github", "mongodb",
    "data structures", "algorithms", "object oriented programming",
    "tensorflow", "keras", "pandas", "numpy", "scikit-learn", "excel"
]

def extract_text_from_pdf_bytes(file_bytes: bytes) -> str:
    text = ""
    if PdfReader:
        try:
            with io.BytesIO(file_bytes) as pdf_stream:
                reader = PdfReader(pdf_stream)
                for page in reader.pages:
                    try:
                        page_text = page.extract_text() or ""
                        text += page_text + "\n"
                    except Exception:
                        continue
        except Exception:
            text = ""
    if text.strip():
        return text.lower()
    # OCR fallback
    if OCR_AVAILABLE and convert_from_bytes and pytesseract:
        try:
            images = convert_from_bytes(file_bytes, first_page=1, last_page=5)
            ocr_text = ""
            for img in images:
                try:
                    ocr_text += pytesseract.image_to_string(img) + "\n"
                except Exception:
                    continue
            if ocr_text.strip():
                return ocr_text.lower()
        except Exception:
            pass
    return ""

def local_get_ats_score_from_pdf(file_bytes: bytes) -> Dict[str, Any]:
    if not file_bytes:
        return {"score": None, "remarks": "No file bytes provided."}
    text = extract_text_from_pdf_bytes(file_bytes)
    if not text.strip():
        return {"score": None, "remarks": "Unable to read text from PDF (scanned or protected)."}
    matched = [kw for kw in ATS_KEYWORDS if kw in text]
    score = round((len(matched) / max(1, len(ATS_KEYWORDS))) * 100)
    remarks = f"Found {len(matched)} of {len(ATS_KEYWORDS)} keywords: {', '.join(matched[:10])}"
    return {"score": score, "remarks": remarks}

def local_get_ats_score(resume_url: str) -> Dict[str, Any]:
    try:
        r = requests.get(resume_url, timeout=15)
        if r.status_code == 200:
            content = r.content or b""
            ct = (r.headers.get("Content-Type") or "").lower()
            if content.startswith(b"%PDF") or "application/pdf" in ct or resume_url.lower().endswith(".pdf"):
                return local_get_ats_score_from_pdf(content)
            return {"score": None, "remarks": "Provided URL is not a direct PDF link."}
        else:
            return {"score": None, "remarks": f"HTTP {r.status_code} fetching resume URL."}
    except Exception as e:
        return {"score": None, "remarks": f"Error fetching resume URL: {e}"}

def get_ats_score_from_pdf_bytes(file_bytes: Optional[bytes]) -> Dict[str, Any]:
    if not file_bytes:
        return {"skipped": True}
    if ats_service and hasattr(ats_service, "get_ats_score_from_pdf"):
        try:
            return ats_service.get_ats_score_from_pdf(file_bytes)
        except Exception:
            return local_get_ats_score_from_pdf(file_bytes)
    else:
        return local_get_ats_score_from_pdf(file_bytes)

def get_ats_score_from_url(resume_url: Optional[str]) -> Dict[str, Any]:
    if not resume_url:
        return {"skipped": True}
    if ats_service and hasattr(ats_service, "get_ats_score"):
        try:
            return ats_service.get_ats_score(resume_url)
        except Exception:
            return local_get_ats_score(resume_url)
    else:
        return local_get_ats_score(resume_url)

def fetch_ats_from_url_or_bytes(resume_url: Optional[str], file_bytes: Optional[bytes]) -> Dict[str, Any]:
    """
    Prefer file_bytes; if file_bytes not present and resume_url provided, try to download it (Drive support).
    Returns dictionary with 'score' and 'remarks', or skipped/error markers.
    """
    if file_bytes:
        try:
            return get_ats_score_from_pdf_bytes(file_bytes)
        except Exception as e:
            return {"error": f"ATS PDF error: {e}"}
    if resume_url:
        if "drive.google.com" in resume_url:
            file_bytes_dl = download_drive_pdf(resume_url)
            if file_bytes_dl:
                try:
                    return get_ats_score_from_pdf_bytes(file_bytes_dl)
                except Exception as e:
                    return {"error": f"ATS PDF error after download: {e}"}
            else:
                return {"error": "Could not download resume from Google Drive (public access required)."}
        try:
            return get_ats_score_from_url(resume_url)
        except Exception as e:
            return {"error": f"ATS URL error: {e}"}
    return {"skipped": True}

# ---------------------------------------------
# LeetCode & GitHub wrappers
# ---------------------------------------------
@st.cache_data(ttl=300)
def fetch_leetcode(username: Optional[str]) -> Dict[str, Any]:
    if not username:
        return {"skipped": True}
    try:
        return get_leetcode_data(username)
    except Exception as e:
        return {"error": f"LeetCode fetch error: {e}"}

@st.cache_data(ttl=300)
def fetch_github(username: Optional[str]) -> Dict[str, Any]:
    if not username:
        return {"skipped": True}
    try:
        return get_github_data(username)
    except Exception as e:
        return {"error": f"GitHub fetch error: {e}"}

# ---------------------------------------------
# Scoring and insights
# ---------------------------------------------
def compute_candidate_score(leet_res, git_res, ats_res) -> int:
    score = 0
    try:
        if leet_res and not leet_res.get("error") and not leet_res.get("skipped"):
            total = int(leet_res.get("total_solved", 0) or 0)
            score += min(40, total // 10)
    except Exception:
        pass
    try:
        if git_res and not git_res.get("error") and not git_res.get("skipped"):
            repos = int(git_res.get("total_repos", 0) or 0)
            score += min(30, repos * 2)
    except Exception:
        pass
    try:
        if ats_res and not ats_res.get("error") and not ats_res.get("skipped"):
            ats_score = ats_res.get("score", 0) or 0
            try:
                ats_score = int(ats_score)
            except Exception:
                ats_score = 0
            score += min(30, ats_score)
    except Exception:
        pass
    return int(score)

def recruiter_insights(leet_res, git_res, ats_res):
    strengths, improvements = [], []
    try:
        solved = int(leet_res.get("total_solved", 0)) if leet_res and not leet_res.get("error") else 0
        if solved >= 300:
            strengths.append("Strong problem-solving skills (LeetCode).")
        elif solved >= 100:
            strengths.append("Good problem-solving foundation on LeetCode.")
        else:
            improvements.append("Solve more DSA problems to strengthen fundamentals.")
    except Exception:
        pass
    try:
        probs = leet_res.get("problems_by_difficulty", {}) if leet_res else {}
        med = int(probs.get("Medium", 0) or 0)
        hard = int(probs.get("Hard", 0) or 0)
        if hard >= 10:
            strengths.append("Experienced with Hard-level problems.")
        elif med >= 20:
            strengths.append("Comfortable with Medium-level problems.")
    except Exception:
        pass
    try:
        repos = int(git_res.get("total_repos", 0)) if git_res and not git_res.get("error") else 0
        if repos > 5:
            strengths.append("Active GitHub contributor.")
        else:
            improvements.append("Add more quality projects on GitHub.")
    except Exception:
        pass
    try:
        ats = int(ats_res.get("score", 0)) if ats_res and not ats_res.get("error") and ats_res.get("score") is not None else 0
        if ats >= 75:
            strengths.append("Resume is ATS-friendly.")
        else:
            improvements.append("Optimize resume with stronger keywords.")
    except Exception:
        pass
    return list(dict.fromkeys(strengths)), list(dict.fromkeys(improvements))

# ---------------------------------------------
# PDF/chart helpers and PDF export
# ---------------------------------------------
def _save_matplotlib_bar(df_long: pd.DataFrame) -> io.BytesIO:
    try:
        plt.style.use("seaborn-v0_8-colorblind")
    except Exception:
        plt.style.use("default")
    fig, ax = plt.subplots(figsize=(10, 6))
    pivot = df_long.pivot(index="display_name", columns="metric", values="value").fillna(0)
    pivot.plot(kind="bar", ax=ax)
    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150)
    plt.close(fig)
    buf.seek(0)
    return buf

def _create_pdf_bytes(rows, title="Profile Report"):
    df = pd.DataFrame(rows).fillna("")
    top = df.sort_values("candidate_score", ascending=False).head(10)
    long_rows = []
    for _, r in top.iterrows():
        long_rows.extend([
            {"display_name": r.get("display_name", ""), "metric": "LeetCode Solved", "value": int(r.get("leetcode.total_solved") or 0)},
            {"display_name": r.get("display_name", ""), "metric": "GitHub Repos", "value": int(r.get("github.total_repos") or 0)},
            {"display_name": r.get("display_name", ""), "metric": "ATS Score", "value": int(r.get("ats.score") or 0) if r.get("ats.score") not in (None, "") else 0},
            {"display_name": r.get("display_name", ""), "metric": "Candidate Score", "value": int(r.get("candidate_score") or 0)},
        ])
    df_long = pd.DataFrame(long_rows) if long_rows else pd.DataFrame([{"display_name":"", "metric":"", "value":0}])
    chart_buf = _save_matplotlib_bar(df_long)
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter)
    styles = getSampleStyleSheet()
    story = [Paragraph(title, styles["Title"]), Spacer(1, 12)]
    story.append(Paragraph(f"Generated at: {time.strftime('%Y-%m-%d %H:%M:%S')}", styles["Normal"]))
    story.append(Spacer(1, 12))
    try:
        story.append(RLImage(chart_buf, width=450, height=300))
    except Exception:
        story.append(Paragraph("Chart not available", styles["Normal"]))
    tbl_data = [["Name", "Email", "LeetCode", "GitHub", "ATS", "Score"]]
    for _, r in top.iterrows():
        tbl_data.append([
            r.get("candidate_name", ""),
            r.get("candidate_email", ""),
            int(r.get("leetcode.total_solved") or 0),
            int(r.get("github.total_repos") or 0),
            int(r.get("ats.score") or 0) if r.get("ats.score") not in (None, "") else 0,
            int(r.get("candidate_score") or 0),
        ])
    story.append(Table(tbl_data, hAlign="LEFT"))
    doc.build(story)
    buf.seek(0)
    return buf.read()

# ---------------------------------------------
# Email sending with fallback
# ---------------------------------------------
def send_bulk_emails_fallback(df: pd.DataFrame, sender: str, app_password: str) -> Dict[str, Any]:
    summary = {"sent": 0, "failed": 0, "errors": []}
    if not sender or not app_password:
        raise ValueError("Missing sender email or app password for sending emails.")
    if "candidate_email" not in df.columns:
        raise ValueError("DataFrame must contain candidate_email column to send emails.")
    smtp_host = os.getenv("SMTP_HOST", "smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT", "465"))
    try:
        with smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=30) as smtp:
            smtp.login(sender, app_password)
            for _, row in df.iterrows():
                to_email = str(row.get("candidate_email") or "").strip()
                if not to_email:
                    summary["failed"] += 1
                    summary["errors"].append(f"Missing email for {row.get('display_name')}")
                    continue
                try:
                    msg = EmailMessage()
                    msg["From"] = sender
                    msg["To"] = to_email
                    name = row.get("candidate_name") or row.get("display_name") or ""
                    score = row.get("candidate_score", "")
                    msg["Subject"] = f"Profile Report — {name or 'Candidate'}"
                    body = f"Hello {name or ''},\n\nWe have evaluated your profile. Your candidate score: {score}.\n\nBest regards,\nRecruiting Team"
                    msg.set_content(body)
                    smtp.send_message(msg)
                    summary["sent"] += 1
                except Exception as e:
                    summary["failed"] += 1
                    summary["errors"].append(f"Error sending to {to_email}: {e}")
    except Exception as e:
        raise RuntimeError(f"SMTP connection/login failed: {e}")
    return summary

def send_bulk_emails_safe(df: pd.DataFrame, sender: Optional[str], app_password: Optional[str]) -> Dict[str, Any]:
    if send_bulk_emails_external:
        try:
            send_bulk_emails_external(df, sender, app_password)
            return {"sent": int(len(df)), "failed": 0, "errors": []}
        except Exception as e:
            try:
                return send_bulk_emails_fallback(df, sender or "", app_password or "")
            except Exception as e2:
                raise RuntimeError(f"External sender failed: {e}; fallback failed: {e2}")
    else:
        return send_bulk_emails_fallback(df, sender or "", app_password or "")

# ---------------------------------------------
# Google Sheets write helpers (optional)
# ---------------------------------------------
def _get_sheet_id_from_url(url: str) -> Optional[str]:
    if not url:
        return None
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if m:
        return m.group(1)
    m = re.search(r"id=([a-zA-Z0-9-_]+)", url)
    return m.group(1) if m else None

def get_gspread_client() -> Optional["gspread.Client"]:
    if gspread is None or Credentials is None:
        return None
    sa_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    sa_file = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE")
    creds = None
    try:
        if sa_json:
            data = json.loads(sa_json)
            creds = Credentials.from_service_account_info(data, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
        elif sa_file and os.path.exists(sa_file):
            creds = Credentials.from_service_account_file(sa_file, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
        else:
            default = "google-credentials.json"
            if os.path.exists(default):
                creds = Credentials.from_service_account_file(default, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
    except Exception as e:
        st.error(f"Error loading Google service account credentials: {e}")
        return None
    if not creds:
        return None
    try:
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"Failed to create gspread client: {e}")
        return None

def append_rows_to_sheet(sheet_url: str, rows: List[List[Any]], worksheet_name: Optional[str] = None) -> Tuple[bool, str]:
    sheet_id = _get_sheet_id_from_url(sheet_url)
    if not sheet_id:
        return False, "Could not parse sheet id from URL."
    client = get_gspread_client()
    if not client:
        return False, "gspread client not available or credentials missing."
    try:
        ss = client.open_by_key(sheet_id)
        if worksheet_name:
            try:
                ws = ss.worksheet(worksheet_name)
            except Exception:
                ws = ss.add_worksheet(title=worksheet_name, rows="1000", cols="20")
        else:
            ws = ss.get_worksheet(0)
        ws.append_rows(rows, value_input_option="USER_ENTERED")
        try:
            st.cache_data.clear()
        except Exception:
            pass
        return True, "Rows appended successfully."
    except Exception as e:
        return False, f"Error appending to sheet: {e}"

# ---------------------------------------------
# UI: Data source selection
# ---------------------------------------------
st.markdown("### Data Source Options")
use_google_sheet = st.checkbox("📥 Fetch candidate data automatically from Google Sheet", value=True)

sheet_url_input = ""
if use_google_sheet:
    st.markdown("Provide your Google Sheet link (regular sheet URL or published CSV URL).")
    sheet_url_input = st.text_input("Google Sheet URL or CSV link", value="", help=(
        "Examples:\n"
        "- Regular sheet URL: https://docs.google.com/spreadsheets/d/<file_id>/edit#gid=0\n"
        "- Published CSV URL: https://docs.google.com/spreadsheets/d/<file_id>/export?format=csv\n"
        "- Published 'Publish to the web' CSV link with output=csv"
    ))
    validate_button = st.button("🔎 Validate & Preview Sheet")
    if validate_button:
        if not sheet_url_input:
            st.error("Please paste your Google Sheet URL or CSV link.")
        else:
            try:
                st.info("Attempting to fetch the sheet...")
                df_preview = fetch_candidates_from_sheet_url(sheet_url_input)
                st.success("Sheet fetched successfully — showing first 5 rows.")
                st.dataframe(df_preview.head(5))
            except Exception as e:
                st.error(f"Failed to fetch the sheet: {e}")

if not use_google_sheet:
    left, right = st.columns(2)
    with left:
        candidate_names = st.text_area("Candidate Names (one per line or comma-separated)", height=180, placeholder="John Doe\nJane Smith")
        leet_bulk = st.text_area("LeetCode usernames (one per line or comma-separated)", height=180, placeholder="tourist\nlabuladong")
    with right:
        candidate_emails = st.text_area("Candidate Emails (one per line or comma-separated)", height=180, placeholder="john@example.com\njane@example.com")
        git_bulk = st.text_area("GitHub usernames (one per line or comma-separated)", height=180, placeholder="torvalds\noctocat")
    st.markdown("---")
    uploaded_files = st.file_uploader("Upload PDF Resumes (multiple)", type=["pdf"], accept_multiple_files=True)
else:
    st.info("Using Google Sheet: all candidates and resumes will be fetched automatically from the provided link.")

# ---------------------------------------------
# Helper: split multiline/comma-separated input
# ---------------------------------------------
def _split_multi(value: str, max_items: int = 100) -> List[str]:
    if not value:
        return []
    parts: List[str] = []
    for line in value.splitlines():
        parts.extend([p.strip() for p in line.split(",")])
    parts = [p for p in parts if p]
    return parts[:max_items]

# ---------------------------------------------
# MAIN PROCESSING LOGIC (Generate Report)
# ---------------------------------------------
if st.button("🚀 Generate Report"):
    start_time = time.time()
    # gather inputs
    if use_google_sheet:
        if not sheet_url_input:
            st.error("Please provide your Google Sheet URL first.")
            st.stop()
        try:
            df_sheet = fetch_candidates_from_sheet_url(sheet_url_input)
        except Exception as e:
            st.error(f"Google Sheet fetch error: {e}")
            st.stop()
        if df_sheet.empty:
            st.error("Google Sheet is empty or inaccessible.")
            st.stop()
        names = df_sheet["candidate_name"].fillna("").astype(str).tolist()
        emails = df_sheet["candidate_email"].fillna("").astype(str).tolist()
        leet_users = df_sheet["leetcode_username"].fillna("").astype(str).tolist()
        git_users = df_sheet["github_username"].fillna("").astype(str).tolist()
        resume_links = df_sheet["resume_link"].fillna("").astype(str).tolist()
        # Attempt to download pdf bytes for any drive links or direct links
        uploaded_bytes_list = []
        for link in resume_links:
            if not link:
                uploaded_bytes_list.append(None)
                continue
            file_bytes = None
            if "drive.google.com" in link:
                file_bytes = download_drive_pdf(link)
            else:
                try:
                    r = requests.get(link, timeout=15)
                    if r.status_code == 200:
                        content = r.content or b""
                        ct = (r.headers.get("Content-Type") or "").lower()
                        if content.startswith(b"%PDF") or "application/pdf" in ct or link.lower().endswith(".pdf"):
                            file_bytes = content
                except Exception:
                    file_bytes = None
            uploaded_bytes_list.append(file_bytes)
    else:
        names = _split_multi(candidate_names)
        emails = _split_multi(candidate_emails)
        leet_users = _split_multi(leet_bulk)
        git_users = _split_multi(git_bulk)
        uploaded_bytes_list = [f.read() for f in uploaded_files] if 'uploaded_files' in locals() and uploaded_files else []
        resume_links = [""] * len(uploaded_bytes_list)  # maintain same length

    n = max(len(names), len(emails), len(leet_users), len(git_users), len(uploaded_bytes_list), len(resume_links))
    if n == 0:
        st.error("No candidates found. Enter at least one candidate or upload a resume.")
        st.stop()
    if n > 500:
        st.warning("Limiting to first 500 entries.")
        n = 500

    rows: List[Dict[str, Any]] = []
    st.info(f"Processing {n} candidates... this may take a while.")

    for i in range(n):
        name = names[i] if i < len(names) else ""
        email = emails[i] if i < len(emails) else ""
        leet = leet_users[i] if i < len(leet_users) and leet_users[i] else None
        git = git_users[i] if i < len(git_users) and git_users[i] else None
        file_bytes = uploaded_bytes_list[i] if i < len(uploaded_bytes_list) else None
        resume_url = resume_links[i] if i < len(resume_links) else None

        leet_res = fetch_leetcode(leet) if leet else {"skipped": True}
        git_res = fetch_github(git) if git else {"skipped": True}

        # IMPORTANT: pass both resume_url and file_bytes so ATS logic can try both
        ats_res = fetch_ats_from_url_or_bytes(resume_url, file_bytes)

        cand_score = compute_candidate_score(leet_res, git_res, ats_res)
        strengths, improvements = recruiter_insights(leet_res, git_res, ats_res)
        display_name = name or git or leet or f"Candidate {i+1}"

        # Render candidate header
        col_icon, col_name = st.columns([1, 10])
        with col_icon:
            st.markdown("🧑‍💻")
        with col_name:
            st.markdown(f"<h2 style='margin:0;padding:0'>{display_name}</h2>", unsafe_allow_html=True)

        # Layout metrics
        c_score, c_leet, c_git, c_ats = st.columns([1.3, 3, 3, 3])
        with c_score:
            st.subheader("Candidate Score")
            st.markdown(f"<h1 style='margin:0'>{cand_score}</h1>", unsafe_allow_html=True)
        with c_leet:
            st.markdown("#### LeetCode:")
            if leet_res.get("skipped"):
                st.write("—")
            elif leet_res.get("error"):
                st.write("LeetCode: Error")
                st.write(leet_res["error"])
            else:
                total = int(leet_res.get("total_solved") or 0)
                probs = leet_res.get("problems_by_difficulty", {})
                easy = int(probs.get("Easy") or 0)
                med = int(probs.get("Medium") or 0)
                hard = int(probs.get("Hard") or 0)
                st.write(f"LeetCode solved: **{total}** (Easy: {easy}, Medium: {med}, Hard: {hard})")
        with c_git:
            st.markdown("#### GitHub:")
            if git_res.get("skipped"):
                st.write("—")
            elif git_res.get("error"):
                st.write("GitHub: Error")
                st.write(git_res["error"])
            else:
                total_repos = int(git_res.get("total_repos") or 0)
                st.write(f"Repos: **{total_repos}**")
                st.write(f"Active months: {git_res.get('active_months', '—')}")
        with c_ats:
            st.markdown("#### ATS:")
            if file_bytes:
                st.write(f"Resume bytes: {len(file_bytes)} bytes")
            else:
                st.write("Resume bytes: — (no bytes downloaded/uploaded)")
            if resume_url:
                st.write("Resume link:", resume_url)
            # Display the raw ats_res for debugging
            #st.write("ATS result:", ats_res)
            if ats_res.get("skipped"):
                st.write("— (no resume provided)")
            elif ats_res.get("error"):
                st.write("ATS: Error")
                st.write(ats_res.get("error"))
            else:
                if ats_res.get("score") is None:
                    st.write("ATS score: — (could not extract text / scanned PDF)")
                    st.write(ats_res.get("remarks", ""))
                else:
                    st.write(f"ATS score: **{ats_res.get('score')}**")
                    if ats_res.get("remarks"):
                        st.write(ats_res.get("remarks"))

        st.markdown("**✅ Strengths**")
        if strengths:
            for s in strengths:
                st.write(f"- {s}")
        else:
            st.write("- —")

        st.markdown("**⚠️ Areas to Improve**")
        if improvements:
            for m in improvements:
                st.write(f"- {m}")
        else:
            st.write("- —")

        st.markdown("---")

        rows.append({
            "candidate_name": name,
            "candidate_email": email,
            "display_name": display_name,
            "leetcode.username": leet or "",
            "leetcode.total_solved": int(leet_res.get("total_solved") or 0) if leet_res and not leet_res.get("error") and not leet_res.get("skipped") else 0,
            "github.username": git or "",
            "github.total_repos": int(git_res.get("total_repos") or 0) if git_res and not git_res.get("error") and not git_res.get("skipped") else 0,
            "ats.score": ats_res.get("score") if ats_res and not ats_res.get("error") and not ats_res.get("skipped") else None,
            "ats.remarks": ats_res.get("remarks", "") if ats_res else "",
            "candidate_score": cand_score,
            "generated_at": time.strftime("%Y-%m-%d %H:%M:%S")
        })

    df = pd.DataFrame(rows)
    st.session_state["df"] = df

    excel_bytes = create_excel_report_bytes(df)
    st.download_button("📊 Download Excel Report", excel_bytes, "profile_report.xlsx")
    st.success(f"✅ Report generated in {time.time()-start_time:.1f}s")

# ---------------------------------------------
# ACTIONS AFTER GENERATION: EMAIL / PDF / CHART / SAVE
# ---------------------------------------------
if "df" in st.session_state:
    df = st.session_state["df"].copy()
    if "candidate_score" not in df.columns:
        df["candidate_score"] = 0
    df_top15 = df.fillna({"candidate_score": 0}).sort_values("candidate_score", ascending=False).head(15)

    st.subheader("Actions for Generated Report")
    colA, colB = st.columns([1, 2])
    with colA:
        sender_env = os.getenv("SENDER_EMAIL") or "ratnaakshithamanchikanti@gmail.com"
        pass_env = os.getenv("EMAIL_APP_PASSWORD") or "ugidsjvdrjsszunk"
        if not sender_env or not pass_env:
            st.warning("Email sender is not configured in environment. Set SENDER_EMAIL and EMAIL_APP_PASSWORD to enable sending.")
        if st.button("📧 Send Email to Top 15 Candidates"):
            try:
                result = send_bulk_emails_safe(df_top15, sender_env, pass_env)
                sent = result.get("sent", 0)
                failed = result.get("failed", 0)
                st.success(f"Emails sent: {sent}. Failed: {failed}.")
                if result.get("errors"):
                    for err in result["errors"]:
                        st.error(err)
            except Exception as e:
                st.error(f"Error sending emails: {e}")

    with colB:
        st.write("Download or visualize top candidates below.")

    st.subheader("📄 Download PDF Summary (Top candidates)")
    try:
        pdf_bytes = _create_pdf_bytes(df.to_dict(orient="records"), title="Profile Summary Report")
        st.download_button(label="Download PDF Report", data=pdf_bytes, file_name=f"profile_report_{int(time.time())}.pdf", mime="application/pdf")
    except Exception as e:
        st.error(f"Error generating PDF: {e}")

    st.subheader("📈 Comparative Overview (Top 15)")
    if not df_top15.empty:
        long_rows = []
        for _, r in df_top15.iterrows():
            long_rows.append({"display_name": r["display_name"], "metric": "LeetCode Solved", "value": int(r.get("leetcode.total_solved") or 0)})
            long_rows.append({"display_name": r["display_name"], "metric": "GitHub Repos", "value": int(r.get("github.total_repos") or 0)})
            long_rows.append({"display_name": r["display_name"], "metric": "ATS Score", "value": int(r.get("ats.score") or 0) if r.get("ats.score") not in (None, "") else 0})
            long_rows.append({"display_name": r["display_name"], "metric": "Candidate Score", "value": int(r.get("candidate_score") or 0)})
        df_long = pd.DataFrame(long_rows)
        try:
            fig = px.bar(df_long, x="display_name", y="value", color="metric", barmode="group", title="Comparative Metrics")
            st.plotly_chart(fig, use_container_width=True)
        except Exception:
            st.write("Unable to render chart.")
    else:
        st.write("No data to plot.")

    # Optional: Save report rows back to Google Sheet
    if use_google_sheet and sheet_url_input:
        st.markdown("### Save generated rows back to Google Sheet (optional)")
        if st.button("💾 Save report rows to Google Sheet"):
            df_to_save = df.fillna("").copy()
            header = list(df_to_save.columns)
            rows_to_append = [header]
            for _, r in df_to_save.iterrows():
                rows_to_append.append([str(r.get(c, "")) for c in header])
            success, msg = append_rows_to_sheet(sheet_url_input, rows_to_append)
            if success:
                st.success(msg)
            else:
                st.error(msg)

else:
    st.info("Generate a report to view charts or send emails.")
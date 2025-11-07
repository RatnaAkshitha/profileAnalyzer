import os
import time
import io
import re
import json
import uuid
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

# Try to import user's mail helper (akshi.send_bulk_emails). If not available, we'll fallback to SMTP sender below.
try:
    from akshi import send_bulk_emails  # type: ignore
except Exception:
    send_bulk_emails = None  # fallback later

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
        # Fallback: write CSV if xlsx writer not available
        buf = io.BytesIO()
        df.to_csv(buf, index=False)
        buf.seek(0)
        return buf.read()

# -----------------------------
# Setup
# -----------------------------
load_dotenv()
st.set_page_config(page_title="Profile & Resume Analyzer", layout="wide")
st.title("Profile & Resume Analyzer — LeetCode | GitHub | ATS")
st.write("You can paste names / emails / usernames OR upload PDF resumes. Max 100 entries.")

# -----------------------------
# Helpers
# -----------------------------
def _split_multi(value: str, max_items: int = 100) -> List[str]:
    if not value:
        return []
    parts: List[str] = []
    for line in value.splitlines():
        parts.extend([p.strip() for p in line.split(",")])
    parts = [p for p in parts if p]
    return parts[:max_items]

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
# ATS: robust Drive downloader + PDF extraction/ATS scoring
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

def download_drive_pdf(url: str) -> Optional[bytes]:
    if not url:
        return None
    try:
        s = str(url).strip()
        m = re.search(r"/d/([a-zA-Z0-9_-]+)", s)
        if not m:
            m = re.search(r"id=([a-zA-Z0-9_-]+)", s)
        if not m:
            return None
        file_id = m.group(1)

        session = requests.Session()
        headers = {"User-Agent": "Profile-Analyzer/1.0"}
        base = f"https://drive.google.com/uc?export=download&id={file_id}"

        r = session.get(base, headers=headers, timeout=20, stream=True)
        if r.status_code != 200:
            preview = f"https://drive.google.com/file/d/{file_id}/view"
            try:
                r = session.get(preview, headers=headers, timeout=20, stream=True)
            except Exception:
                return None
            if r.status_code != 200:
                return None

        content = r.content or b""
        ct = (r.headers.get("Content-Type") or "").lower()

        if content.startswith(b"%PDF") or "application/pdf" in ct:
            return content

        text = ""
        try:
            text = r.text or ""
        except Exception:
            text = ""

        token = None
        m1 = re.search(r"confirm=([0-9A-Za-z-_]+)&", text)
        if m1:
            token = m1.group(1)
        if not token:
            m2 = re.search(r'name="confirm"\s+value="([0-9A-Za-z-_]+)"', text)
            if m2:
                token = m2.group(1)
        if not token:
            m3 = re.search(r'href="[^"]*?confirm=([0-9A-Za-z-_]+)[^"]*"', text)
            if m3:
                token = m3.group(1)
        if not token:
            for cookie_name, cookie in session.cookies.items():
                if cookie_name.startswith("download_warning"):
                    token = cookie.value
                    break

        if token:
            confirm_url = f"https://drive.google.com/uc?export=download&confirm={token}&id={file_id}"
            r2 = session.get(confirm_url, headers=headers, timeout=30, stream=True)
            if r2.status_code == 200:
                content2 = r2.content or b""
                ct2 = (r2.headers.get("Content-Type") or "").lower()
                if content2.startswith(b"%PDF") or "application/pdf" in ct2:
                    return content2
                idx = content2.find(b"%PDF")
                if idx != -1:
                    return content2[idx:]
            return None

        idx = content.find(b"%PDF")
        if idx != -1:
            return content[idx:]

        if "docs.google.com/document" in s or "docs.google.com/spreadsheets" in s or "docs.google.com/presentation" in s:
            export_url = None
            if "document" in s:
                export_url = f"https://docs.google.com/document/d/{file_id}/export?format=pdf"
            elif "spreadsheets" in s:
                export_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=pdf"
            elif "presentation" in s:
                export_url = f"https://docs.google.com/presentation/d/{file_id}/export/pdf"
            if export_url:
                r3 = session.get(export_url, headers=headers, timeout=30)
                if r3.status_code == 200:
                    c3 = r3.content or b""
                    if c3.startswith(b"%PDF") or "application/pdf" in (r3.headers.get("Content-Type") or "").lower():
                        return c3

        return None
    except Exception:
        return None

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

def fetch_ats_from_url_or_bytes(resume_input: Optional[str], file_bytes: Optional[bytes]) -> Dict[str, Any]:
    if file_bytes:
        try:
            return get_ats_score_from_pdf_bytes(file_bytes)
        except Exception as e:
            return {"error": f"ATS PDF error: {e}"}

    if resume_input:
        if "drive.google.com" not in resume_input.lower():
            return {
                "skipped": True,
                "remarks": "Non-Google-Drive resume URLs are ignored. Please upload a PDF or provide a public Google Drive link (drive.google.com)."
            }
        file_bytes_dl = download_drive_pdf(resume_input)
        if not file_bytes_dl:
            return {"error": "Could not download resume from Google Drive (public access required or unsupported URL format)."}
        try:
            return get_ats_score_from_pdf_bytes(file_bytes_dl)
        except Exception as e:
            return {"error": f"ATS PDF error after download: {e}"}
    return {"skipped": True}

# ---------------------------------------------
# Email sending fallback (to avoid undefined names)
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
    """
    Uses akshi.send_bulk_emails if available, otherwise tries fallback.
    Returns summary dict with sent/failed/errors.
    """
    if send_bulk_emails:
        try:
            send_bulk_emails(df, sender, app_password)
            return {"sent": int(len(df)), "failed": 0, "errors": []}
        except Exception as e:
            try:
                return send_bulk_emails_fallback(df, sender or "", app_password or "")
            except Exception as e2:
                raise RuntimeError(f"External sender failed: {e}; fallback failed: {e2}")
    else:
        return send_bulk_emails_fallback(df, sender or "", app_password or "")

# ---------------------------------------------
# Scoring & Insights
# ---------------------------------------------
def compute_candidate_score(leet_res: Dict[str, Any], git_res: Dict[str, Any], ats_res: Dict[str, Any]) -> int:
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

def recruiter_insights(leet_res: Dict[str, Any], git_res: Dict[str, Any], ats_res: Dict[str, Any]) -> Tuple[List[str], List[str]]:
    strengths: List[str] = []
    improvements: List[str] = []
    try:
        solved = int(leet_res.get("total_solved", 0)) if leet_res and not leet_res.get("error") else 0
        if solved >= 300:
            strengths.append("Strong problem-solving skills (LeetCode).")
        elif solved >= 100:
            strengths.append("Good problem-solving foundation on LeetCode.")
        else:
            improvements.append("Solve more DSA problems on LeetCode to strengthen fundamentals.")
    except Exception:
        improvements.append("LeetCode data incomplete or unavailable.")
    try:
        probs = leet_res.get("problems_by_difficulty", {}) if leet_res else {}
        hard = int(probs.get("Hard", 0) or 0)
        med = int(probs.get("Medium", 0) or 0)
        if hard >= 10:
            strengths.append("Has experience solving Hard problems.")
        elif med >= 20:
            strengths.append("Comfortable with Medium-level problems.")
    except Exception:
        pass
    try:
        langs = leet_res.get("languages", {}) if leet_res else {}
        top_langs = sorted(langs.items(), key=lambda x: x[1], reverse=True)[:3]
        if top_langs:
            strengths.append("Top languages: " + ", ".join([f"{ln} ({cnt})" for ln, cnt in top_langs]))
    except Exception:
        pass
    try:
        repos = int(git_res.get("total_repos", 0)) if git_res and not git_res.get("error") else 0
        if repos > 5:
            strengths.append("Active GitHub with multiple projects.")
        else:
            improvements.append("Add more meaningful GitHub projects (with README & tests).")
        active_months = git_res.get("active_months", 0) if git_res else 0
        if active_months and active_months >= 6:
            strengths.append(f"Consistent GitHub activity (~{active_months} months).")
    except Exception:
        pass
    try:
        ats = int(ats_res.get("score", 0)) if ats_res and not ats_res.get("error") else 0
        if ats >= 75:
            strengths.append("Resume is ATS-friendly.")
        else:
            improvements.append("Optimize resume for ATS (keywords, layout, sectioning).")
    except Exception:
        pass
    strengths = list(dict.fromkeys(strengths))
    improvements = list(dict.fromkeys(improvements))
    return strengths, improvements

# ---------------------------------------------
# Charts / PDF generation
# ---------------------------------------------
def _save_matplotlib_bar(df_long: pd.DataFrame) -> io.BytesIO:
    try:
        plt.style.use("seaborn-v0_8-colorblind")
    except Exception:
        plt.style.use("default")
    fig, ax = plt.subplots(figsize=(10, 6))
    pivot = df_long.pivot(index="display_name", columns="metric", values="value").fillna(0)
    pivot.plot(kind="bar", ax=ax)
    ax.set_ylabel("Value")
    ax.set_xlabel("")
    ax.legend(title="Metric", bbox_to_anchor=(1.02, 1), loc="upper left")
    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150)
    plt.close(fig)
    buf.seek(0)
    return buf

def _create_pdf_bytes(rows: List[Dict[str, Any]], title: str = "Profile Report") -> bytes:
    df = pd.DataFrame(rows).fillna("")
    top = df.sort_values("candidate_score", ascending=False).head(10)
    long_rows: List[Dict[str, Any]] = []
    for _, r in top.iterrows():
        long_rows.append({"display_name": r.get("display_name", ""), "metric": "LeetCode solved", "value": int(r.get("leetcode.total_solved") or 0)})
        long_rows.append({"display_name": r.get("display_name", ""), "metric": "GitHub repos", "value": int(r.get("github.total_repos") or 0)})
        long_rows.append({"display_name": r.get("display_name", ""), "metric": "ATS score", "value": int(r.get("ats.score") or 0)})
        long_rows.append({"display_name": r.get("display_name", ""), "metric": "Candidate score", "value": int(r.get("candidate_score") or 0)})
    df_long = pd.DataFrame(long_rows)
    chart_buf = _save_matplotlib_bar(df_long)
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    story: List[Any] = []
    story.append(Paragraph(title, styles["Title"]))
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"Generated at: {time.strftime('%Y-%m-%d %H:%M:%S')}", styles["Normal"]))
    story.append(Spacer(1, 12))
    try:
        img = utils.ImageReader(chart_buf)
        iw, ih = img.getSize()
        aspect = ih / float(iw)
        width = 450
        height = width * aspect
        story.append(RLImage(chart_buf, width=width, height=height))
    except Exception:
        story.append(Paragraph("Chart not available", styles["Normal"]))
    story.append(Spacer(1, 12))
    table_data = [["Name", "Email", "LeetCode", "GitHub", "ATS", "Score"]]
    for _, r in top.iterrows():
        table_data.append([
            r.get("candidate_name", r.get("display_name", "")),
            r.get("candidate_email", ""),
            int(r.get("leetcode.total_solved") or 0),
            int(r.get("github.total_repos") or 0),
            int(r.get("ats.score") or 0),
            int(r.get("candidate_score") or 0),
        ])
    tbl = Table(table_data, hAlign="LEFT")
    story.append(tbl)
    doc.build(story)
    buffer.seek(0)
    return buffer.read()

def create_pdf(rows: List[Dict[str, Any]], title: str = "Profile Report") -> bytes:
    return _create_pdf_bytes(rows, title=title)

# ---------------------------------------------
# UI Inputs (names, emails, usernames, file uploads, resume URL)
# ---------------------------------------------
left, right = st.columns(2)
with left:
    candidate_names = st.text_area(
        "Candidate Names (one per line or comma-separated)", height=180,
        placeholder="e.g. John Doe\nJane Smith"
    )
    leet_bulk = st.text_area(
        "LeetCode usernames (one per line or comma-separated)", height=180,
        placeholder="e.g. tourist\nlabuladong"
    )
with right:
    candidate_emails = st.text_area(
        "Candidate Emails (one per line or comma-separated)", height=180,
        placeholder="e.g. john@example.com\njane@example.com"
    )
    git_bulk = st.text_area(
        "GitHub usernames (one per line or comma-separated)", height=180,
        placeholder="e.g. torvalds\noctocat"
    )

st.markdown("----")
st.write("Upload resumes as PDF files (optional). If you upload files, they will be matched by index to names/emails/usernames. Max 100 files.")
uploaded_files = st.file_uploader("Upload PDF Resumes (multiple)", type=["pdf"], accept_multiple_files=True)

st.markdown("Optional: Paste resume URLs (one per line, matched by index). Only Google Drive links (drive.google.com) will be processed for ATS scoring; other URLs will be ignored.")
resume_urls_input = st.text_area("Resume URLs (one per line, optional)", height=120, placeholder="https://drive.google.com/file/d/FILE_ID/view")

# -----------------------------
# Generate Report button
# -----------------------------
if st.button("Generate Report", type="primary"):
    start_time = time.time()

    names = _split_multi(candidate_names)
    emails = _split_multi(candidate_emails)
    leet_users = _split_multi(leet_bulk)
    git_users = _split_multi(git_bulk)
    resume_urls = _split_multi(resume_urls_input)

    uploaded_bytes_list: List[Optional[bytes]] = []
    if uploaded_files:
        for f in uploaded_files[:100]:
            try:
                uploaded_bytes_list.append(f.read())
            except Exception:
                uploaded_bytes_list.append(None)

    n = max(len(names), len(emails), len(leet_users), len(git_users), len(uploaded_bytes_list), len(resume_urls))
    if n == 0:
        st.error("Please enter at least one entry or upload at least one resume.")
        st.stop()
    if n > 100:
        st.warning("Processing only the first 100 entries.")
        n = 100

    st.info(f"Processing {n} candidate(s)... (this might take a while)")
    rows: List[Dict[str, Any]] = []

    for i in range(n):
        name = names[i] if i < len(names) else ""
        email = emails[i] if i < len(emails) else ""
        leet = leet_users[i] if i < len(leet_users) else None
        git = git_users[i] if i < len(git_users) else None
        file_bytes = uploaded_bytes_list[i] if i < len(uploaded_bytes_list) else None
        resume_url = resume_urls[i] if i < len(resume_urls) else None

        leet_res = fetch_leetcode(leet) if leet else {"skipped": True}
        git_res = fetch_github(git) if git else {"skipped": True}
        ats_res = fetch_ats_from_url_or_bytes(resume_url, file_bytes)

        cand_score = compute_candidate_score(leet_res, git_res, ats_res)
        strengths, improvements = recruiter_insights(leet_res, git_res, ats_res)

        display_name = name or git or leet or f"Candidate {i+1}"

        st.markdown(f"## 👤 {display_name}")
        c1, c2, c3, c4 = st.columns(4)

        with c1:
            st.metric("Candidate Score", cand_score)

        with c2:
            st.write("**LeetCode:**")
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
                badges = leet_res.get("badges", []) or []
                if badges:
                    st.write("Badges: " + ", ".join(badges[:8]))
                else:
                    st.write("Badges: —")
                langs = leet_res.get("languages", {}) or {}
                if langs:
                    top_langs = sorted(langs.items(), key=lambda x: x[1], reverse=True)[:5]
                    st.write("Languages: " + ", ".join([f"{ln} ({cnt})" for ln, cnt in top_langs]))
                else:
                    st.write("Languages: —")

        with c3:
            st.write("**GitHub:**")
            if git_res.get("skipped"):
                st.write("—")
            elif git_res.get("error"):
                st.write("GitHub: Error")
                st.write(git_res["error"])
            else:
                total_repos = int(git_res.get("total_repos") or 0)
                st.write(f"Repos: **{total_repos}**")
                st.write(f"Active months: {git_res.get('active_months', '—')}")
                st.write(f"Recent push events: {git_res.get('recent_push_events', '—')}")

        with c4:
            st.write("**ATS:**")
            if ats_res.get("skipped"):
                st.write("— (no resume provided or non-Drive URL ignored)")
                if ats_res.get("remarks"):
                    st.write(ats_res.get("remarks"))
            elif ats_res.get("error"):
                st.write("ATS: Error")
                st.write(ats_res.get("error"))
            else:
                st.write(f"ATS score: **{ats_res.get('score', '—')}**")
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
            "leetcode.total_solved": int(leet_res.get("total_solved") or 0) if leet_res and not leet_res.get("error") and not leet_res.get("skipped") else None,
            "leetcode.easy": int(leet_res.get("problems_by_difficulty", {}).get("Easy") or 0) if leet_res and not leet_res.get("error") else None,
            "leetcode.medium": int(leet_res.get("problems_by_difficulty", {}).get("Medium") or 0) if leet_res and not leet_res.get("error") else None,
            "leetcode.hard": int(leet_res.get("problems_by_difficulty", {}).get("Hard") or 0) if leet_res and not leet_res.get("error") else None,
            "leetcode.badges": ", ".join(leet_res.get("badges", [])) if leet_res and not leet_res.get("error") else None,
            "leetcode.languages": ", ".join([f"{ln} ({cnt})" for ln, cnt in (leet_res.get("languages") or {}).items()]) if leet_res and not leet_res.get("error") else None,
            "github.username": git or "",
            "github.total_repos": int(git_res.get("total_repos") or 0) if git_res and not git_res.get("error") else None,
            "github.active_months": git_res.get("active_months", None) if git_res and not git_res.get("error") else None,
            "github.recent_push_events": git_res.get("recent_push_events", None) if git_res and not git_res.get("error") else None,
            "ats.url": resume_url or (f"uploaded_pdf_{i+1}" if file_bytes else ""),
            "ats.score": int(ats_res.get("score") or 0) if ats_res and not ats_res.get("error") and not ats_res.get("skipped") else None,
            "ats.remarks": ats_res.get("remarks", "") if ats_res else None,
            "candidate_score": cand_score,
            "generated_at": time.strftime("%Y-%m-%d %H:%M:%S")
        })

    df = pd.DataFrame(rows)
    st.session_state["df"] = df

    st.subheader("📥 Download Excel Report (All Candidates)")
    excel_bytes_all = create_excel_report_bytes(df)
    st.download_button(
        label="Download Excel (All)",
        data=excel_bytes_all,
        file_name=f"profile_report_all_{int(time.time())}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_excel_all_generate"
    )

    st.success(f"Report generated successfully! Done in {time.time() - start_time:.1f}s")

# -----------------------------
# Email & Visualization (uses session_state)
# -----------------------------
if "df" in st.session_state:
    df = st.session_state["df"].copy()
    if "candidate_score" not in df.columns:
        df["candidate_score"] = 0
    df_top15 = df.fillna(0).sort_values("candidate_score", ascending=False).head(15)

    st.subheader("Actions for Generated Report")

    colA, colB = st.columns([1, 2])
    with colA:
        if st.button("📧 Send Email to Top 15 Candidates"):
            sender = os.getenv("SENDER_EMAIL") or "ratnaakshithamanchikanti"
            app_password = os.getenv("EMAIL_APP_PASSWORD") or "ugidsjvdrjsszunk"
            try:
                result = send_bulk_emails_safe(df_top15, sender, app_password)
                sent = result.get("sent", 0)
                failed = result.get("failed", 0)
                st.success(f"Emails sent: {sent}. Failed: {failed}.")
                for err in result.get("errors", []):
                    st.error(err)
            except Exception as e:
                st.error(f"Error sending emails: {e}")

    with colB:
        st.write("Download or visualize top candidates below.")

    st.subheader("📄 Download PDF Summary (Top candidates)")
    try:
        pdf_bytes = create_pdf(df.to_dict(orient="records"), title="Profile Summary Report")
        st.download_button(
            label="Download PDF Report",
            data=pdf_bytes,
            file_name=f"profile_report_{int(time.time())}.pdf",
            mime="application/pdf",
            key="download_pdf_top"
        )
    except Exception as e:
        st.error(f"Error generating PDF: {e}")

    # New: Excel downloads in Actions area (Top 15 and All)
    st.subheader("📥 Download Excel Reports")
    try:
        excel_top15 = create_excel_report_bytes(df_top15)
        st.download_button(
            label="Download Excel (Top 15)",
            data=excel_top15,
            file_name=f"profile_report_top15_{int(time.time())}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excel_top15_actions"
        )
        
    except Exception as e:
        st.error(f"Error preparing Excel files: {e}")

    st.subheader("📈 Comparative Overview (Top 15)")
    if not df_top15.empty:
        long_rows = []
        for _, r in df_top15.iterrows():
            long_rows.append({"display_name": r["display_name"], "metric": "LeetCode Solved", "value": int(r.get("leetcode.total_solved") or 0)})
            long_rows.append({"display_name": r["display_name"], "metric": "GitHub Repos", "value": int(r.get("github.total_repos") or 0)})
            long_rows.append({"display_name": r["display_name"], "metric": "ATS Score", "value": int(r.get("ats.score") or 0)})
            long_rows.append({"display_name": r["display_name"], "metric": "Candidate Score", "value": int(r.get("candidate_score") or 0)})
        df_long = pd.DataFrame(long_rows)
        try:
            fig = px.bar(df_long, x="display_name", y="value", color="metric", barmode="group", title="Comparative Metrics")
            st.plotly_chart(fig, use_container_width=True)
        except Exception:
            st.write("Unable to render chart.")
    else:
        st.write("No data to plot.")
else:
    st.info("Please generate a report first (enter names/emails/usernames or upload PDFs) to enable actions.")

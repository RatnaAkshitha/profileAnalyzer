# app.py
import os
import time
import io
from typing import List, Dict, Any
import pandas as pd
import streamlit as st
import plotly.express as px
import matplotlib.pyplot as plt
from dotenv import load_dotenv
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import utils

# services (should exist in services/ and return expected keys)
from services.leetcode_service import get_leetcode_data
from services.github_service import get_github_data
from services.ats_service import get_ats_score
from utils.excel_generator import create_excel_report_bytes

# -----------------------------
# Setup
# -----------------------------
load_dotenv()
st.set_page_config(page_title="Profile & Resume Analyzer", layout="wide")
st.title("Profile & Resume Analyzer — LeetCode | GitHub | ATS")
st.write("Paste up to **100** entries in each box (one per line or comma-separated). Resume URLs optional (matched by index).")

# -----------------------------
# Helpers
# -----------------------------
def _split_multi(value: str, max_items: int = 100) -> List[str]:
    if not value:
        return []
    parts = []
    for line in value.splitlines():
        parts.extend([p.strip() for p in line.split(",")])
    parts = [p for p in parts if p]
    return parts[:max_items]

@st.cache_data(ttl=300)
def fetch_leetcode(username: str) -> Dict[str, Any]:
    return get_leetcode_data(username)

@st.cache_data(ttl=300)
def fetch_github(username: str) -> Dict[str, Any]:
    return get_github_data(username)

@st.cache_data(ttl=300)
def fetch_ats(url: str) -> Dict[str, Any]:
    return get_ats_score(url)

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
            ats_score = int(ats_res.get("score", 0) or 0)
            score += min(30, ats_score)
    except Exception:
        pass
    return int(score)

def recruiter_insights(leet_res, git_res, ats_res):
    strengths = []
    improvements = []

    # LeetCode summary
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

    # Difficulty breakdown
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

    # Languages
    try:
        langs = leet_res.get("languages", {}) if leet_res else {}
        top_langs = sorted(langs.items(), key=lambda x: x[1], reverse=True)[:3]
        if top_langs:
            strengths.append("Top languages: " + ", ".join([f"{ln} ({cnt})" for ln, cnt in top_langs]))
    except Exception:
        pass

    # GitHub
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

    # ATS
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

def _save_matplotlib_bar(df_long):
    try:
        plt.style.use("seaborn-v0_8-colorblind")
    except OSError:
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

def _create_pdf_bytes(rows: List[Dict[str, Any]], title="Profile Report"):
    df = pd.DataFrame(rows).fillna("")
    top = df.sort_values("candidate_score", ascending=False).head(10)

    long_rows = []
    for _, r in top.iterrows():
        long_rows.append({"display_name": r["display_name"], "metric": "LeetCode solved", "value": int(r.get("leetcode.total_solved") or 0)})
        long_rows.append({"display_name": r["display_name"], "metric": "GitHub repos", "value": int(r.get("github.total_repos") or 0)})
        long_rows.append({"display_name": r["display_name"], "metric": "ATS score", "value": int(r.get("ats.score") or 0)})
        long_rows.append({"display_name": r["display_name"], "metric": "Candidate score", "value": int(r.get("candidate_score") or 0)})
    df_long = pd.DataFrame(long_rows)

    chart_buf = _save_matplotlib_bar(df_long)

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []

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

# -----------------------------
# UI Inputs
# -----------------------------
left, right = st.columns(2)
with left:
    candidate_names = st.text_area("Candidate Names (one per line or comma-separated)", height=180,
                                   placeholder="e.g. John Doe\nJane Smith")
    leet_bulk = st.text_area("LeetCode usernames (one per line or comma-separated)", height=180,
                             placeholder="e.g. tourist\nlabuladong")
with right:
    candidate_emails = st.text_area("Candidate Emails (one per line or comma-separated)", height=180,
                                    placeholder="e.g. john@example.com\njane@example.com")
    git_bulk = st.text_area("GitHub usernames (one per line or comma-separated)", height=180,
                            placeholder="e.g. torvalds\noctocat")
    resume_bulk = st.text_area("Resume URLs (optional, one per line; matched by index)", height=180,
                               placeholder="https://example.com/resume1.pdf")

if st.button("Generate Report", type="primary"):
    start_time = time.time()
    names = _split_multi(candidate_names)
    emails = _split_multi(candidate_emails)
    leet_users = _split_multi(leet_bulk)
    git_users = _split_multi(git_bulk)
    resumes = _split_multi(resume_bulk)

    if not any([names, emails, leet_users, git_users, resumes]):
        st.error("Please enter at least one entry in any field.")
        st.stop()

    n = max(len(names), len(emails), len(leet_users), len(git_users), len(resumes))
    if n == 0:
        st.error("No entries found after parsing.")
        st.stop()
    if n > 100:
        st.warning("Processing only the first 100 entries.")
        n = 100

    st.info(f"Processing {n} candidate(s)... (this might take a while)")

    rows = []
    for i in range(n):
        name = names[i] if i < len(names) else ""
        email = emails[i] if i < len(emails) else ""
        leet = leet_users[i] if i < len(leet_users) else None
        git = git_users[i] if i < len(git_users) else None
        resu = resumes[i] if i < len(resumes) else None

        leet_res = fetch_leetcode(leet) if leet else {"skipped": True}
        git_res = fetch_github(git) if git else {"skipped": True}
        ats_res = fetch_ats(resu) if resu else {"skipped": True}

        cand_score = compute_candidate_score(leet_res, git_res, ats_res)
        strengths, improvements = recruiter_insights(leet_res, git_res, ats_res)

        display_name = name or git or leet or f"Candidate {i+1}"
        st.markdown(f"## 👤 {display_name}")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("Candidate Score", cand_score)
        with c2:
            if leet_res.get("skipped"):
                st.write("LeetCode: —")
            elif leet_res.get("error"):
                st.write("LeetCode: Error")
                st.write(leet_res["error"])
            else:
                total = leet_res.get("total_solved", "—")
                probs = leet_res.get("problems_by_difficulty", {})
                easy = probs.get("Easy", 0)
                med = probs.get("Medium", 0)
                hard = probs.get("Hard", 0)
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
            if git_res.get("skipped"):
                st.write("GitHub: —")
            elif git_res.get("error"):
                st.write("GitHub: Error")
                st.write(git_res["error"])
            else:
                st.write(f"Repos: **{git_res.get('total_repos', '—')}**")
                st.write(f"Active months: {git_res.get('active_months', '—')}")
                st.write(f"Recent push events: {git_res.get('recent_push_events', '—')}")
        with c4:
            if ats_res.get("skipped"):
                st.write("ATS: —")
            elif ats_res.get("error"):
                st.write("ATS: Error")
                st.write(ats_res["error"])
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
            "leetcode.total_solved": int(leet_res.get("total_solved")) if leet_res and not leet_res.get("error") and leet_res.get("total_solved") else None,
            "leetcode.easy": int(leet_res.get("problems_by_difficulty", {}).get("Easy", 0)) if leet_res and not leet_res.get("error") else None,
            "leetcode.medium": int(leet_res.get("problems_by_difficulty", {}).get("Medium", 0)) if leet_res and not leet_res.get("error") else None,
            "leetcode.hard": int(leet_res.get("problems_by_difficulty", {}).get("Hard", 0)) if leet_res and not leet_res.get("error") else None,
            "leetcode.badges": ", ".join(leet_res.get("badges", [])) if leet_res and not leet_res.get("error") else None,
            "leetcode.languages": ", ".join([f"{ln} ({cnt})" for ln, cnt in (leet_res.get("languages") or {}).items()]) if leet_res and not leet_res.get("error") else None,
            "github.username": git or "",
            "github.total_repos": int(git_res.get("total_repos")) if git_res and not git_res.get("error") and git_res.get("total_repos") else None,
            "ats.url": resu or "",
            "ats.score": int(ats_res.get("score")) if ats_res and not ats_res.get("error") and ats_res.get("score") else None,
            "candidate_score": cand_score,
            "generated_at": time.strftime("%Y-%m-%d %H:%M:%S")
        })

    # Excel download
    st.subheader("📥 Download Excel Report")
    df = pd.DataFrame(rows)
    excel_bytes = create_excel_report_bytes(df)
    st.download_button(
        label="Download Excel Report",
        data=excel_bytes,
        file_name=f"profile_report_{int(time.time())}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # PDF download
    st.subheader("📄 Download PDF Summary (top candidates)")
    pdf_bytes = _create_pdf_bytes(rows, title="Profile Summary Report")
    st.download_button(
        label="Download PDF Report",
        data=pdf_bytes,
        file_name=f"profile_report_{int(time.time())}.pdf",
        mime="application/pdf"
    )

    # Comparative chart (top 15)
    st.subheader("📈 Comparative Overview (Top 15)")
    if rows:
        df_plot = pd.DataFrame(rows).fillna(0).sort_values("candidate_score", ascending=False).head(15)
        long_rows = []
        for _, r in df_plot.iterrows():
            long_rows.append({"display_name": r["display_name"], "metric": "LeetCode Solved", "value": int(r.get("leetcode.total_solved") or 0)})
            long_rows.append({"display_name": r["display_name"], "metric": "GitHub Repos", "value": int(r.get("github.total_repos") or 0)})
            long_rows.append({"display_name": r["display_name"], "metric": "ATS Score", "value": int(r.get("ats.score") or 0)})
            long_rows.append({"display_name": r["display_name"], "metric": "Candidate Score", "value": int(r.get("candidate_score") or 0)})

        df_long = pd.DataFrame(long_rows)
        fig = px.bar(df_long, x="display_name", y="value", color="metric", barmode="group", title="Comparative Metrics")
        st.plotly_chart(fig, use_container_width=True)

    st.success(f"Done in {time.time() - start_time:.1f}s")

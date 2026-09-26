"""Hackathon Team Screener — single-file Streamlit app.

Reads a Google Forms CSV/XLSX export (one row = one team of 5), runs rule-based
eligibility checks, fetches resume/portfolio links, scores eligible teams with
Claude against a fixed rubric, and produces a ranked shortlist for the review
panel.
"""

import base64
import hashlib
import io
import json
import os
import re
import time
from collections import Counter
from urllib.parse import urlparse

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

try:
    import anthropic
except ImportError:
    anthropic = None

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

try:
    from docx import Document as DocxDocument
except ImportError:
    DocxDocument = None


# ---------------------------------------------------------------------------
# Config / constants
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Hackathon Team Screener", layout="wide")

MODEL_NAME = "claude-sonnet-5"
MAX_TOKENS = 1000
REQUEST_TIMEOUT = 15
MIN_READABLE_CHARS = 200
RESUME_TRUNCATE_CHARS = 6000
PORTFOLIO_TRUNCATE_CHARS = 8000

CACHE_DIR = ".cache"
LINK_CACHE_PATH = os.path.join(CACHE_DIR, "link_cache.json")
SCORE_CACHE_PATH = os.path.join(CACHE_DIR, "score_cache.json")

SKILL_CATEGORIES = ["Developer", "Product Lead", "Marketing/Pitch", "Hybrid"]
IDEAL_COMPOSITION = {"Developer": 2, "Product Lead": 1, "Marketing/Pitch": 1, "Hybrid": 1}

SCORE_MAXES = {
    "technical_capability": 25,
    "portfolio_quality": 25,
    "problem_thinking": 20,
    "execution_track_record": 15,
}
SCORE_LABELS = {
    "technical_capability": "Technical Capability",
    "portfolio_quality": "Portfolio Quality",
    "problem_thinking": "Problem Thinking",
    "execution_track_record": "Execution Track Record",
}

MEMBER_REQUIRED_FIELDS = [
    "Full Name",
    "Email",
    "School and Program",
    "Year",
    "Skill Set",
    "Resume Link",
]
# Present per member and used for fetching, but an empty value never triggers a hard
# eligibility flag (matches the spec: portfolio issues are softer than resume issues).
MEMBER_SOFT_FIELDS = ["Portfolio or Case Study Link"]
MEMBER_OPTIONAL_FIELDS = ["LinkedIn or GitHub"]

BASE_REQUIRED_COLUMNS = ["Team Name", "Captain Full Name", "Captain Email"]
OTHER_REQUIRED_COLUMNS = [
    "All 5 members can attend the full event",
    "AI-assisted review consent",
]
Q_PREFIXES = ["Q1:", "Q2:", "Q3:"]

TOP_CUTOFF = 30
WAITLIST_CUTOFF = 35
REVIEW_BAND = (25, 40)

DRIVE_FILE_RE = re.compile(r"drive\.google\.com/file/d/([a-zA-Z0-9_-]+)")
DRIVE_OPEN_RE = re.compile(r"drive\.google\.com/open\?id=([a-zA-Z0-9_-]+)")
DOCS_RE = re.compile(r"docs\.google\.com/document/d/([a-zA-Z0-9_-]+)")
GITHUB_REPO_RE = re.compile(r"github\.com/([\w.-]+)/([\w.-]+?)(?:/.*)?$")


def member_field(i, field):
    return f"Member {i} {field}"


def build_required_columns():
    cols = list(BASE_REQUIRED_COLUMNS)
    for i in range(1, 6):
        for f in MEMBER_REQUIRED_FIELDS + MEMBER_SOFT_FIELDS:
            cols.append(member_field(i, f))
    cols += OTHER_REQUIRED_COLUMNS
    return cols


# ---------------------------------------------------------------------------
# Small utilities
# ---------------------------------------------------------------------------

def get_val(row, col):
    if col not in row:
        return ""
    v = row[col]
    if pd.isna(v):
        return ""
    return str(v).strip()


def normalize_skill(value):
    if not value:
        return None
    v = re.sub(r"[\s\-/]+", " ", value.strip().lower())
    mapping = {
        "developer": "Developer",
        "product lead": "Product Lead",
        "marketing pitch": "Marketing/Pitch",
        "marketing": "Marketing/Pitch",
        "pitch": "Marketing/Pitch",
        "hybrid": "Hybrid",
    }
    return mapping.get(v)


def looks_like_html_login(text, extracted_len):
    """Detect a short 'sign in to view this file' page, without flagging a real,
    content-rich page just because it happens to have a "Log in" nav link."""
    if extracted_len >= 1000:
        return False
    lowered = text.lower()
    if "<html" not in lowered and "<!doctype html" not in lowered:
        return False
    signals = ["sign in to continue", "accounts.google.com", "request access",
                "you need access", "you need permission"]
    return any(s in lowered for s in signals)


def load_json_cache(path):
    if os.path.exists(path):
        try:
            with open(path, "r") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_json_cache(path, data):
    os.makedirs(CACHE_DIR, exist_ok=True)
    with open(path, "w") as f:
        json.dump(data, f)


def redact(text, values):
    if not text:
        return text
    for value in values:
        if value and len(value.strip()) > 1:
            text = re.sub(re.escape(value.strip()), "[REDACTED]", text, flags=re.IGNORECASE)
    return text


# ---------------------------------------------------------------------------
# Column validation
# ---------------------------------------------------------------------------

def validate_columns(df):
    missing = [c for c in build_required_columns() if c not in df.columns]
    for prefix in Q_PREFIXES:
        if not any(str(c).startswith(prefix) for c in df.columns):
            missing.append(f'A column starting with "{prefix}"')
    return missing


def find_q_columns(df):
    found = {}
    for prefix in Q_PREFIXES:
        for c in df.columns:
            if str(c).startswith(prefix):
                found[prefix] = c
                break
    return found


# ---------------------------------------------------------------------------
# Stage 1: eligibility
# ---------------------------------------------------------------------------

def build_email_team_map(df):
    """Map each lowercase email -> set of team names that contain it."""
    mapping = {}
    for _, row in df.iterrows():
        team = get_val(row, "Team Name")
        emails = set()
        cap = get_val(row, "Captain Email").lower()
        if cap:
            emails.add(cap)
        for i in range(1, 6):
            e = get_val(row, member_field(i, "Email")).lower()
            if e:
                emails.add(e)
        for e in emails:
            mapping.setdefault(e, set()).add(team)
    return mapping


def check_structural_eligibility(row, email_team_map):
    reasons = []
    team = get_val(row, "Team Name")

    if not get_val(row, "Captain Full Name"):
        reasons.append("Captain Full Name is empty")
    captain_email = get_val(row, "Captain Email")
    if not captain_email:
        reasons.append("Captain Email is empty")
    elif not captain_email.lower().endswith("@nyu.edu"):
        reasons.append(f"Captain Email ({captain_email}) is not an @nyu.edu address")

    for i in range(1, 6):
        for f in MEMBER_REQUIRED_FIELDS:
            val = get_val(row, member_field(i, f))
            if not val:
                reasons.append(f"Member {i} {f} is empty")
        email = get_val(row, member_field(i, "Email"))
        if email and not email.lower().endswith("@nyu.edu"):
            reasons.append(f"Member {i} Email ({email}) is not an @nyu.edu address")

    # Cross-team duplicate emails (captain == member 1 within the same team is fine)
    seen_here = set()
    for e in [captain_email] + [get_val(row, member_field(i, "Email")) for i in range(1, 6)]:
        e_low = e.lower()
        if e_low and e_low not in seen_here:
            seen_here.add(e_low)
            other_teams = email_team_map.get(e_low, set()) - {team}
            if other_teams:
                reasons.append(f"Email {e} also appears on team(s): {', '.join(sorted(other_teams))}")

    attend = get_val(row, "All 5 members can attend the full event")
    if attend.lower() != "yes":
        reasons.append('"All 5 members can attend the full event" is not checked "Yes"')

    consent = get_val(row, "AI-assisted review consent")
    if consent.lower() != "yes":
        reasons.append('"AI-assisted review consent" is not checked "Yes"')

    return reasons


# ---------------------------------------------------------------------------
# Stage 2: link extraction
# ---------------------------------------------------------------------------

def fetch_url(url, timeout=REQUEST_TIMEOUT):
    headers = {"User-Agent": "Mozilla/5.0 (compatible; HackathonScreener/1.0)"}
    return requests.get(url, timeout=timeout, headers=headers, allow_redirects=True)


def extract_pdf_text(content_bytes):
    if PdfReader is None:
        return ""
    try:
        reader = PdfReader(io.BytesIO(content_bytes))
        return "\n".join((page.extract_text() or "") for page in reader.pages)
    except Exception:
        return ""


def extract_docx_text(content_bytes):
    if DocxDocument is None:
        return ""
    try:
        doc = DocxDocument(io.BytesIO(content_bytes))
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception:
        return ""


def fetch_resume_text(url, cache):
    url = (url or "").strip()
    if not url:
        return {"status": "not_readable", "reason": "Resume link is empty", "text": ""}
    if url in cache:
        return cache[url]

    result = None
    try:
        docs_m = DOCS_RE.search(url)
        drive_m = DRIVE_FILE_RE.search(url) or DRIVE_OPEN_RE.search(url)

        if docs_m:
            doc_id = docs_m.group(1)
            r = fetch_url(f"https://docs.google.com/document/d/{doc_id}/export?format=txt")
            text = r.text if r.status_code == 200 else ""
            if len(text) < MIN_READABLE_CHARS or looks_like_html_login(text, len(text)):
                result = {"status": "not_readable",
                          "reason": "Google Doc not readable (check sharing is set to Anyone with the link)",
                          "text": ""}
            else:
                result = {"status": "ok", "reason": "", "text": text}

        elif drive_m:
            file_id = drive_m.group(1)
            r = fetch_url(f"https://drive.google.com/uc?export=download&id={file_id}")
            content = r.content
            if content[:4] == b"%PDF":
                text = extract_pdf_text(content)
                if len(text) < MIN_READABLE_CHARS:
                    result = {"status": "not_readable",
                              "reason": "PDF has little or no extractable text (may be a scanned image; "
                                        "check sharing is set to Anyone with the link)", "text": ""}
                else:
                    result = {"status": "ok", "reason": "", "text": text}
            elif content[:2] == b"PK":
                text = extract_docx_text(content)
                if len(text) < MIN_READABLE_CHARS:
                    result = {"status": "not_readable",
                              "reason": "Word doc has little or no extractable text "
                                        "(check sharing is set to Anyone with the link)", "text": ""}
                else:
                    result = {"status": "ok", "reason": "", "text": text}
            else:
                text = content.decode("utf-8", errors="ignore")
                if len(text) < MIN_READABLE_CHARS or looks_like_html_login(text, len(text)):
                    result = {"status": "not_readable",
                              "reason": "Not readable (check sharing is set to Anyone with the link)", "text": ""}
                else:
                    result = {"status": "ok", "reason": "", "text": text}
        else:
            r = fetch_url(url)
            raw = r.text
            soup_text = BeautifulSoup(raw, "html.parser").get_text(separator=" ", strip=True)
            if len(soup_text) < MIN_READABLE_CHARS or looks_like_html_login(raw, len(soup_text)):
                result = {"status": "not_readable",
                          "reason": "Not readable (check sharing is set to Anyone with the link, "
                                    "or the link may require login)", "text": ""}
            else:
                result = {"status": "ok", "reason": "", "text": soup_text}

    except requests.exceptions.Timeout:
        result = {"status": "not_readable", "reason": "Request timed out after 15 seconds", "text": ""}
    except Exception as e:
        result = {"status": "not_readable", "reason": f"Could not fetch link: {e}", "text": ""}

    cache[url] = result
    return result


def fetch_portfolio(url, cache):
    url = (url or "").strip()
    if not url:
        return {"status": "needs_review", "note": "Portfolio or Case Study Link is empty", "text": ""}
    if url in cache:
        return cache[url]

    result = None
    try:
        m = GITHUB_REPO_RE.search(url) if "github.com" in url else None
        if m:
            owner, repo = m.group(1), m.group(2)
            repo = re.sub(r"\.git$", "", repo)
            api_base = f"https://api.github.com/repos/{owner}/{repo}"
            r = fetch_url(api_base)
            if r.status_code != 200:
                result = {"status": "needs_review",
                          "note": f"GitHub repo not accessible (HTTP {r.status_code})", "text": ""}
            else:
                data = r.json()
                description = data.get("description") or "(none)"
                language = data.get("language") or "Unknown"
                stars = data.get("stargazers_count", 0)

                commit_count_note = "unknown"
                last_commit_date = "unknown"
                commits_r = fetch_url(f"{api_base}/commits?per_page=1")
                if commits_r.status_code == 200:
                    link_header = commits_r.headers.get("Link", "")
                    m2 = re.search(r'page=(\d+)>; rel="last"', link_header)
                    commit_count_note = m2.group(1) if m2 else "1"
                    commits_json = commits_r.json()
                    if commits_json:
                        last_commit_date = commits_json[0].get("commit", {}).get("author", {}).get("date", "unknown")

                readme_text = ""
                readme_r = fetch_url(f"{api_base}/readme")
                if readme_r.status_code == 200:
                    rd = readme_r.json()
                    if rd.get("content"):
                        try:
                            readme_text = base64.b64decode(rd["content"]).decode("utf-8", errors="ignore")
                        except Exception:
                            readme_text = ""

                summary = (
                    f"GitHub repository: {owner}/{repo}\n"
                    f"Description: {description}\n"
                    f"Primary language: {language}\n"
                    f"Stars: {stars}\n"
                    f"Approx. commit count (last page of 1/page): {commit_count_note}\n"
                    f"Last commit date: {last_commit_date}\n\n"
                    f"README:\n{readme_text}"
                )
                if len(summary.strip()) < MIN_READABLE_CHARS:
                    result = {"status": "needs_review",
                              "note": "GitHub repo has very little content (needs manual review)", "text": summary}
                else:
                    result = {"status": "ok", "note": "", "text": summary}
        else:
            r = fetch_url(url)
            raw = r.text
            soup_text = BeautifulSoup(raw, "html.parser").get_text(separator=" ", strip=True)
            if len(soup_text) < MIN_READABLE_CHARS or looks_like_html_login(raw, len(soup_text)):
                result = {"status": "needs_review",
                          "note": "Portfolio page has very little content or requires login "
                                  "(needs manual review)", "text": soup_text}
            else:
                result = {"status": "ok", "note": "", "text": soup_text}

    except requests.exceptions.Timeout:
        result = {"status": "needs_review", "note": "Portfolio link timed out after 15 seconds "
                                                      "(needs manual review)", "text": ""}
    except Exception as e:
        result = {"status": "needs_review", "note": f"Could not fetch portfolio link: {e} "
                                                      "(needs manual review)", "text": ""}

    cache[url] = result
    return result


# ---------------------------------------------------------------------------
# Stage 3: AI scoring
# ---------------------------------------------------------------------------

def build_prompt(q_answers, resumes_by_member, portfolios_by_member):
    lines = [
        "You are helping a hackathon review panel score a team application against a fixed rubric. "
        "Score only what is provided. Do not try to guess anyone's identity. Respond with JSON only, "
        "no other text, no markdown code fences.",
        "",
        "Rubric:",
        "- technical_capability (0-25): from resumes only",
        "- portfolio_quality (0-25): depth, clarity, evidence of real work across the members' individual "
        "portfolio/case study links. No-code and business case studies count equally with code.",
        "- problem_thinking (0-20): specificity in the team's answers below, real users and problems, "
        "not buzzwords or filler",
        "- execution_track_record (0-15): shipped projects, past hackathons, internships. Reward "
        "substance over writing style.",
        "",
        "Team answers:",
    ]
    for label, ans in q_answers:
        lines.append(f"{label}: {ans if ans else '(no answer provided)'}")

    lines.append("")
    lines.append("Member resumes:")
    for label, text in resumes_by_member:
        snippet = text[:RESUME_TRUNCATE_CHARS] if text else "(resume text not available)"
        lines.append(f"--- {label} ---\n{snippet}")

    lines.append("")
    lines.append("Member portfolios / case studies:")
    for label, text in portfolios_by_member:
        snippet = text[:PORTFOLIO_TRUNCATE_CHARS] if text else "(portfolio not available)"
        lines.append(f"--- {label} ---\n{snippet}")

    lines.append("")
    lines.append(
        'Respond with JSON only, in exactly this shape:\n'
        '{"technical_capability": {"score": <0-25>, "reason": "..."}, '
        '"portfolio_quality": {"score": <0-25>, "reason": "..."}, '
        '"problem_thinking": {"score": <0-20>, "reason": "..."}, '
        '"execution_track_record": {"score": <0-15>, "reason": "..."}}'
    )
    return "\n".join(lines)


def strip_code_fences(raw):
    cleaned = raw.strip()
    cleaned = re.sub(r"^```(json)?", "", cleaned, flags=re.IGNORECASE).strip()
    cleaned = re.sub(r"```$", "", cleaned).strip()
    return cleaned


def try_parse_scores(raw):
    try:
        data = json.loads(strip_code_fences(raw))
        out = {}
        for key, mx in SCORE_MAXES.items():
            entry = data[key]
            score = float(entry["score"])
            score = max(0, min(mx, score))
            out[key] = {"score": score, "reason": str(entry.get("reason", ""))}
        return out
    except Exception:
        return None


def call_claude(client, prompt):
    msg = client.messages.create(
        model=MODEL_NAME,
        max_tokens=MAX_TOKENS,
        messages=[{"role": "user", "content": prompt}],
    )
    return "".join(getattr(block, "text", "") for block in msg.content)


def score_team_with_ai(client, prompt, score_cache):
    key = hashlib.sha256(prompt.encode("utf-8")).hexdigest()
    if key in score_cache:
        return score_cache[key]

    try:
        raw = call_claude(client, prompt)
    except Exception as e:
        return {"error": str(e)}

    parsed = try_parse_scores(raw)
    if parsed is None:
        try:
            raw2 = call_claude(client, prompt + "\n\nReturn ONLY valid JSON. No markdown, no commentary.")
        except Exception as e:
            return {"error": str(e)}
        parsed = try_parse_scores(raw2)
        if parsed is None:
            return {"error": "Model did not return valid JSON after one retry."}

    out = {"result": parsed}
    score_cache[key] = out
    return out


# ---------------------------------------------------------------------------
# Stage 4: team composition
# ---------------------------------------------------------------------------

def composition_score(skill_values):
    normalized = [normalize_skill(v) for v in skill_values]
    counts = Counter(s for s in normalized if s)
    unrecognized = sum(1 for s in normalized if not s)

    diff = sum(abs(counts.get(k, 0) - v) for k, v in IDEAL_COMPOSITION.items())
    diff += unrecognized

    counts_display = dict(counts)
    if unrecognized:
        counts_display["Unrecognized"] = unrecognized

    if diff == 0:
        return 15, f"Ideal composition (counts: {counts_display})"
    elif diff == 2:
        return 10, f"Off by one member from ideal composition (counts: {counts_display})"
    else:
        return 5, f"Composition does not match ideal distribution (counts: {counts_display})"


# ---------------------------------------------------------------------------
# Pipeline
# ---------------------------------------------------------------------------

def get_api_client():
    key = None
    try:
        key = st.secrets.get("ANTHROPIC_API_KEY")
    except Exception:
        key = None
    if not key:
        key = os.environ.get("ANTHROPIC_API_KEY")
    if not key or anthropic is None:
        return None, key
    return anthropic.Anthropic(api_key=key), key


def read_uploaded_file(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith(".xlsx"):
        df = pd.read_excel(uploaded_file, dtype=str)
    else:
        df = pd.read_csv(uploaded_file, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
    return df


def run_pipeline(df, client, progress_callback=None):
    q_cols = find_q_columns(df)
    email_team_map = build_email_team_map(df)
    link_cache = load_json_cache(LINK_CACHE_PATH)
    score_cache = load_json_cache(SCORE_CACHE_PATH)

    flagged = []
    scored = []
    errored = []

    rows = list(df.iterrows())
    total = len(rows)

    for idx, (_, row) in enumerate(rows):
        team = get_val(row, "Team Name")
        captain_email = get_val(row, "Captain Email")
        reasons = check_structural_eligibility(row, email_team_map)

        member_names = [get_val(row, member_field(i, "Full Name")) for i in range(1, 6)]
        member_emails = [get_val(row, member_field(i, "Email")) for i in range(1, 6)]
        redact_values = [get_val(row, "Captain Full Name"), captain_email] + member_names + member_emails

        resumes_by_member = []
        for i in range(1, 6):
            link = get_val(row, member_field(i, "Resume Link"))
            skill = get_val(row, member_field(i, "Skill Set"))
            res = fetch_resume_text(link, link_cache)
            label = f"Member {i} ({skill or 'Unknown role'})"
            if res["status"] != "ok":
                reasons.append(f"Member {i} resume link could not be read: {res['reason']}")
                resumes_by_member.append((label, ""))
            else:
                resumes_by_member.append((label, redact(res["text"], redact_values)))

        portfolios_by_member = []
        portfolio_review_notes = []
        for i in range(1, 6):
            link = get_val(row, member_field(i, "Portfolio or Case Study Link"))
            skill = get_val(row, member_field(i, "Skill Set"))
            res = fetch_portfolio(link, link_cache)
            label = f"Member {i} ({skill or 'Unknown role'})"
            if res["status"] != "ok":
                portfolio_review_notes.append(f"Member {i}: {res.get('note', '')}")
            portfolios_by_member.append((label, redact(res.get("text", ""), redact_values)))
        needs_manual_review = bool(portfolio_review_notes)
        portfolio_note = "; ".join(portfolio_review_notes)

        base_record = {
            "Team Name": team,
            "Captain Email": captain_email,
            "Member Emails": member_emails,
        }

        if reasons:
            flagged.append({**base_record, "Reasons": reasons})
        else:
            q_answers = []
            for prefix in Q_PREFIXES:
                col = q_cols.get(prefix)
                ans = get_val(row, col) if col else ""
                label = col if col else prefix
                q_answers.append((label, redact(ans, redact_values)))

            prompt = build_prompt(q_answers, resumes_by_member, portfolios_by_member)
            outcome = score_team_with_ai(client, prompt, score_cache)

            if "error" in outcome:
                errored.append({**base_record, "Error": outcome["error"]})
            else:
                ai_scores = outcome["result"]
                skill_values = [get_val(row, member_field(i, "Skill Set")) for i in range(1, 6)]
                comp_score, comp_reason = composition_score(skill_values)
                ai_total = sum(v["score"] for v in ai_scores.values())
                total_score = ai_total + comp_score

                scored.append({
                    **base_record,
                    "AI Scores": ai_scores,
                    "Composition Score": comp_score,
                    "Composition Reason": comp_reason,
                    "Total": total_score,
                    "Needs Manual Review": needs_manual_review,
                    "Portfolio Note": portfolio_note,
                })

        if progress_callback:
            progress_callback((idx + 1) / total if total else 1.0)

    save_json_cache(LINK_CACHE_PATH, link_cache)
    save_json_cache(SCORE_CACHE_PATH, score_cache)

    scored.sort(key=lambda r: r["Total"], reverse=True)
    for rank, rec in enumerate(scored, start=1):
        rec["Rank"] = rank
        if rank <= TOP_CUTOFF:
            rec["Bucket"] = "Top 30"
        elif rank <= WAITLIST_CUTOFF:
            rec["Bucket"] = "Waitlist"
        else:
            rec["Bucket"] = ""
        rec["Manual Review Band"] = REVIEW_BAND[0] <= rank <= REVIEW_BAND[1]

    return {"flagged": flagged, "scored": scored, "errored": errored}


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

def render_results_table(scored):
    rows = []
    for rec in scored:
        row = {
            "Rank": rec["Rank"],
            "Bucket": rec["Bucket"],
            "Team Name": rec["Team Name"],
            "Captain Email": rec["Captain Email"],
        }
        for key, label in SCORE_LABELS.items():
            row[label] = rec["AI Scores"][key]["score"]
        row["Composition"] = rec["Composition Score"]
        row["Total (/100)"] = rec["Total"]
        flags = []
        if rec["Needs Manual Review"]:
            flags.append("needs manual review (portfolio)")
        if rec["Manual Review Band"]:
            flags.append("borderline — review by hand")
        row["Notes"] = "; ".join(flags)
        rows.append(row)
    return pd.DataFrame(rows)


def build_download_csv(scored, flagged, errored):
    rows = []
    for rec in scored:
        row = {
            "Status": "Scored",
            "Rank": rec["Rank"],
            "Bucket": rec["Bucket"],
            "Team Name": rec["Team Name"],
            "Captain Email": rec["Captain Email"],
            "Total": rec["Total"],
            "Composition Score": rec["Composition Score"],
            "Composition Reason": rec["Composition Reason"],
            "Needs Manual Review": rec["Needs Manual Review"],
            "Portfolio Note": rec["Portfolio Note"],
            "Reasons/Errors": "",
        }
        for key, label in SCORE_LABELS.items():
            row[f"{label} Score"] = rec["AI Scores"][key]["score"]
            row[f"{label} Reason"] = rec["AI Scores"][key]["reason"]
        rows.append(row)

    for rec in flagged:
        rows.append({
            "Status": "Flagged",
            "Rank": "",
            "Bucket": "",
            "Team Name": rec["Team Name"],
            "Captain Email": rec["Captain Email"],
            "Total": "",
            "Composition Score": "",
            "Composition Reason": "",
            "Needs Manual Review": "",
            "Portfolio Note": "",
            "Reasons/Errors": "; ".join(rec["Reasons"]),
        })

    for rec in errored:
        rows.append({
            "Status": "Scoring Error",
            "Rank": "",
            "Bucket": "",
            "Team Name": rec["Team Name"],
            "Captain Email": rec["Captain Email"],
            "Total": "",
            "Composition Score": "",
            "Composition Reason": "",
            "Needs Manual Review": "",
            "Portfolio Note": "",
            "Reasons/Errors": rec["Error"],
        })

    return pd.DataFrame(rows).to_csv(index=False).encode("utf-8")


def main():
    st.title("Hackathon Team Screener")

    client, api_key = get_api_client()

    with st.sidebar:
        st.header("Setup")
        if api_key and client is not None:
            st.success("API key loaded ✓")
        else:
            st.error("API key not found")
            st.caption("Set ANTHROPIC_API_KEY in Streamlit secrets or your environment.")

        uploaded_file = st.file_uploader("Upload team roster (.csv or .xlsx)", type=["csv", "xlsx"])
        top_n = st.number_input("Top N (for Luma email export)", min_value=1, max_value=500, value=30, step=1)
        run_clicked = st.button("Run screening", type="primary", disabled=(uploaded_file is None or client is None))

    if run_clicked and uploaded_file is not None:
        try:
            df = read_uploaded_file(uploaded_file)
        except Exception as e:
            st.error(f"Could not read the uploaded file: {e}")
            return

        if "Team Name" not in df.columns:
            st.error('Missing required column: "Team Name"')
            return

        missing = validate_columns(df)
        if missing:
            st.error("The uploaded file is missing these required columns:")
            st.write(missing)
            return

        df = df[df["Team Name"].fillna("").astype(str).str.strip() != ""].reset_index(drop=True)
        total_teams = len(df)

        progress_bar = st.progress(0.0, text="Processing teams...")

        def update_progress(frac):
            progress_bar.progress(min(frac, 1.0), text=f"Processing teams... ({int(frac * 100)}%)")

        with st.spinner("Running eligibility checks, fetching links, and scoring..."):
            results = run_pipeline(df, client, progress_callback=update_progress)

        progress_bar.empty()
        st.session_state["results"] = results
        st.session_state["total_teams"] = total_teams

    results = st.session_state.get("results")
    total_teams = st.session_state.get("total_teams", 0)

    if not results:
        st.info("Upload a roster file and click **Run screening** to begin.")
        return

    scored = results["scored"]
    flagged = results["flagged"]
    errored = results["errored"]

    st.subheader("Summary")
    cols = st.columns(4)
    cols[0].metric("Total teams", total_teams)
    cols[1].metric("Eligible", len(scored) + len(errored))
    cols[2].metric("Flagged", len(flagged))
    cols[3].metric("Scored", len(scored))

    if errored:
        st.warning(f"{len(errored)} eligible team(s) could not be scored due to API errors:")
        st.dataframe(pd.DataFrame(errored)[["Team Name", "Captain Email", "Error"]], use_container_width=True)

    st.subheader("Ranked teams")
    if scored:
        table_df = render_results_table(scored)
        st.dataframe(table_df, use_container_width=True, hide_index=True)

        team_names = [f"#{r['Rank']} — {r['Team Name']}" for r in scored]
        pick = st.selectbox("View score reasoning for a team", ["(select a team)"] + team_names)
        if pick != "(select a team)":
            rec = scored[team_names.index(pick)]
            with st.expander(f"Score details — {rec['Team Name']}", expanded=True):
                for key, label in SCORE_LABELS.items():
                    entry = rec["AI Scores"][key]
                    st.markdown(f"**{label}: {entry['score']}/{SCORE_MAXES[key]}**")
                    st.write(entry["reason"])
                st.markdown(f"**Composition: {rec['Composition Score']}/15**")
                st.write(rec["Composition Reason"])
                if rec["Needs Manual Review"]:
                    st.info(f"Portfolio flagged for manual review: {rec['Portfolio Note']}")
    else:
        st.write("No teams were scored.")

    st.subheader("Flagged teams (not scored)")
    if flagged:
        flagged_rows = [{"Team Name": r["Team Name"], "Captain Email": r["Captain Email"],
                          "Reasons": "; ".join(r["Reasons"])} for r in flagged]
        st.dataframe(pd.DataFrame(flagged_rows), use_container_width=True, hide_index=True)
    else:
        st.write("No teams were flagged.")

    st.subheader("Export")
    csv_bytes = build_download_csv(scored, flagged, errored)
    st.download_button("Download results CSV", data=csv_bytes, file_name="screening_results.csv",
                        mime="text/csv")

    if scored:
        top_teams = scored[:top_n]
        email_lines = []
        for rec in top_teams:
            emails = [e for e in rec["Member Emails"] if e]
            email_lines.append(", ".join(emails))
        email_text = "\n".join(email_lines)
        st.text_area(f"Top {top_n} teams — member emails (comma-separated per team, for Luma invites)",
                     value=email_text, height=200)


if __name__ == "__main__":
    main()

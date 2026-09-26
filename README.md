# Hackathon Team Screener

A single-file Streamlit app that screens NYU AI Hackathon team applications: it
checks eligibility, reads resumes and portfolio links, scores each team with
Claude against the panel's 100-point rubric, and produces a ranked shortlist.

## Quick start (local)

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml   # then paste your real key in
streamlit run app.py
```

Try it first with `sample_teams.csv` (one strong team, one weak team, one
flagged team) before uploading a real roster export.

## Deploying privately (Streamlit Community Cloud)

1. Push this repo to GitHub (already done if you're reading this from there).
2. On share.streamlit.io, create a new app pointing at `app.py`.
3. In the app's Settings → Secrets, paste `ANTHROPIC_API_KEY = "sk-ant-..."`.
4. In Settings → Sharing, choose "Only specific people can view this app" and
   list your reviewer panel's email addresses. Do not make it public — it
   handles student resumes.

Full walkthrough, architecture notes, and troubleshooting: see the
project documentation shared alongside this repo.

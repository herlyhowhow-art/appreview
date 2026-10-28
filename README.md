# AI Scientific Article Reviewer

A runnable Streamlit web app for expert-style scientific manuscript review with PDF and DOCX report export. The output is a detailed AI review report only (not a completed checklist form), styled with the navy/blue table language and section hierarchy of the supplied ORI peer-review form.

## Features
- Upload PDF or DOCX manuscripts
- Extract text and preview before review
- AI Expert Review using an OpenAI API key (optional but recommended for deep critique)
- Preliminary rule-based screening when no API key is configured
- Structured scorecard, major/minor issues, section reviews, statistical audit, ethics/transparency audit, reference audit, and revision plan
- Download professionally formatted PDF and editable Word reports
- Detailed section-by-section report aligned with the supplied form headings: Title, Abstract, Introduction, Literature review, Methodology, Results / Findings, Discussion, Conclusion & suggestions, References & citations, Language, and Other issues
- Editorial recommendation is explicitly advisory; final decision remains with a human editor/reviewer
- Explicit limitations and no claims of guaranteed acceptance or verified Scopus/Web of Science indexing

## Requirements
- Python 3.10+
- Internet access for package installation; AI mode also requires API access

## Run locally

```bash
python -m venv .venv
# Windows PowerShell:
.venv\Scripts\Activate.ps1
# macOS/Linux:
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Open the local URL shown by Streamlit (usually http://localhost:8501).

## Enable AI review
1. Obtain an API key from your model provider account.
2. Enter the key in the app sidebar, or set `OPENAI_API_KEY` in your shell environment.
3. Choose an API model available to your account. The default is `gpt-4.1-mini`; availability can vary by account and provider changes.
4. Upload a text-based PDF or DOCX, preview extracted text, and click **Run expert review**.

The app sends extracted manuscript text to the configured model provider in AI mode. Do not upload confidential or unpublished manuscripts unless you are authorized to share them and have reviewed the provider's data-handling terms. The key is not written to a local database by this app.

## Rule-based screening
Choose **Preliminary rule-based screening** to run a local keyword/structure scan without an AI API key. It is not an expert review and should not be interpreted as a valid scientific-quality score.

## Output design and review governance
- The report borrows the supplied form’s navy/blue visual style and detailed-review section structure, but does not reproduce or fill the original checklist form because the selected output is the detailed report only.
- The AI is instructed not to invent line numbers. If reliable line numbering is unavailable in the extracted manuscript, it should refer to section headings or quoted text instead.
- Editorial recommendations require human confirmation.

## Known limitations
- Scanned/image-only PDFs may extract little or no text; OCR is not built in.
- Complex tables, equations, multi-column layouts, footnotes, and figure captions may extract imperfectly.
- No external bibliographic database lookup is implemented. DOI validity, reference matching, journal instructions, and Scopus/WoS indexing are not verified.
- No raw data or statistical software output is independently inspected. Statistical findings require human verification.
- No login, multi-user authorization, database, usage limits, audit trail, or production privacy controls are included.
- The app's score and recommendation are advisory and are not an acceptance probability or editorial decision.

## Production hardening before public launch
- Add authentication, role-based access, rate limiting, secure secret management, and monitoring.
- Establish data retention/deletion controls and user consent for AI processing.
- Add OCR and robust document parsing for tables, equations, and figures.
- Add DOI/reference verification against trusted bibliographic services and explicit source citations.
- Build test cases using known manuscripts and statistical edge cases; validate score calibration with qualified reviewers.
- Keep journal-specific checklists versioned and verify indexing only against official current sources.

## Architecture
- `app.py`: Streamlit UI, document extraction, review orchestration, and PDF/DOCX export.
- AI mode: manuscript text -> structured JSON review prompt -> validated/parsed report -> PDF/DOCX generators.
- Fallback mode: extracted text -> transparent heuristic checks -> same report schema -> PDF/DOCX generators.

## Deploy from GitHub to Streamlit Community Cloud

1. Create a GitHub repository named `appreview` (or another name).
2. Upload the contents of this folder to the repository root. `app.py`, `requirements.txt`, and `README.md` must be at the root.
3. Open https://share.streamlit.io/ and sign in with GitHub.
4. Choose **Create app**, select your repository and branch (usually `main`), and set the main file path to `app.py`.
5. Deploy. Choose an available `*.streamlit.app` subdomain in the deployment settings if prompted.
6. For AI review, add your provider key under **App settings → Secrets** using TOML syntax:

```toml
OPENAI_API_KEY = "your-key-here"
```

Never commit secrets to GitHub. Keep the repository private while testing if the manuscript workflow or code should not be public.

## GitHub smoke check

A GitHub Actions workflow at `.github/workflows/python-check.yml` runs a Python syntax check and verifies that the required deployment files exist on each push or pull request. This is a basic repository check; it does not replace a full live deployment test or scientific validation of review quality.

# GitHub-backed HTML Dashboard on Streamlit

## Architecture
- `dashboard_template.html`: the original interactive HTML dashboard.
- `data.xlsx`: the Excel source of truth, containing sheet `Raw Data`.
- `app.py`: fetches Excel from GitHub, creates the dashboard HTML, displays it in Streamlit, and lets the admin replace Excel through GitHub.
- Fixed admin password in code: `0000`.

## One-time configuration
Edit these values near the top of `app.py`:
```python
GITHUB_OWNER = "your-github-user-or-org"
GITHUB_REPO = "your-repository"
GITHUB_BRANCH = "main"
GITHUB_DATA_PATH = "data.xlsx"
```

In Streamlit App settings > Secrets, add a GitHub token with permission to write repository contents:
```toml
GITHUB_TOKEN = "your-token"
```
Do not commit the real token to GitHub.

## Deploy
1. Upload this folder's contents to the configured GitHub repository.
2. Create a Streamlit Community Cloud app using `app.py`.
3. Add `GITHUB_TOKEN` in Streamlit Secrets.
4. Open the app. The HTML dashboard will use `data.xlsx` from GitHub.

## Updating Excel
1. Enter admin password `0000` in the Streamlit sidebar.
2. Upload an `.xlsx` workbook containing sheet `Raw Data`.
3. Select **Commit Excel to GitHub**.
4. The app validates the workbook, commits it to GitHub, then reloads the dashboard.

TMLS is derived as `gAC + gMX + gTN`; `mMX` is excluded.

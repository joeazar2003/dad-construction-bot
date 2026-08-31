"""
Makes the GitHub repo the source of truth for the Excel file, so a Render
free-tier restart or spin-down can never lose data — only the file that's
sitting on Render's local (ephemeral) disk is at risk; the copy in GitHub
isn't.

Flow:
  - On process startup, `pull_latest()` downloads whatever is currently in
    the repo and writes it to the local XLSX_PATH, so the app always starts
    from the real data instead of an empty workbook.
  - After every write (add_expense, undo_last, importing an edited file),
    the caller saves locally as before and then calls `push_current()`,
    which commits the new bytes back to the repo. Every expense becomes one
    git commit, so you also get a full audit trail for free.

Requires three env vars: GITHUB_TOKEN (a fine-grained or classic Personal
Access Token with read/write access to the repo's contents), GITHUB_REPO
("owner/repo"), and GITHUB_FILE_PATH (where the workbook lives in the repo,
e.g. "data/site1.xlsx"). GITHUB_BRANCH defaults to "main".

If these aren't set (e.g. while testing locally), every function becomes a
harmless no-op so the rest of the bot still works against the local file only.
"""
import base64
import os
import time

import requests

GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")
GITHUB_REPO = os.environ.get("GITHUB_REPO", "")
GITHUB_FILE_PATH = os.environ.get("GITHUB_FILE_PATH", "")
GITHUB_BRANCH = os.environ.get("GITHUB_BRANCH", "main")

ENABLED = bool(GITHUB_TOKEN and GITHUB_REPO and GITHUB_FILE_PATH)

API_ROOT = "https://api.github.com"
_HEADERS = {
    "Authorization": f"Bearer {GITHUB_TOKEN}",
    "Accept": "application/vnd.github+json",
    "User-Agent": "dad-construction-bot",
}

# Cached so we don't need a GET before every single PUT.
_last_sha = None


def _contents_url():
    return f"{API_ROOT}/repos/{GITHUB_REPO}/contents/{GITHUB_FILE_PATH}"


def pull_latest(local_path):
    """Download the current file from GitHub into local_path. Returns True if a file was found and written."""
    global _last_sha
    if not ENABLED:
        return False

    resp = requests.get(_contents_url(), headers=_HEADERS, params={"ref": GITHUB_BRANCH}, timeout=15)
    if resp.status_code == 404:
        _last_sha = None
        return False
    resp.raise_for_status()
    data = resp.json()
    _last_sha = data.get("sha")
    content = base64.b64decode(data["content"])
    with open(local_path, "wb") as f:
        f.write(content)
    return True


def push_current(local_path, message):
    """Commit the current local file back to GitHub. Retries once on a stale-sha conflict."""
    global _last_sha
    if not ENABLED:
        return False

    with open(local_path, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode("ascii")

    payload = {"message": message, "content": content_b64, "branch": GITHUB_BRANCH}
    if _last_sha:
        payload["sha"] = _last_sha

    resp = requests.put(_contents_url(), headers=_HEADERS, json=payload, timeout=20)

    if resp.status_code == 409 or (resp.status_code == 422 and _last_sha):
        # Someone/something else changed the file since we last read it (sha is stale).
        # Refresh sha and retry once.
        get_resp = requests.get(_contents_url(), headers=_HEADERS, params={"ref": GITHUB_BRANCH}, timeout=15)
        if get_resp.status_code == 200:
            _last_sha = get_resp.json().get("sha")
            payload["sha"] = _last_sha
            resp = requests.put(_contents_url(), headers=_HEADERS, json=payload, timeout=20)

    resp.raise_for_status()
    _last_sha = resp.json()["content"]["sha"]
    return True


def push_current_with_retry(local_path, message, attempts=3, delay=1.5):
    """Best-effort push — logs are still safe locally within this process even if every attempt fails,
    but callers should still let the user know if this returns False."""
    for i in range(attempts):
        try:
            return push_current(local_path, message)
        except Exception:
            if i == attempts - 1:
                return False
            time.sleep(delay)
    return False

# Zebrafish Bestand Statistics — cloud deployment

This folder is a self-contained copy meant to become its own GitHub
repository, deployed on Streamlit Community Cloud (free). See the
docstring at the top of `zf_ufz_statistics_cloud.py` for the full
explanation of how the pieces fit together (GitHub hosts the code,
Streamlit Cloud runs it, Nextcloud stays purely the data source).

## Files in this folder

| File | Purpose |
|---|---|
| `zf_ufz_statistics_cloud.py` | The app Streamlit Cloud runs (its "main file path") |
| `zf_ufz_statistics.py` | Shared data-loading/plotting logic it imports |
| `ufz_logo.png` | Logo shown in the app |
| `requirements.txt` | Tells Streamlit Cloud which Python packages to install |
| `secrets.toml.example` | Reference only — what to paste into the Cloud dashboard, not a file it reads |
| `.gitignore` | Keeps a local `.streamlit/secrets.toml`, if you ever create one for testing, out of the repo |

## One-time setup

1. **Create a GitHub account**, if you don't have one: https://github.com/join
   (free).

2. **Create a new, empty repository** on GitHub (e.g. named
   `zf-bestand-cloud`). Leave it empty — don't let GitHub add a README,
   license, or .gitignore, since this folder already has those.

3. **Push this folder to it.** Open a terminal in this folder and run:

   ```
   git init
   git add .
   git commit -m "Initial cloud deployment"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<your-repo>.git
   git push -u origin main
   ```

   (Replace `<your-username>/<your-repo>` with the real repository URL —
   GitHub shows you this exact command on the new repository's page.)

4. **Create a Streamlit Community Cloud account**: go to
   https://share.streamlit.io and sign in with the same GitHub account
   (click "Continue with GitHub" — no separate password to set up).

5. **Deploy**: click "New app", choose the repository and branch you just
   pushed, and set **Main file path** to `zf_ufz_statistics_cloud.py`.
   Click "Deploy".

6. **Set the Nextcloud share password** (optional but recommended for a
   hands-off deployment): once the app exists, open its Settings >
   Secrets in the Streamlit Cloud dashboard, and paste in the contents of
   `secrets.toml.example` with the real password filled in. Save — the
   app restarts automatically and picks it up.

   Skipping this step is also fine: the page will simply ask each visitor
   to type the share password in themselves before it loads any data.

7. You'll get a URL like `https://<something>.streamlit.app` — that's
   what you share with the people who should have access.

## Updating it later

Whenever the code changes, from this same folder:

```
git add .
git commit -m "describe the change"
git push
```

Streamlit Community Cloud watches the repository and redeploys
automatically within a minute or so of each push — no separate "deploy"
step needed after the first time.

## Keeping this copy and the internal one in sync

This is a deliberately separate copy from `zf_ufz_statistics_nc.py` and
`zf_ufz_statistics.py` in the parent folder (the ones the office PC runs).
If you make a fix or improvement in one, remember to copy it across to
the other if it should apply to both.

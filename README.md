# CV / LSV Data Processor (web app)

A browser-based version of the CV/LSV analysis tool: upload raw voltammetry
data, pick a reference electrode and pH, enter electrode area, and download
a processed Excel file with the RHE potential and current density already
calculated (using live formulas, same as the desktop version).

Gated behind Google sign-in.

## Files

- `app.py` -- the whole app (UI + processing logic).
- `requirements.txt` -- Python dependencies.
- `.streamlit/config.toml` -- raises the file upload size limit to 400MB.
- `.streamlit/secrets.toml.example` -- template for the Google OAuth config.
  **Never commit a real `secrets.toml`** -- it's in `.gitignore`.

## 1. Set up Google Sign-In (one-time)

1. Go to the [Google Cloud Console credentials page](https://console.cloud.google.com/apis/credentials).
2. Create a project (or use an existing one).
3. Under **OAuth consent screen**, configure it (External user type is fine;
   add your app name, your email as support contact). You can leave it in
   "Testing" mode while you're the only user, or publish it once ready for
   others.
4. Under **Credentials** -> **Create Credentials** -> **OAuth client ID**:
   - Application type: **Web application**
   - Authorized redirect URIs, add BOTH:
     - `http://localhost:8501/oauth2callback` (for local testing)
     - `https://YOUR-APP-NAME.streamlit.app/oauth2callback` (fill in your
       real app name once you know it -- see step 3 below)
5. Copy the **Client ID** and **Client Secret** it generates.

## 2. Run locally (optional, to test before deploying)

```bash
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
```

Edit `.streamlit/secrets.toml` and fill in:
- `client_id` / `client_secret` from step 1
- `cookie_secret`: any long random string (e.g. run
  `python -c "import secrets; print(secrets.token_hex(32))"`)

Then:
```bash
streamlit run app.py
```
Open http://localhost:8501, click "Log in with Google", and test the CV/LSV
upload flow with a real raw data file.

## 3. Deploy on Streamlit Community Cloud (free)

1. Push this folder to a **public** GitHub repository (free tier requires
   public repos for unlimited apps; you get exactly one free private app if
   you'd rather keep it private).
   Make sure `.streamlit/secrets.toml` is NOT included (check `.gitignore`).
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in with
   GitHub, click **New app**, and point it at this repo and `app.py`.
3. Deploy. Note the URL Streamlit assigns you, e.g.
   `https://your-app-name.streamlit.app`.
4. Go back to Google Cloud Console -> your OAuth client -> add
   `https://your-app-name.streamlit.app/oauth2callback` to the authorized
   redirect URIs (if you didn't already know the exact name in step 1).
5. In the Streamlit Cloud dashboard, open your app -> **Settings** ->
   **Secrets**, and paste in the same `[auth]` block from your local
   `secrets.toml`, but with `redirect_uri` updated to
   `https://your-app-name.streamlit.app/oauth2callback`.
6. Reboot the app from the dashboard. Google sign-in should now work on the
   live site.

## Restricting who can sign in (optional)

By default, `st.login()` lets anyone with a Google account sign in. If you
only want specific people (e.g. your lab) to have access, add a check in
`app.py` right after the login gate, e.g.:

```python
ALLOWED_EMAILS = {"you@youruni.edu", "labmate@youruni.edu"}
if st.user.email not in ALLOWED_EMAILS:
    st.error("Your account isn't authorized to use this app.")
    st.stop()
```

## Notes

- Free tier: app sleeps after ~12 hours with no traffic (wakes up
  automatically on the next visit, takes a few seconds), ~1GB RAM, no custom
  domain.
- Uploaded files are processed in memory for that session only -- not saved
  to disk, not visible to other users.
- To add peak current / diffusion coefficient analysis later, that's a
  separate feature to build on top of this -- not included yet.

# IGI report data

Paste IGI report numbers and get each report's data in the 28-column sheet layout
(Lot no … Ratio), ready to download as Excel/CSV or copy straight into Excel.

Where the data comes from (both give the same JSON record):

1. IGI's API: `https://api.igi.org/ReportDetail.php?Printno=<report number>`, asked first.
2. The report-diagnosis page: `https://www.igi.org/API-IGI/report-diagnosis.php?r=<report number>`.
   It's used when the API answers with Cloudflare's "Verify you are human" check instead of data.

The app tries a direct request first, then a browser, and keeps using whichever way answered.

| File | What it is |
|---|---|
| `streamlit_app.py` | The app |
| `requirements.txt` | Python packages |
| `packages.txt` | Chromium + chromedriver, installed by Streamlit Cloud (IGI's page is read in a hidden browser) |

## Put it on Streamlit Community Cloud

1. On GitHub, create a new repository (for example `igi-report-data`), set it to **Private**, and create it.
2. On the new repository's page, click **uploading an existing file**, drag in the files from this folder, and click **Commit changes**.
3. Go to **share.streamlit.io** and sign in with GitHub. When GitHub asks, allow access to private repositories.
4. Click **Create app** → **Yup, I have an app**, then fill in:
   - Repository: `<your GitHub name>/igi-report-data`
   - Branch: `main`
   - Main file path: `streamlit_app.py`
5. Open **Advanced settings**, choose Python **3.12**, click **Save**, then **Deploy**.
   The first build takes a few minutes because it installs Chromium.
6. When the app opens, click **Check IGI connection** in the sidebar.

An app from a private repository is private: use **Share** (top right) to invite people by email.
Community Cloud allows one private app per account.

## If IGI blocks Streamlit Cloud

If the check says **IGI's Cloudflare is refusing this server**, IGI does not accept requests
from Streamlit Cloud's servers (Google Cloud, United States). Two ways round it:

- Ask IGI (igi-api-help@igi.org) to allow Streamlit Community Cloud's outgoing IP addresses
  on `api.igi.org`. The addresses are listed at https://docs.streamlit.io/deploy/streamlit-community-cloud/status
  (Streamlit says they can change without notice). Once IGI allows them, the app reads the API
  directly: fast, and no browser needed.
- Run the same app on a PC or server in your office, where IGI already works, and share it
  with a Cloudflare Tunnel link.

## Run it on a Windows PC

```
pip install -r requirements.txt
python streamlit_app.py
```

On Windows it opens a visible Chrome (undetected-chromedriver, like the IGI report script)
and adds a **Save Excel** button for `D:\Gaurav\Download`.

## Settings at the top of streamlit_app.py

- `SHAPE_CODES`: IGI shape name → your code. Only `OV` is confirmed; set the others to the codes your sheet uses.
- `DATE_FORMATS`, `DEFAULT_SAVE_FOLDER`, `CHROME_PAGE_TIMEOUT`.

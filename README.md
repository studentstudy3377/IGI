# IGI Report Data

Paste IGI report numbers and get each report's data in the 28-column sheet (Lot no … Ratio),
ready to download as Excel.

The same `streamlit_app.py` runs on your Windows PC and on Streamlit Cloud, with the same page:
report numbers, then the result; **Fetch**, **Clear all** and **Download Excel** in the sidebar.
Only what the Fetch button does is different:

| | Windows PC | Streamlit Cloud (igi-data.streamlit.app) |
|---|---|---|
| Fetch | The app reads IGI itself | Opens IGI in a new tab; your **IGI → sheet** bookmark reads it there and sends it back |
| Why | IGI answers your office connection | IGI's Cloudflare blocks Streamlit's servers |

## Use it on Streamlit Cloud

Once per browser (desktop Chrome or Edge):

1. Open the app. In the sidebar, open **First time: add the IGI → sheet bookmark**.
2. Drag the red **IGI → sheet** button onto your bookmarks bar (Ctrl+Shift+B shows the bar).
   If you added the bookmark from an older version of the app, add it again and delete the old one.

Every time:

1. Paste the report numbers, or rows copied from Excel including the header row.
2. Click **Fetch N report(s) from IGI** in the sidebar. IGI opens in a new tab.
3. On the IGI tab, click the **IGI → sheet** bookmark. It reads the reports, sends them to the app
   and closes the tab.
4. The table appears. **Download Excel** is in the sidebar.

If the result does not come back by itself, the bookmark copies it instead: open
**Result not back? Paste it here** in the sidebar, paste, and press Ctrl+Enter.

## Files

| File | What it is |
|---|---|
| `streamlit_app.py` | The app |
| `requirements.txt` | Python packages (the last three lines are only for the Windows PC) |
| `.streamlit/config.toml` | Colours, so the cloud app looks like the PC one |

There is no `packages.txt`: Streamlit Cloud does not need Chromium, since IGI is read in your browser.

## Update the app on Streamlit Cloud

The app is deployed from `studentstudy3377/IGI`, branch `main`, main file `streamlit_app.py`.

1. On GitHub: **Add file → Upload files**, drag in `streamlit_app.py`, `requirements.txt`, `README.md`
   and the `.streamlit` folder, then **Commit changes**.
   (Or create the theme file by hand: **Add file → Create new file**, name it `.streamlit/config.toml`,
   paste its content.)
2. Open `packages.txt` → **⋯** (top right) → **Delete file** → **Commit changes**.
3. Streamlit Cloud redeploys by itself. If the page still looks old after a few minutes, go to
   share.streamlit.io → **⋮** next to the app → **Reboot**.

## Run it on the Windows PC

```
pip install -r requirements.txt
python streamlit_app.py
```

On Windows the Fetch button reads IGI itself (a direct request first, then a visible Chrome through
undetected-chromedriver).

## Settings at the top of streamlit_app.py

- `SHAPE_CODES`: IGI shape name → your code. Only `OV` is confirmed; set the others to your sheet's codes.
- `REPORT_NO_AS_IGI`, `DATE_FORMAT`, `SKIP_MISSING`, `NO_DATA`, `NO_DATA_COLUMN`, `SAVE_FILE_NAME`.
- `READ_ON_SERVER`: True on Windows (the app reads IGI), False elsewhere (your browser reads IGI).

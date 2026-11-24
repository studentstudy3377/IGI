# IGI Report Data

Paste IGI report numbers, click **Fetch**, and get your 28-column sheet; **Download Excel** saves it.
Live at <https://igi-data.streamlit.app>.

## Files

| File | What it is |
|---|---|
| `streamlit_app.py` | The app. It also holds the IGI Reader extension, which the page offers as a download. |
| `requirements.txt` | streamlit, pandas, openpyxl, requests |
| `.streamlit/config.toml` | The colours of the local app (red #FF4B4B, grey boxes, no borders) |

## How the reports are read

IGI's API (`https://api.igi.org/ReportDetail.php?Printno=<number>`) refuses Streamlit Cloud's servers
(HTTP 403, Cloudflare's "Verify you are human" page) and sends no CORS header, so a page on
streamlit.app can't call it with JavaScript either. The page therefore reads IGI through your own
browser, with the **IGI Reader** extension:

1. **Fetch** hands the report numbers to IGI Reader.
2. IGI Reader asks IGI's API. If the API answers with its human check instead of data, it uses IGI's
   report page (`https://www.igi.org/API-IGI/report-diagnosis.php?r=<number>`), which carries the same data.
3. The reports come back to the page and fill the table. Reports IGI doesn't return show **No Data**.

## Install IGI Reader (once per PC)

1. Open the app, paste a report number and click **Fetch**. The page shows **Download IGI Reader extension**.
2. Unzip it to a folder you will keep, for example `C:\IGI Reader`.
3. Chrome: open `chrome://extensions` (Edge: `edge://extensions`) and switch on **Developer mode**.
4. Click **Load unpacked** and choose that folder.
5. Back in the app, press **F5** and click **Fetch**.

IGI Reader only works on `https://igi-data.streamlit.app` and `localhost`, and only when you click
Fetch there. If the app's address changes, change `APP_PAGES` in `streamlit_app.py` and install the
extension again.

## Running it on a PC

`streamlit run streamlit_app.py`. On Windows without the extension, the app reads the same two IGI
addresses itself. With the extension in the browser, it reads through the browser, as on the cloud.

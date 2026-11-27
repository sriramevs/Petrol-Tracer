# Petrol Maintenance Tracker: setup (about 10 minutes)

Files: `index.html` (the page you open on your phone) and `Code.gs` (the Google Apps Script backend).

## 1. Backend
1. Open your Google Sheet, then **Extensions → Apps Script**.
2. Keep a copy of your old script somewhere, then replace it with the contents of `Code.gs`.
   Entries go to a new tab called **Fuel Log**, so your old data tab is not touched.
3. **Project Settings** (gear icon):
   - Set the time zone to **Asia/Kolkata**.
   - Under **Script properties**, add:
     - `APP_PIN` = any PIN you choose (you type it once in the web page)
     - `GEMINI_API_KEY` = free key from https://aistudio.google.com/apikey
       (or use Claude instead: add `ANTHROPIC_API_KEY` and set `OCR_PROVIDER: 'claude'` at the top of `Code.gs`)
4. In the editor, pick `authorizeAll` in the function dropdown, click **Run**, and accept the permissions.
5. Pick `testPrice` and Run. Open **Execution log**: you should see one or two Chennai prices.
   If it shows `[]`, the price pages changed layout; the app then falls back to your last logged price and
   you can type the rate yourself.
6. **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy.**
   Keep "Execute as: Me" and "Who has access: Anyone". This keeps your existing `/exec` URL,
   which is already in `index.html`.

## 2. Frontend (GitHub Pages)
GitHub Pages is configured to publish the root of the `ver-1.2` branch.
After pushing changes to that branch, open https://sriramevs.github.io/Petrol-Tracer/ on your phone,
enter your PIN, and add it to your home screen. The HTTPS URL is required for camera access.

GitHub Pages hosts only the frontend; the backend remains the Google Apps Script web app above.
Do not put API keys or other secrets in `index.html`.

## 3. Using it
- Weekly at the pump: take a photo of the dashboard → odometer fills in → type the amount paid → Save.
- Litres = amount ÷ price. Mileage = km since last fill ÷ litres. This is accurate because you always fill to full.
- The first entry only records the starting odometer; mileage starts from the second fill-up.
- Photos are saved in the Drive folder **Dio Odometer Photos**, with an "Open photo" link in each sheet row.
  Set `SHOW_IMAGE_IN_SHEET: true` in `Code.gs` for thumbnails inside the sheet
  (this makes each photo viewable by anyone who has its link).

## Notes
- The price is the Chennai city price, not your specific pump's. Pumps of different companies in one city
  usually differ by only a few paise to a rupee at most. The pump slip or Kotak SMS shows the exact rate, so edit it if needed.
- If Gemini ever returns "model not found", update `GEMINI_MODEL` in `Code.gs`.

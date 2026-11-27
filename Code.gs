/**
 * Petrol Maintenance Tracker - backend (Google Apps Script web app)
 *
 * Actions (all via POST, JSON body sent as text/plain):
 *   summary -> last odometer, averages, recent fills, today's Chennai petrol price
 *   ocr     -> reads the TOTAL odometer from a dashboard photo (vision AI)
 *   log     -> saves the photo to Drive and appends a row to the "Fuel Log" sheet
 *
 * Secrets live in Project Settings -> Script properties (never in the HTML):
 *   APP_PIN            required. Same PIN you type once in the web page.
 *   GEMINI_API_KEY     if OCR_PROVIDER = 'gemini' (free key from aistudio.google.com)
 *   ANTHROPIC_API_KEY  if OCR_PROVIDER = 'claude'
 *   SHEET_ID           optional. Only needed if this script is NOT bound to your sheet.
 */

const CONFIG = {
  SHEET_NAME: 'Fuel Log',
  DRIVE_FOLDER_NAME: 'Dio Odometer Photos',

  OCR_PROVIDER: 'gemini',            // 'gemini' or 'claude'
  GEMINI_MODEL: 'gemini-2.5-flash',  // update if Google retires this name
  CLAUDE_MODEL: 'claude-haiku-5-5',

  // Sanity limits (warnings only, never block saving except odometer going backwards)
  TANK_CAPACITY_L: 5.3,              // check your Dio manual
  USUAL_MIN_FILL_L: 3.5,
  MAX_GAP_KM: 600,                   // weekly fills: more than this between fills is suspicious
  PRICE_MIN: 80,
  PRICE_MAX: 160,

  // false = photo link only. true = also show a thumbnail inside the sheet
  // (this makes the photo viewable by anyone who has its link).
  SHOW_IMAGE_IN_SHEET: false,

  PRICE_CACHE_SECONDS: 3 * 3600,
};

// Public pages that list the Chennai city petrol price. Text is stripped of HTML
// first, then matched with these patterns. If a site changes layout, that source
// just stops contributing; the others (or your manual override) still work.
const PRICE_SOURCES = [
  {
    name: 'bankbazaar',
    url: 'https://bb.bankbazaar.com/fuel/petrol-price-chennai.html',
    re: /Today(?:'|’)?s Petrol Price in Chennai\s*(?:₹|Rs\.?)\s*(\d{2,3}\.\d{1,2})/i,
  },
  {
    name: 'cardekho',
    url: 'https://www.cardekho.com/petrol-price-in-chennai-city',
    re: /Petrol price in Chennai stands at\s*(?:₹|Rs\.?)\s*(\d{2,3}\.\d{1,2})/i,
  },
];

const HEADERS = [
  'Logged at', 'Date', 'Odometer (km)', 'Amount (₹)', 'Price (₹/L)', 'Litres',
  'Km since last fill', 'Mileage (km/L)', '₹ per km', 'Price source', 'Warnings',
  'Photo link', 'Photo',
];

// ---------------------------------------------------------------- entry points

function doGet() {
  return json_({ ok: true, message: 'Petrol tracker backend is running.' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const pin = PropertiesService.getScriptProperties().getProperty('APP_PIN');
    if (!pin) return json_({ ok: false, error: 'APP_PIN is not set in Script properties.' });
    if (String(req.pin) !== pin) return json_({ ok: false, error: 'Wrong PIN.' });

    switch (req.action) {
      case 'summary': return json_(summary_());
      case 'ocr':     return json_(ocr_(req));
      case 'log':     return json_(log_(req));
      default:        return json_({ ok: false, error: 'Unknown action.' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------- sheet helpers

function getSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setColumnWidth(13, 130);
  }
  return sh;
}

function readRows_() {
  const sh = getSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const vals = sh.getRange(2, 1, n, HEADERS.length).getValues();
  return vals
    .map(function (r) {
      return {
        date: r[1],
        odo: Number(r[2]),
        amount: Number(r[3]),
        price: Number(r[4]),
        litres: Number(r[5]),
        km: r[6] === '' ? null : Number(r[6]),
        mileage: r[7] === '' ? null : Number(r[7]),
      };
    })
    .filter(function (r) { return r.odo > 0; });
}

function round_(x, d) {
  const k = Math.pow(10, d);
  return Math.round(x * k) / k;
}

// ---------------------------------------------------------------- summary

function summary_() {
  const tz = Session.getScriptTimeZone();
  const rows = readRows_();
  const last = rows.length ? rows[rows.length - 1] : null;

  // Full-tank-to-full-tank mileage: total km / total litres over all intervals.
  let kmSum = 0, litSum = 0;
  rows.forEach(function (r) {
    if (r.km > 0 && r.litres > 0) { kmSum += r.km; litSum += r.litres; }
  });

  const ym = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  let monthSpend = 0;
  rows.forEach(function (r) {
    if (r.date instanceof Date && Utilities.formatDate(r.date, tz, 'yyyy-MM') === ym) {
      monthSpend += r.amount;
    }
  });

  const recent = rows.slice(-5).reverse().map(function (r) {
    return {
      date: r.date instanceof Date ? Utilities.formatDate(r.date, tz, 'dd MMM yyyy') : String(r.date),
      odometer: r.odo,
      amount: r.amount,
      litres: round_(r.litres, 2),
      mileage: r.mileage === null ? null : round_(r.mileage, 1),
    };
  });

  return {
    ok: true,
    today: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
    lastOdometer: last ? last.odo : null,
    lastPrice: last ? last.price : null,
    fills: rows.length,
    avgMileage: litSum > 0 ? round_(kmSum / litSum, 1) : null,
    monthSpend: round_(monthSpend, 2),
    recent: recent,
    price: getPrice_(last ? last.price : null),
    limits: {
      tankMax: CONFIG.TANK_CAPACITY_L,
      usualMin: CONFIG.USUAL_MIN_FILL_L,
      maxGapKm: CONFIG.MAX_GAP_KM,
    },
  };
}

// ---------------------------------------------------------------- petrol price

function scrapePrices_() {
  const found = [];
  PRICE_SOURCES.forEach(function (s) {
    try {
      const res = UrlFetchApp.fetch(s.url, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36' },
      });
      if (res.getResponseCode() !== 200) return;
      const text = res.getContentText()
        .replace(/&#8377;|&#x20b9;/gi, '₹')
        .replace(/&nbsp;|&#160;/g, ' ')
        .replace(/<script[\s\S]*?<\/script>/gi, ' ')
        .replace(/<style[\s\S]*?<\/style>/gi, ' ')
        .replace(/<[^>]+>/g, ' ')
        .replace(/\s+/g, ' ');
      const m = text.match(s.re);
      if (!m) return;
      const p = parseFloat(m[1]);
      if (p >= CONFIG.PRICE_MIN && p <= CONFIG.PRICE_MAX) found.push({ source: s.name, price: p });
    } catch (err) { /* ignore this source */ }
  });
  return found;
}

function getPrice_(fallbackPrice) {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('price');
  if (hit) return JSON.parse(hit);

  const found = scrapePrices_();
  let out;
  if (found.length) {
    const sorted = found.map(function (f) { return f.price; }).sort(function (a, b) { return a - b; });
    const mid = sorted.length % 2
      ? sorted[(sorted.length - 1) / 2]
      : (sorted[sorted.length / 2 - 1] + sorted[sorted.length / 2]) / 2;
    out = {
      price: round_(mid, 2),
      source: 'Chennai city price from ' + found.map(function (f) { return f.source; }).join(' + '),
      live: true,
    };
    cache.put('price', JSON.stringify(out), CONFIG.PRICE_CACHE_SECONDS);
  } else {
    out = {
      price: fallbackPrice || null,
      source: fallbackPrice ? 'last logged price (online lookup failed)' : 'unavailable, enter manually',
      live: false,
    };
  }
  return out;
}

// ---------------------------------------------------------------- OCR

const OCR_PROMPT =
  'This is a photo of a motorcycle/scooter digital instrument cluster. ' +
  'Read ONLY the TOTAL odometer (the row labelled TOTAL, in km). ' +
  'Ignore the speed (km/h), clock, trip meters and the fuel gauge. ' +
  'Faint unlit segments are ghost digits, not part of the number. ' +
  'Reply with strict JSON only: ' +
  '{"digits":"<odometer digits as a string; use ? for any digit that is scratched, glared or unreadable; null if no odometer visible>",' +
  '"confidence":"high|medium|low","note":"<very short note, e.g. glare on first digit>"}';

function ocr_(req) {
  if (!req.image) return { ok: false, error: 'No image received.' };
  const mime = req.mime || 'image/jpeg';
  const raw = CONFIG.OCR_PROVIDER === 'claude'
    ? callClaude_(req.image, mime)
    : callGemini_(req.image, mime);

  const m = String(raw).match(/\{[\s\S]*\}/);
  if (!m) return { ok: false, error: 'Could not parse the OCR reply.' };
  const j = JSON.parse(m[0]);

  const digits = j.digits ? String(j.digits).replace(/[^0-9?]/g, '') : '';
  return {
    ok: true,
    digits: digits || null,
    odometer: digits && digits.indexOf('?') === -1 ? parseInt(digits, 10) : null,
    confidence: j.confidence || 'low',
    note: j.note || '',
  };
}

function callGemini_(b64, mime) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEY is not set in Script properties.');
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
    CONFIG.GEMINI_MODEL + ':generateContent';
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': key },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      contents: [{ parts: [{ text: OCR_PROMPT }, { inline_data: { mime_type: mime, data: b64 } }] }],
      generationConfig: { temperature: 0, responseMimeType: 'application/json' },
    }),
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Gemini error ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200));
  }
  const data = JSON.parse(res.getContentText());
  return data.candidates[0].content.parts[0].text;
}

function callClaude_(b64, mime) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('ANTHROPIC_API_KEY is not set in Script properties.');
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      model: CONFIG.CLAUDE_MODEL,
      max_tokens: 300,
      messages: [{
        role: 'user',
        content: [
          { type: 'image', source: { type: 'base64', media_type: mime, data: b64 } },
          { type: 'text', text: OCR_PROMPT },
        ],
      }],
    }),
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Claude API error ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200));
  }
  return JSON.parse(res.getContentText()).content[0].text;
}

// ---------------------------------------------------------------- logging a fill

function saveImage_(b64, mime, dateStr, odo) {
  const it = DriveApp.getFoldersByName(CONFIG.DRIVE_FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.DRIVE_FOLDER_NAME);
  const ext = mime === 'image/png' ? 'png' : 'jpg';
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, dateStr + '_' + odo + 'km.' + ext);
  return folder.createFile(blob);
}

function log_(req) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const odo = Number(req.odometer);
    const amount = Number(req.amount);
    const price = Number(req.price);
    if (!(odo > 0) || !(amount > 0) || !(price >= CONFIG.PRICE_MIN && price <= CONFIG.PRICE_MAX)) {
      return { ok: false, error: 'Please check odometer, amount and price.' };
    }

    const rows = readRows_();
    const last = rows.length ? rows[rows.length - 1] : null;
    if (last && odo <= last.odo) {
      return { ok: false, error: 'Odometer (' + odo + ') must be higher than the last entry (' + last.odo + ').' };
    }

    const litres = amount / price;
    const km = last ? odo - last.odo : null;
    const mileage = km && litres > 0 ? km / litres : null;

    const warnings = [];
    if (litres < CONFIG.USUAL_MIN_FILL_L || litres > CONFIG.TANK_CAPACITY_L) {
      warnings.push('Litres ' + round_(litres, 2) + ' is outside ' + CONFIG.USUAL_MIN_FILL_L + '-' +
        CONFIG.TANK_CAPACITY_L + ' L: tank not full, or price/amount wrong?');
    }
    if (km !== null && km > CONFIG.MAX_GAP_KM) {
      warnings.push(km + ' km since last fill: a fill-up may be missing, so mileage is unreliable.');
    }

    const tz = Session.getScriptTimeZone();
    const dateStr = /^\d{4}-\d{2}-\d{2}$/.test(req.date || '')
      ? req.date
      : Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const p = dateStr.split('-').map(Number);
    const dateObj = new Date(p[0], p[1] - 1, p[2]);

    let link = '', preview = '';
    if (req.image) {
      const f = saveImage_(req.image, req.mime || 'image/jpeg', dateStr, odo);
      link = '=HYPERLINK("' + f.getUrl() + '","Open photo")';
      if (CONFIG.SHOW_IMAGE_IN_SHEET) {
        f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        preview = '=IMAGE("https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w400")';
      }
    }

    const sh = getSheet_();
    const r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, HEADERS.length).setValues([[
      new Date(), dateObj, odo, amount, price,
      '=D' + r + '/E' + r,
      r > 2 ? '=C' + r + '-C' + (r - 1) : '',
      r > 2 ? '=IF(F' + r + '>0,G' + r + '/F' + r + ',"")' : '',
      r > 2 ? '=IF(G' + r + '>0,D' + r + '/G' + r + ',"")' : '',
      req.priceSource || '',
      warnings.join(' | '),
      link,
      preview,
    ]]);
    sh.getRange(r, 6).setNumberFormat('0.00');
    sh.getRange(r, 8).setNumberFormat('0.0');
    sh.getRange(r, 9).setNumberFormat('0.00');
    if (preview) sh.setRowHeight(r, 90);

    return {
      ok: true,
      row: r,
      litres: round_(litres, 2),
      km: km,
      mileage: mileage === null ? null : round_(mileage, 1),
      costPerKm: km ? round_(amount / km, 2) : null,
      warnings: warnings,
    };
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- run these once from the editor

/** Run once: triggers the Google permission prompts (Sheets, Drive, external requests). */
function authorizeAll() {
  getSheet_();
  DriveApp.getRootFolder();
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  Logger.log('Authorised.');
}

/** Run to see what the price lookup returns (View -> Logs). */
function testPrice() {
  Logger.log(JSON.stringify(scrapePrices_()));
}

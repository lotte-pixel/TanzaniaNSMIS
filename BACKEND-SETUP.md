# NSMIS Form — Backend Setup (Google Sheets, ~5 minutes)

The form (`nsmis-form.html`) runs in **two parts on different days**:

- **Day 1 (Before training)** — participants fill in name, phone, three open questions, and their *Before* scores. They see their personal "Before" radar.
- **Day 6 (After training)** — participants come back, enter the same phone number, fill in their *After* scores. They see the combined Before/After radar; the admin can toggle the class average.

For this to work across devices (e.g. participants use their phones on Day 1 and laptops on Day 6, and the admin sees everything), the form needs a shared backend. The instructions below set one up using **Google Sheets + Apps Script** in about 5 minutes. It's free and runs entirely in your Google account.

If you skip this setup, the form falls back to `localStorage` and only works on a single device — fine for a quick offline demo, but not for a real training.

---

## 1. Create the Google Sheet

1. Go to <https://sheets.google.com> and create a new blank spreadsheet.
2. Rename it something like **"NSMIS Training Submissions"**.
3. Leave the first row blank — the script will write the header automatically on the first submission.

(If you'd rather set the header manually, paste this into row 1 — tab-separated:)

```
id	timestamp	timestamp_after	name	phone	q1	q2	q3	before_knowledge	before_skills	before_quality	before_confidence	before_ownership	after_knowledge	after_skills	after_quality	after_confidence	after_ownership
```

---

## 2. Add the Apps Script

1. In your sheet, open **Extensions → Apps Script**.
2. Delete the starter code and paste in the script below.
3. Click 💾 to save. Give the project a name (e.g. *NSMIS Backend*).

```javascript
const SHEET_NAME = 'Sheet1'; // change if your tab is named differently

const COLS = [
  'id','timestamp','timestamp_after','name','phone','q1','q2','q3',
  'before_knowledge','before_skills','before_quality','before_confidence','before_ownership',
  'after_knowledge','after_skills','after_quality','after_confidence','after_ownership'
];

function getSheet_() {
  return SpreadsheetApp.getActive().getSheetByName(SHEET_NAME)
      || SpreadsheetApp.getActive().getActiveSheet();
}

function normalizePhone_(p) {
  // Match on last 9 digits to tolerate country code / leading-zero variation.
  return String(p || '').replace(/\D/g, '').slice(-9);
}

function ok_(extra) {
  return ContentService.createTextOutput(JSON.stringify(Object.assign({ok:true}, extra || {})))
    .setMimeType(ContentService.MimeType.JSON);
}
function fail_(msg) {
  return ContentService.createTextOutput(JSON.stringify({ok:false, error: msg}))
    .setMimeType(ContentService.MimeType.JSON);
}

function findRowByPhone_(sheet, headers, phone) {
  const data = sheet.getDataRange().getValues();
  const phoneCol = headers.indexOf('phone');
  const target = normalizePhone_(phone);
  for (let i = data.length - 1; i >= 1; i--) {
    if (normalizePhone_(data[i][phoneCol]) === target) {
      return { rowIndex: i + 1, rowValues: data[i] }; // 1-based sheet row
    }
  }
  return null;
}

function rowToObj_(headers, r) {
  const obj = {};
  headers.forEach((h, i) => obj[h] = r[i]);
  const hasAfter = obj.timestamp_after !== '' && obj.timestamp_after !== null && obj.timestamp_after !== undefined;
  return {
    id: obj.id,
    timestamp: obj.timestamp,
    timestamp_after: obj.timestamp_after || '',
    name: obj.name, phone: obj.phone,
    q1: obj.q1, q2: obj.q2, q3: obj.q3,
    before: {
      knowledge:  Number(obj.before_knowledge)  || 0,
      skills:     Number(obj.before_skills)     || 0,
      quality:    Number(obj.before_quality)    || 0,
      confidence: Number(obj.before_confidence) || 0,
      ownership:  Number(obj.before_ownership)  || 0
    },
    after: hasAfter ? {
      knowledge:  Number(obj.after_knowledge)  || 0,
      skills:     Number(obj.after_skills)     || 0,
      quality:    Number(obj.after_quality)    || 0,
      confidence: Number(obj.after_confidence) || 0,
      ownership:  Number(obj.after_ownership)  || 0
    } : null
  };
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    if (sheet.getLastRow() === 0) sheet.appendRow(COLS);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    if (body.phase === 'after') {
      // Find existing row by phone and write After columns.
      const found = findRowByPhone_(sheet, headers, body.phone);
      if (!found) return fail_('not_found');
      const updates = {
        timestamp_after:     body.timestamp_after || new Date().toISOString(),
        after_knowledge:    (body.after || {}).knowledge,
        after_skills:       (body.after || {}).skills,
        after_quality:      (body.after || {}).quality,
        after_confidence:   (body.after || {}).confidence,
        after_ownership:    (body.after || {}).ownership
      };
      Object.keys(updates).forEach(k => {
        const c = headers.indexOf(k);
        if (c >= 0 && updates[k] !== undefined && updates[k] !== null) {
          sheet.getRange(found.rowIndex, c + 1).setValue(updates[k]);
        }
      });
      return ok_();
    }

    // Default = "before" submission → append a new row.
    // If the same phone already has a Day-1 row, overwrite it instead of
    // creating a duplicate (so someone who restarts Day 1 doesn't double up).
    const existing = findRowByPhone_(sheet, headers, body.phone);
    const row = COLS.map(c => {
      switch (c) {
        case 'id':              return body.id || '';
        case 'timestamp':       return body.timestamp || new Date().toISOString();
        case 'timestamp_after': return '';
        case 'name':            return body.name  || '';
        case 'phone':           return body.phone || '';
        case 'q1':              return body.q1 || '';
        case 'q2':              return body.q2 || '';
        case 'q3':              return body.q3 || '';
        case 'before_knowledge':  return (body.before || {}).knowledge  ?? '';
        case 'before_skills':     return (body.before || {}).skills     ?? '';
        case 'before_quality':    return (body.before || {}).quality    ?? '';
        case 'before_confidence': return (body.before || {}).confidence ?? '';
        case 'before_ownership':  return (body.before || {}).ownership  ?? '';
        default: return '';
      }
    });
    if (existing) {
      // Preserve an already-completed After section, if any.
      const afterCols = ['after_knowledge','after_skills','after_quality','after_confidence','after_ownership','timestamp_after'];
      afterCols.forEach(c => {
        const idx = headers.indexOf(c);
        if (idx >= 0) row[idx] = existing.rowValues[idx] || '';
      });
      sheet.getRange(existing.rowIndex, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    return ok_();
  } catch (err) {
    return fail_(String(err));
  }
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  const sheet = getSheet_();
  if (sheet.getLastRow() < 2) {
    if (action === 'lookup') return fail_('not_found');
    if (action === 'list')   return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON);
    return ContentService.createTextOutput('NSMIS backend OK').setMimeType(ContentService.MimeType.TEXT);
  }
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1).map(r => rowToObj_(headers, r));

  if (action === 'lookup') {
    const phone = (e.parameter.phone || '').toString();
    const target = normalizePhone_(phone);
    const match = [...rows].reverse().find(r => normalizePhone_(r.phone) === target);
    if (!match) return fail_('not_found');
    return ok_({ submission: match });
  }
  if (action === 'list') {
    return ContentService.createTextOutput(JSON.stringify(rows))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('NSMIS backend OK').setMimeType(ContentService.MimeType.TEXT);
}
```

---

## 3. Deploy as a web app

1. Click **Deploy → New deployment**.
2. Next to "Select type", click the gear icon and choose **Web app**.
3. Fill in:
   - *Description*: `NSMIS submissions endpoint`
   - *Execute as*: **Me**
   - *Who has access*: **Anyone**  ← important, otherwise participants can't submit
4. Click **Deploy**.
5. Google will ask you to **authorize** the script. Click through, choose your Google account, click "Advanced → Go to … (unsafe)" if it warns you (your own script is safe), and **Allow**.
6. Copy the **Web app URL** at the end. It looks like:
   `https://script.google.com/macros/s/AKfycb...../exec`

> ⚠ Whenever you change the script, redeploy via **Deploy → Manage deployments → ✏️ → New version → Deploy**. The URL stays the same.

---

## 4. Wire the URL into the form

Open `nsmis-form.html` and find this line near the top of the `<script>` block:

```javascript
const BACKEND_URL = "";
```

Paste your URL between the quotes:

```javascript
const BACKEND_URL = "https://script.google.com/macros/s/AKfycb...../exec";
```

Save the file. The form now:

- on **Day 1**, POSTs the participant's info + Before scores → new row in the sheet,
- on **Day 6**, GETs `?action=lookup&phone=…` to find that row, then POSTs the After scores → same row gets updated,
- on the admin view (`?admin=1`), GETs `?action=list` and renders the table + averages.

---

## 5. (Optional) Lock the admin view with a passcode

In the same `<script>` block:

```javascript
const ADMIN_PASSCODE = "1";
```

Change it to a passcode of your choice, e.g. `"akvo2026"`. The admin URL then becomes:

```
…/nsmis-form.html?admin=akvo2026
```

Any other `?admin=` value will not unlock the admin view.

---

## 6. Hosting the HTML

You can host `nsmis-form.html` anywhere static — for example:

- **GitHub Pages** — push the file to a repo, enable Pages.
- **Netlify Drop** — drag-and-drop the file at <https://app.netlify.com/drop>.
- **Akvo's own web hosting** — upload alongside any other static page.

Share the same URL with participants for both Day 1 and Day 6. They pick which path applies (start vs. return) on the welcome screen.

---

## How phone matching works

Participants might type their number differently on Day 1 vs Day 6 — `+254 700 000 000`, `254-700-000-000`, `0700 000 000` all refer to the same person. Both the form and the Apps Script normalize phone numbers by stripping non-digits and matching on the **last 9 digits**, so all of the above are treated as the same person.

If a country in your training uses fewer than 9-digit local numbers, change `slice(-9)` to a smaller value in two places: in the Apps Script `normalizePhone_` function above, and in the HTML's `normalizePhone` function.

---

## Troubleshooting

- **Submissions are not appearing in the sheet** → Re-check step 3.3: *Who has access* must be **Anyone**. Re-deploy (`Deploy → Manage deployments → ✏️ → New version → Deploy`) and use the URL it gives you.
- **Day 6 lookup says "couldn't find an entry"** → Confirm the participant submitted Day 1 with this backend connected (not just localStorage). Check the Google Sheet for their row. If the phone formats differ wildly (e.g. with vs. without country code), the last-9-digits match should still work; if not, check both rows and adjust `normalizePhone_` if needed.
- **"Show class average" only shows my own line** → The form falls back to `localStorage` if the backend can't be reached. Confirm `BACKEND_URL` is set and that opening it in a browser shows `NSMIS backend OK`.
- **CORS error on POST** → The form posts with `mode: "no-cors"`, so the browser will not show you the response, but the data still arrives at the sheet. This is expected. GET requests (lookup + admin list) **do** return responses — if those fail, your Apps Script is not deployed as *Anyone* access.

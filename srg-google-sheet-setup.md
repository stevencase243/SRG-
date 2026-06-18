# SRG Rules Page — Google Sheet Setup Guide

Follow these steps once. After setup, every driver who submits the form will automatically
appear as a new row in your spreadsheet.

---

## Step 1 — Create the Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a new blank spreadsheet.
2. Name it something like **SRG Rule Acknowledgments**.
3. In **Row 1**, add these column headers exactly:

   | A | B | C | D |
   |---|---|---|---|
   | Timestamp | Name | Email | Acknowledged |

---

## Step 2 — Create the Apps Script

1. With the spreadsheet open, click **Extensions → Apps Script**.
2. Delete everything in the editor and paste the code below:

```javascript
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data  = JSON.parse(e.postData.contents);

    sheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.name       || '',
      data.email      || '',
      data.acknowledged ? 'Yes' : 'No'
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

3. Click **Save** (floppy disk icon) and give the project a name like **SRG Form Handler**.

---

## Step 3 — Deploy as a Web App

1. Click **Deploy → New deployment**.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Fill in the settings:
   - **Description**: SRG Rule Acknowledgment
   - **Execute as**: Me
   - **Who has access**: **Anyone**
4. Click **Deploy**.
5. Google will ask you to authorize the script — click through the prompts (you may need to click "Advanced → Go to SRG Form Handler" if you see a warning).
6. After deployment, copy the **Web app URL** — it looks like:
   `https://script.google.com/macros/s/XXXXXXXX/exec`

---

## Step 4 — Connect it to the HTML page

1. Open **srg-rules.html** in a text editor.
2. Near the top of the `<script>` section, find this line:

   ```javascript
   const APPS_SCRIPT_URL = 'YOUR_APPS_SCRIPT_URL_HERE';
   ```

3. Replace `YOUR_APPS_SCRIPT_URL_HERE` with the URL you copied in Step 3:

   ```javascript
   const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/XXXXXXXX/exec';
   ```

4. Save the file and upload it to your website.

---

## Step 5 — Test it

1. Open the page in a browser, fill in a test name/email, check the box, and submit.
2. Check your Google Sheet — a new row should appear within a few seconds.

> **Note:** The form uses `mode: 'no-cors'` which means the page can't read the response
> from Apps Script, but the data still gets written to the sheet. This is normal behavior
> and submissions will still be recorded reliably.

---

## Adding More Rule Sections Later

In **srg-rules.html**, find the comment block:

```html
<!-- ADD MORE RULE SECTIONS HERE -->
```

Copy the `.rules-section` block above it, update the section number and content, and paste it before that comment.

---

That's it! You now have a live rules page that logs every driver acknowledgment to Google Sheets.

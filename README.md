# Get to Know Your Posse

A mobile-first web app of conversation questions for incoming **UW–Madison Posse**
scholars during **SOAR 2026**. Scholars open it from a QR code on their phones and use
it to spark real conversations in pairs ("dyads") and with roommates — questions are
organized by **category** and by **depth** (warm-up → deep).

**Live site:** https://adelatorre2.github.io/uwposse-dyad-questions/

---

## ⚠️ This repository is PUBLIC

It must contain **questions only**. Never add scholar names, rosters, surveys, face
sheets, flight manifests, or any personal data. A `.gitignore` blocks common
document/PII file types (`.pdf`, `.docx`, `.pptx`, `.csv`) as a safety net, but the
final responsibility is yours — **if in doubt, don't commit it.**

---

## How it works

It's a plain static site — **no framework, no backend, no build step.**

- `index.html`, `css/styles.css`, `js/app.js` — the app.
- `data/questions.xlsx` — the questions, read **live in the browser** via
  [SheetJS](https://sheetjs.com/) (loaded from a CDN).
- To change the questions you just edit the Excel file and push. Nothing to rebuild.

---

## Adding or editing questions

1. Open **`data/questions.xlsx`** in Excel (or Numbers / Google Sheets).
2. Use the sheet named **`Questions`**. It has three columns:

   | Column     | What to put                                                        |
   |------------|--------------------------------------------------------------------|
   | `category` | A group name, e.g. `Icebreakers`, `Roots & Cities`. Reused names are grouped together. A brand-new category name automatically gets its own color and a default emoji. |
   | `depth`    | A number **1–4**: `1` = warm-up, `2` = light, `3` = deeper, `4` = deep. Blank defaults to `1`. |
   | `question` | The question text. Rows with no question are skipped.              |

3. **Save** the file (keep it as `.xlsx`).
4. Publish the change one of two ways:

   **A. Commit & push (command line)**
   ```bash
   git add data/questions.xlsx
   git commit -m "Update questions"
   git push
   ```

   **B. Upload via GitHub's website (no command line)**
   - Go to the repo → open the `data/` folder.
   - Click **Add file → Upload files**, drag in your updated `questions.xlsx`
     (same filename), and **Commit changes**.

> ⏳ **GitHub Pages can take up to ~10 minutes** to publish the update. The site
> also cache-busts the file, so once Pages redeploys, a refresh shows the new questions.

### Optional: add a logo
Drop a `assets/logo.png` into the repo and the header uses it automatically; without
one, a styled text header is shown.

---

## Regenerating the QR code

The printable QR code (`assets/dyad-questions-qr.png`) points at the live site. To
regenerate it (e.g. if the URL changes):

```bash
python3 -m pip install --user "qrcode[pil]"
python3 -c "import qrcode; qrcode.make('https://adelatorre2.github.io/uwposse-dyad-questions/').save('assets/dyad-questions-qr.png')"
git add assets/dyad-questions-qr.png && git commit -m "Update QR code" && git push
```

---

## Deploying / GitHub Pages settings

Pages is served from the **`main`** branch, **root (`/`)** folder. If you ever need to
re-enable it: repo **Settings → Pages → Source: Deploy from a branch → `main` / `/root`
→ Save**. The `.nojekyll` file ensures GitHub serves the files as-is.

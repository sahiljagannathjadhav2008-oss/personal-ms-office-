# Office Web — Word, Excel & PowerPoint in the browser

Developer: **Sahil Jadhav**

Ek pura static website jo MS Office jaisa kaam karta hai. Koi server, build ya library nahi — sirf HTML/CSS/JS.

## Apps
- **Word** – rich text editor, styles, tables, images, find/replace, templates, export `.docx` / `.html` / `.txt` / `.md` / PDF.
- **Excel** – 400+ formulas, charts, formatting, sort/filter, validation, conditional formatting, freeze panes, multiple sheets, export `.xlsx` / `.csv` / `.html` / PDF.
- **PowerPoint** – slides, shapes, images, tables, themes, transitions, animations, full-screen slide show, export `.pptx` / PDF / PNG.

Teeno apps phone, tablet, laptop aur PC par chalte hain. Data kahin upload nahi hota; sab kuch aapke browser mein rehta hai.
**Website band karne ke baad data save nahi rahta — isliye File → Export / Save As se file export kar lijiye.**

## GitHub Pages par host karna
1. GitHub par naya repository banayein (jaise `office-web`).
2. Is zip ke *andar ki saari files* (index.html, css/, js/, icons/, sw.js, manifest.webmanifest, .nojekyll) repository ke root mein upload karein.
3. Repository → **Settings → Pages → Build and deployment → Source: Deploy from a branch → Branch: `main` / `(root)`** → Save.
4. 1–2 minute baad site `https://<username>.github.io/<repo>/` par live hogi.

Local chalane ke liye: folder mein `python3 -m http.server` chalaiye aur `http://localhost:8000` kholiye (index.html seedha double-click karne se bhi zyada tar kaam karta hai, lekin PWA/service worker ke liye http chahiye).

## Shortcuts (kuch)
Ctrl+Z / Ctrl+Y undo-redo, Ctrl+B/I/U, Ctrl+C/X/V, Ctrl+F find, Ctrl+1 Format Cells (Excel), F2 cell edit, F5 slide show (PowerPoint).

## Known limits
- PowerPoint mein text formatting poore text-box par lagti hai (per-word nahi); animations sirf slide show mein chalte hain (.pptx mein export nahi hote).
- Browser clipboard permissions ke kaaran ribbon ka Paste kabhi kabhi Ctrl+V maangta hai.

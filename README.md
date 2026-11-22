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


## Version 2 – full feature pass (developer: Sahil Jadhav)

**Word** – track changes, comments, mail merge, captions/citations/index/TOC, equations, charts/SmartArt/WordArt, drop cap, columns, section breaks, watermark, page borders, themes, styles, macros (JS), content controls.
**Excel** – charts (many types + elements), sparklines, shapes/objects, tables + structured references, pivot tables, what-if (scenarios, goal seek, data tables), Solver, forecast, Data Analysis, Flash Fill, subtotal/group/consolidate, advanced filter, audit tools, protect sheet/workbook, page setup/page-break view, 400+ functions.
**PowerPoint** – rich text, groups, shapes/pictures effects, picture tools, charts/SmartArt/WordArt/icons/equations/video/audio, Slide Master + layouts, header/footer, sections, comments, outline/sorter/notes views, **animations** (entrance / emphasis / exit / motion paths, With/After previous, delay, repeat, triggers, Animation Pane, Animation Painter), **20 slide transitions** (direction, advance-after), **slide show** (Presenter View with notes+next slide+timer, pen/highlighter/eraser/laser, black/white screen, go-to-slide, see-all-slides, hyperlinks/actions, Set Up Show, kiosk/loop, Custom Shows, Rehearse Timings, hidden slides), export to .pptx **with animations, transitions and hyperlinks**.

Files: `js/ppt-x1.js` (objects/format), `js/ppt-x2.js` (master/views), `js/ppt-x3.js` (animation, transitions, slide show, export hooks).

### Known limits
- No cloud/OneDrive/co-authoring/password protection (by design – static site).
- Translate and online pictures/videos need internet; not testable offline.
- In .pptx export: charts, SmartArt, WordArt, icons, equations, groups and filtered/cropped pictures are flattened to images; a few transitions (flip, cube, vortex, flash, morph…) map to the nearest standard PowerPoint transition; comments, sections and notes/handout masters are not exported; importing .pptx does not read animations back.
- Excel shapes and sparklines are not written to .xlsx.
- Spell-check/thesaurus exists in Word only.

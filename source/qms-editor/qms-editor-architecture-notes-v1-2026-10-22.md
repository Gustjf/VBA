# QMS Editor: System Architecture

**What it is:** a Windows desktop application that edits Doosan Bobcat North America QMS process documents and turns them into finished, branded PDFs.
**Version:** 2.0, 22 Oct 2026
**Owner:** Gust Faragallah, Manufacturing Engineer II
**Proof of concept:** QMS-022, ECN Factory Implementation Process

---

## 1. Summary

| Item | Description |
|---|---|
| Purpose | Edit QMS process documents in a simple form and produce a standard, Doosan Bobcat-branded PDF without touching LaTeX or Word formatting |
| Input | One `.tex` file for a QMS document, or its `<name>_source.zip` (the `.tex` plus its pictures), taken from a SharePoint folder or from the PC |
| Output | `<name>.pdf` (the final document) and `<name>_source.zip` (the `.tex`, its process-map pictures and the files needed to rebuild it) |
| Delivery | One portable file, `QMS Editor.exe`. Double-click to run; a setup wizard runs the first time. No installer and no administrator rights |
| Where documents live | Outside the app, in a SharePoint folder. The app stores no documents |
| Document control | Handled by the SharePoint document-control system: reviewer, review cycle and overdue reminders. The app records only who changed the document (Revised by and Job Title) |

---

## 2. Inputs and outputs

### 2.1 Inputs

| Input | Source | How it gets into the app |
|---|---|---|
| Existing document (`QMS-0NN.tex`) | SharePoint document folder or the PC | **Upload Current Document (.tex)**: pick or drag in the file, or use **Pick from the SharePoint folder** to see the folder's file list |
| Existing document with pictures (`QMS-0NN_source.zip`) | Same | Same buttons. The zip brings the `.tex` and its process-map pictures |
| New document | Standard template built into the app | **Create a new document**: enter the document ID, title and owner |
| Process maps and figures | Exported from Visio, Mermaid, PowerPoint or captured from the Word originals | **Choose picture…** in an appendix or figure. Accepts PNG, JPG, PDF or SVG |
| Name and job title of the person revising | Company directory (Active Directory) or the user's SharePoint profile; can be typed over | Filled in automatically during setup and when the app starts; editable in Settings and in the Finish dialog |
| Text changes | The person editing | A form laid out in PDF order: Header, Purpose, Turtle diagram, Sections and tables, Appendices, Revision History |

### 2.2 Outputs

| Output | Contents | Where it goes |
|---|---|---|
| `QMS-0NN.pdf` | The finished, branded document: header with the Doosan and Bobcat logos, turtle diagram, sections, Revision History, appendices with process maps, and the controlled-document footer | **Save to this PC** (folder picker) or **Upload to SharePoint** (asks before replacing an existing file) |
| `QMS-0NN_source.zip` | `QMS-0NN.tex`, `maps/*.png/.jpg/.pdf`, `doosanqms.cls`, `turtle.tex`, `brand/doosan-bobcat-logo.png` | Same places. Open this zip next time to keep editing |
| Revision History row | Rev. Date, Revised by, Job Title and Summary of Changes, added when the user clicks **Finish & create files** | Inside the PDF and the `.tex` |

### 2.3 What a `.tex` file holds

Each `.tex` the app writes is a complete LaTeX document, so any TeX engine can rebuild the PDF. After `\end{document}` the app also stores the document's data as commented lines between `%QMS-DATA-BEGIN` and `%QMS-DATA-END`; TeX ignores them. When the file is opened again, the app reads that block back exactly, so no information is lost between edits. Older `.tex` files from the first editor have no data block; the app rebuilds the document data from the TeX itself.

---

## 3. Architecture

### 3.1 Layers

| Layer | Technology | Responsibility |
|---|---|---|
| User interface | HTML, CSS and JavaScript with no framework, shown in Microsoft Edge WebView2 | Setup wizard, Home screen (two main buttons plus the Input and Output panels), editor (outline, form, PDF preview), Finish dialog |
| Core engine | JavaScript (`core.js`) | Turns document data into `.tex` and reads `.tex` back into data, checks documents and writes plain-English errors and warnings, sizes table columns, lists field changes |
| Desktop shell | Rust with Tauri 2 (`main.rs`) | Runs the PDF engine, builds and reads zip files, saves files, looks up the job title in the company directory, opens the SharePoint sign-in window |
| PDF engine | Tectonic 0.15.0 (XeTeX), embedded in the `.exe` | Compiles the `.tex` into a PDF |
| Document design | LaTeX class `doosanqms.cls`, `turtle.tex` (TikZ), template `process-standard.tex` | Page layout, branding, header and footer, turtle diagram, tables, appendix pages |
| Storage | SharePoint folder or files on the PC | All documents. The app keeps only settings and a package cache on the PC |

### 3.2 Data flow

```
SharePoint folder / PC
   │  QMS-0NN.tex  or  QMS-0NN_source.zip
   ▼
[Upload Current Document] or [Pick from SharePoint] or [Create a new document]
   ▼
Input panel (documents loaded in this session) ──► click Edit
   ▼
Editor form (outline | form | PDF preview)
   │  live checks: errors and warnings shown at the field
   │  Build preview: core.js writes the .tex ──► Tectonic ──► PDF preview
   ▼
Finish & create files
   │  adds a Revision History row (Rev. Date, Revised by, Job Title, Summary)
   │  writes the .tex ──► Tectonic ──► QMS-0NN.pdf
   │  zips the .tex + pictures + class files ──► QMS-0NN_source.zip
   ▼
Output panel ──► Save to this PC  or  Upload to SharePoint
   ▼
SharePoint document-control system (reviewer, review cycle, reminders)
```

### 3.3 Main components

| Component | File | Lines | Role |
|---|---|---|---|
| Screens | `ui/app.js` | about 1,670 | Wizard, Home, Input and Output panels, editor form, Finish dialog, Settings |
| Core engine | `ui/core.js` | about 1,270 | `.tex` writer and reader, checks, change list, table sizing |
| Bridge | `ui/host.js` | about 320 | Calls to the desktop shell; SharePoint REST calls run inside the sign-in window |
| Styles | `ui/app.css` | about 310 | Look and layout in Doosan Endeavour Blue |
| Font widths | `ui/fontwidths.js` | data | Carlito (Calibri-compatible) character widths, used to size table columns |
| Desktop shell | `src-tauri/src/main.rs` | about 530 | Tectonic runner, zip, files, directory lookup, SharePoint window |
| Document class | `ui/templates/doosanqms.cls` | about 255 | US Letter page, 0.6 in margins, branded header and footer, section bars, tables, appendices |
| Turtle diagram | `ui/templates/turtle.tex` | about 180 | Symmetric layout: rows of equal height, Inputs and Outputs centred on the circle, straight flow arrows |
| Page template | `ui/templates/process-standard.tex` | about 175 | Order of the PDF: banner and purpose, turtle, sections, Revision History, appendices |
| Rules | `ui/templates/process-standard.schema.json` | data | Shape of a document's data (JSON Schema), version 2 |
| New-document starter | `ui/templates/new-document.json` | data | Standard banner, purpose text, usual sections and the placeholder appendix |
| Build script | `BUILD QMS EDITOR.bat` and `build.ps1` | — | One-click build of `QMS Editor.exe` |

### 3.4 Data kept on the PC (not documents)

| Folder | Contents |
|---|---|
| `%LOCALAPPDATA%\QMS-Editor\` | Settings (name, job title, SharePoint folders), the PDF engine, temporary build folders |
| `%LOCALAPPDATA%\QMS-TeX\` | LaTeX package cache, downloaded once per PC (about 100–150 MB) |

---

## 4. Languages and tools used to build it

| Area | Language or tool | Version | Licence and cost |
|---|---|---|---|
| User interface | HTML5, CSS3, JavaScript (ES5/ES2015), no framework | — | Project code, no cost |
| Desktop shell | Rust | stable (1.99 used for the build) | MIT/Apache-2.0, free |
| App framework | Tauri | 2.x | MIT/Apache-2.0, free |
| Rust libraries | serde, serde_json, base64, zip, tauri-plugin-dialog | current | MIT/Apache-2.0, free |
| Web view | Microsoft Edge WebView2 | built into Windows 10/11 | Included with Windows |
| PDF engine | Tectonic (XeTeX plus TeX Live packages) | 0.15.0 | MIT, plus LPPL and other free licences for packages; free |
| Document design | LaTeX: fontspec, TikZ, longtable, tabularx, fancyhdr, hyperref, pdflscape and others | TeX Live via Tectonic | LPPL, free |
| Fonts | Calibri (Windows), with Carlito as a bundled fallback that has the same character widths | Carlito 1.104 | Calibri already licensed with Windows/Office; Carlito SIL OFL 1.1, free |
| Build | PowerShell 5.1, Windows batch file, Cargo | Windows built-in | Free |
| C++ linker for the build | Microsoft Visual Studio 2022 Build Tools (C++ workload) | 2022 | Needed only on the one PC that builds the app. Licence terms to be confirmed with IT (normally covered by the company's Visual Studio licensing) |
| Tests (development only) | Node.js, Playwright with Chromium | — | Free |
| Process-map capture from the Word originals (one-time conversion) | Python, LibreOffice (headless), poppler `pdftoppm`; connector-line fix `fix_connectors.py` | — | Free |

---

## 5. Tests

All tests are in `qms-editor\tests\` and run during development; they are not part of the `.exe`.

| Test | What it checks | Result |
|---|---|---|
| `parity.js` | The JavaScript `.tex` writer produces character-for-character the same output as the original Python builder for every document | 32 of 32 identical |
| `validate-parity.js` | The JavaScript checks give the same errors and warnings as the original Python checker | 32 of 32 documents, plus 16 of 16 deliberately broken documents |
| `roundtrip.js` | A document written to `.tex` and read back is unchanged; older `.tex` files read back and print the same | Data block 32 of 32 exact; older files 27 of 30 field-exact and 30 of 30 printing identically |
| `ui-test.js` | End to end: setup wizard (name and job title from the directory, PDF engine, SharePoint folder), upload, Input list, edit, preview, Finish (Revision History with Revised by and Job Title, no "Reviewed by"), save as PDF and zip, upload to SharePoint, older file, pick from SharePoint, create a new document, no "Rev date is empty" warning | All checks pass (simulated desktop shell and SharePoint) |
| `scroll-test.js` | 12 inputs and 8 outputs stay reachable with the mouse wheel at 1440×920, 1280×720 and 1024×640 | Passes; it also fails on the old version, so it catches the original bug |
| `picture-test.js` | Appendix pictures: **Choose picture…** on placeholders, SVG converted to PNG, PDF pictures with a preview, both pictures go into the PDF build | Passes |
| Zip import check | `QMS-022_source.zip` imports with 2 pictures, no warnings, and both maps reach the PDF build | Passes |
| Independent code reviews | Tauri 2, zip and base64 API use; SharePoint sign-in detection | Issues found were fixed |
| On the user's PC | Real build, PDF engine, directory title lookup, turtle diagram look | Confirmed by the user (app builds and runs, title lookup works, turtle approved) |

---

## 6. Design requirements and how they were met

### Requirement 1: No cost to implement and create documents

| How it is met | Detail |
|---|---|
| Only free, open-licence components | Tectonic (MIT), TeX Live packages (LPPL), Tauri and Rust (MIT/Apache), Carlito font (SIL OFL). No subscriptions, cloud services, AI services or paid add-ins |
| Uses what Doosan already pays for | Windows (WebView2, PowerShell), Microsoft 365 (SharePoint storage, sign-in, document control), Calibri font |
| No admin rights or IT installs for users | One portable `.exe`. The PDF engine unpacks into the user's own profile; the setup wizard does the rest |
| No per-document cost | Creating, editing and publishing a document costs nothing beyond staff time |
| One open point | The free C++ Build Tools are needed only on the single PC that builds the `.exe`; confirm the licence with IT |

### Requirement 2: All final versions of documents are PDFs

| How it is met | Detail |
|---|---|
| PDF is the only finished output | **Finish & create files** always produces `<name>.pdf`. The `.tex` and pictures go into `<name>_source.zip` for future editing only |
| Same look every time | One LaTeX class and one template set the layout, fonts, colours, header, footer and Revision History for every document |
| Controlled-document elements built in | Header (Title, Owner, Location, Section, Rev Date, Page x of y), the three footer statements, and the Revision History on every PDF |
| Branding | Header logo combining the new Doosan logo (Endeavour Blue) and the current Bobcat logo; accent colour Endeavour Blue #005EB8 |
| Process maps print exactly as supplied | Pictures are placed at full resolution and never redrawn; PDF pictures stay vector-sharp |

### Requirement 3: Updating documents must be time-effective

| How it is met | Detail |
|---|---|
| Two clear starting points | **Upload Current Document (.tex)** and **Create a new document** on the Home screen |
| Edit text, not formatting | A form in PDF order. The turtle diagram is typed as one line per bullet; tables are rows of cards; the layout is automatic |
| Instant feedback | Checks run while typing (errors and warnings at the field); **Build preview** shows the PDF in a few seconds; the preview jumps to the page being edited |
| Automatic bookkeeping | The Revision History row and the header Rev Date are filled in on Finish; name and job title come from the company directory |
| Fast hand-off | Save the PDF and zip to the PC, or upload both to SharePoint, in one click each |
| Several documents per session | The Input list holds every loaded document; the Output list holds every finished set of files |
| Review handled elsewhere | Reviewer, review cycle and reminders are left to the SharePoint document-control system, so the editor has no extra steps |

### Requirement 4: The process must be repeatable and scalable

| How it is met | Detail |
|---|---|
| One template for every document | All 30 North America QMS documents use the same template, class and rules; 28 of the 30 follow the QMS-022 pattern directly |
| Any number of documents | No documents are stored in the app; each `.tex` or zip is self-contained, so the library grows in SharePoint, not in the app |
| Anyone can follow the process | Non-technical users only click **Upload → Edit → Finish → Save/Upload**. The first-start wizard replaces the old `.bat` files |
| Easy to share | One `.exe` to give to another person or function; a guided setup on each PC |
| Self-contained documents | A `.tex` carries its own data; a `_source.zip` also carries its pictures and the class files, so it can be rebuilt with any TeX engine, even without the app |
| Same result every time | Automated tests confirm identical output, exact round trips and stable checks across all 30 documents |
| Batch conversion ready | The 30 Word originals are already converted to `.tex` in `source\QMS documents (.tex)`. Process maps can be captured from the Word originals in one batch with the connector-line fix (`fix_connectors.py`) |
| Room to extend | The data format is versioned (schema 2) and older files are upgraded automatically. New pictures from Visio, Mermaid and similar tools work through **Choose picture…** |

---

## 7. Other design requirements from the project brief

| Requirement | How it was met |
|---|---|
| Scale back to a `.tex` editor; remove the document library | The library, drafts, logs and published PDFs moved to the `source` backup folder; the app now holds only the template and what it needs to run |
| Input is a `.tex`; output is a `.tex` and a PDF | `.tex` (or `_source.zip`) in; `<name>.pdf` and `<name>_source.zip` out |
| Two obvious buttons: "Upload Current Document (.tex)" and "Create a new document" | Two large buttons at the top of the Home screen |
| Input/Output window instead of the library | **Input** panel lists the files loaded in this session; **Output** panel lists the finished files ready to save |
| SharePoint: manual upload and folder link; download or upload output | Both supported. **Pick from the SharePoint folder** lists the folder's `.tex` and `_source.zip` files; **Upload to SharePoint** asks before replacing a file. Users sign in with their own Microsoft account |
| Separate saved documents from the app; keep a backup | Everything moved to `TeX\source` (old library, drafts, logs, published PDFs, Word originals, old templates, old HTML app) |
| Tauri app for Windows; one file; setup wizard; non-technical audience | `QMS Editor.exe` with a five-step first-start wizard: Welcome, Name and job title, PDF engine, SharePoint, Done |
| New Doosan logo (Endeavour Blue); keep the Bobcat logo | Combined Doosan and Bobcat header logo in the PDF and the app |
| Review outside the program; record only who changed the document | "Reviewed by" removed; "Rev. By" renamed "Revised by"; Job Title column added (auto-filled, editable) |
| No "Rev date is empty" warning | Removed; the Rev date is filled in automatically on Finish |
| Professional turtle diagram | Symmetric layout: rows of equal height, Inputs and Outputs centred on the circle, straight horizontal flow arrows, mirrored corner lines |
| Long Input and Output lists | The Home screen scrolls; every item stays reachable |
| Don't jump straight into the editor | Uploaded and new documents land in the Input list; the user clicks **Edit** |

---

## 8. Proof of concept: QMS-022

| Item | Detail |
|---|---|
| File | `QMS-022_source.zip`, saved in the TeX folder |
| Contents | `QMS-022.tex`; `maps/QMS-022_A.png` (Quality Change Notification Process, from page 2 of the Word original); `maps/QMS-022_B.png` (Engineer Change Notification, from page 3); `doosanqms.cls`; `turtle.tex`; `brand/doosan-bobcat-logo.png` |
| Process maps | Captured from the Word original at 300 dpi, not redrawn; connector lines follow the same routes as in Word |
| Corrections | "Notificaiton" corrected to "Notification" |
| Revision History | Row dated 22 Oct 2026: Gust Faragallah, Manufacturing Engineer II, "Converted to the QMS Editor format; process maps added as Appendix A and Appendix B; corrected the spelling of "Notification"" |
| Checks | 0 errors, 0 warnings; reopens exactly; both pictures included in the PDF build |

---

## 9. Folder layout (`TeX` folder in OneDrive)

```
TeX\
├── QMS Editor.exe                  the app (made by the build)
├── QMS-022_source.zip              proof-of-concept document
├── README.md                       user guide
├── qms-editor\                     app source code and BUILD QMS EDITOR.bat
│   ├── ui\                         screens, core engine, templates, brand
│   ├── src-tauri\                  desktop shell (Rust), icons, fonts, permissions
│   └── tests\                      development tests
└── source\                         backup of everything separated from the app
    ├── QMS documents (.tex)\       the 30 documents in the new format
    ├── TeX source data\            the 30 Word originals
    ├── content\ build\ published\ log\ maps\ templates\ schema\
    ├── legacy-html-app\            the first (HTML) editor and its tools
    └── project-notes\              roadmap, file descriptions, brainstorming, fix_connectors.py
```

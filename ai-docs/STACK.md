# STACK.md — Tech Stack, Folder Structure & Conventions

Last updated: 2026-04-09

---

## 1. Tech Stack

| Layer              | Technology           | Version / Notes                               |
| ------------------ | -------------------- | --------------------------------------------- |
| Language           | Python 3             | Windows; runs as .py or .exe                  |
| GUI                | tkinter + ttk        | Desktop app, tabbed interface                 |
| Browser Automation | Selenium WebDriver   | v4.15.2, Microsoft Edge                       |
| Edge Driver        | webdriver-manager    | Auto-managed via msedgedriver                 |
| Excel I/O          | openpyxl             | ≥3.1.2                                        |
| PDF extraction     | pdfplumber           | ≥0.10.0 — digital PDFs                        |
| PDF fallback OCR   | ocrmypdf + tesseract | Image-based PDFs                              |
| AI name matching   | google-genai         | ≥1.0.0, model `gemini-3.1-flash-lite-preview` |
| Fuzzy matching     | rapidfuzz            | ≥3.5.0                                        |
| Environment vars   | python-dotenv        | ≥1.0.0                                        |
| Packaging          | PyInstaller          | ≥6.0.0 → `MedAfrica.exe`                      |
| System utils       | psutil               | ≥5.9.5                                        |

> **IMPORTANT**: `google-generativeai` is deprecated. Always use `google-genai` (Gemini 3.x API).

---

## 2. Environment Variables (`.env`)

```
EDGE_PATH=<path to msedge.exe>
DRIVER_PATH=<path to msedgedriver.exe>
BADR_PASSWORD=<badr operator password>
```

Optional (for Gemini AI):

```
GEMINI_API_KEY=<api key>
```

---

## 3. Configuration Files (`config/`)

| File                               | Purpose                                                |
| ---------------------------------- | ------------------------------------------------------ |
| `config/settings.json`             | GUI defaults: window size, last folder, username       |
| `config/license.json`              | License expiry metadata (`LTA_sys_ts`, `LTA_validity`) |
| `config/email_config.json`         | SMTP options for email sender                          |
| `config/email_config.json.example` | Template for email config                              |

License is validated at startup (`gui/utils/license_validator.py`). App exits if expired.

---

## 4. Folder Structure

```
Application_auto_ds/
├── badr_login_test.py          # Main automation engine (~11,900+ lines)
├── script_all_fuzy_match.py    # Preparation/preprocessing pipeline
├── known_companies.json        # FROZEN legacy fuzzy-match list (39 names) — never written by the app
├── known_companies_shared.json # shared names, curated by the maintainer (tracked)
├── known_companies_local.json  # names learned on THIS PC (untracked, only file the app writes)
├── result_LTAS.txt             # Append-only DUM reference output log
├── requirements.txt            # pip dependencies
├── build_exe.py                # PyInstaller build helper
├── MedAfrica.spec              # PyInstaller spec file
├── .env                        # Environment vars (not committed)
├── antigravity_instructions.md # Architectural design overview
├── FULL_APP_CONTEXT_FOR_AI.md  # Full AI briefing document
├── context.md                  # System overview + workflow diagrams
│
├── config/                     # App config files
│   ├── settings.json
│   ├── license.json
│   └── email_config.json(.example)
│
├── gui/                        # tkinter GUI package
│   ├── main.py                 # Entry point (also does git pull auto-update)
│   ├── app.py                  # BADRApp: root window, tabs, state
│   ├── __init__.py
│   ├── assets/                 # Icons, images
│   ├── screens/
│   │   ├── preparation.py      # Tab 1: folder select, prep script, DS/shipper edit
│   │   ├── phase1_ed.py        # Tab 2: ED creation
│   │   ├── phase2_dum.py       # Tab 3: DUM declaration
│   │   ├── logs.py             # Tab 4: log viewer
│   │   └── partial_config_dialog.py  # Dialog for partial LTA split config
│   └── utils/
│       ├── script_manager.py   # Subprocess runner for .py scripts → streams to logs
│       ├── file_utils.py       # LTA file I/O helpers (read/write txt, detect LTAs)
│       ├── validators.py       # DS series, location, folder validation
│       ├── email_sender.py     # SMTP email utility
│       ├── license_validator.py # License check at startup
│       └── theme.py            # tkinter theme helpers, footer
│
├── <X>eme LTA/                 # LTA working folders (e.g. "10eme LTA/")
│   ├── <MAWB>.pdf              # Master air waybill PDF
│   ├── Sheet <N>.xlsx          # Per-DUM cargo Excel files
│   ├── generated_excel*.xlsx   # Summary workbook (Summary sheet + Sheet N sheets)
│   └── <folder>_partial_config.json  # (if partial LTA)
│
├── <X>eme_LTA_shipper_name.txt # Shipper + DS info (one per LTA, in root folder)
├── <X>eme LTA.txt              # LTA reference data file (in root folder)
│
├── inspect_excel.py             # Standalone utility: inspect Excel structure (not GUI-integrated)
├── validation.py                # Standalone utility: manual validation runner (not GUI-integrated)
├── badr_connection.py           # Standalone Edge connection helper (legacy; logic duplicated in badr_login_test.py)
├── test_integration.py          # Ad-hoc integration test script
├── test_distribution.py         # Ad-hoc distribution test script
├── badr_login_test_logs.txt     # Leftover debug log from manual runs
├── result_LTAS.txt              # Append-only DUM reference output log (no archival mechanism)
│
├── process-return-colis-script/ # Separate sub-project: return parcel processing (out of scope)
└── retour-waybill-app/          # Separate sub-project: waybill web app (out of scope)
```

### Dead Code in `gui/screens/`

| File                                       | Status                                                         |
| ------------------------------------------ | -------------------------------------------------------------- |
| `partial_config_dialog_check_first_dum.py` | Old iteration of partial config dialog — not imported anywhere |
| `partial_config_dialog_old_logic.py`       | Previous logic version — not imported anywhere                 |

Both files are superseded by `partial_config_dialog.py` and can be safely deleted.

---

## 5. Key File Naming Conventions

| Pattern                        | Meaning                                           |
| ------------------------------ | ------------------------------------------------- |
| `<N>eme LTA/`                  | LTA folder (ordinal: 1er, 2eme, ..., 10eme, etc.) |
| `<N>eme_LTA_shipper_name.txt`  | Shipper data file for LTA N (in project root)     |
| `<N>eme LTA.txt`               | LTA reference file (in project root)              |
| `<folder>_partial_config.json` | Partial split config inside LTA folder            |
| `generated_excel*.xlsx`        | Summary + sheet data workbook                     |
| `Sheet <N>.xlsx`               | Single DUM cargo sheet                            |

Sorting: LTA folders sort by numeric ordinal using `_extract_lta_number()` — NOT lexicographic. `1er`, `2eme`, `10eme` sort correctly (not `10` before `2`).

---

## 6. Script Execution Model

`script_manager.py` runs automation scripts as **subprocesses** and streams their stdout/stderr into the GUI Logs tab in real time.

Invocation signatures:

```
python badr_login_test.py 1 all          # phase 1, all LTAs
python badr_login_test.py 1 0,2,4        # phase 1, specific indices
python badr_login_test.py 2 all          # phase 2, all LTAs
python badr_login_test.py 2 1,3          # phase 2, specific indices
python script_all_fuzy_match.py          # preparation, no args
```

For selective LTA processing, unselected LTA folders are temporarily moved to `.temp_unselected/` inside the working directory, then restored after the script finishes.

---

## 7. Browser Automation Conventions

- **Driver**: Microsoft Edge via Selenium 4
- **Connection**: Edge started in debug mode; `connect_to_edge()` attaches
- **Waiting**: `WebDriverWait` + `expected_conditions` — no `time.sleep` patterns
- **Iframe pattern**: Switch to iframe for forms → `driver.switch_to.default_content()` when done
- **JS fallback**: `driver.execute_script("arguments[0].click()", elem)` when normal `.click()` fails
- **Home reset**: `return_to_home_after_error(driver)` navigates to BADR home safely
- **Profile management**: Fresh Edge profile per session (`get_fresh_profile_path()`)

---

## 8. Excel Data Conventions

`read_dum_data_from_summary(summary_excel_path)` reads the **Summary** sheet from `generated_excel*.xlsx`:

- Returns list of DUM dicts with: `dum_number`, `weight`, `positions`, `serie` (initially empty), etc.

`save_dum_series_to_excel(lta_folder_path, dum_number, serie)` writes DUM series back to the same workbook.

---

## 9. GUI Architecture

```
BADRApp (gui/app.py)
├── self.current_folder     — currently selected working folder
├── self.lta_list           — detected LTA folder names
├── PreparationScreen       — Tab 1
├── Phase1EDScreen          — Tab 2
├── Phase2DUMScreen         — Tab 3
└── LogsScreen              — Tab 4
```

Tabs are all enabled from start. Switching tabs auto-refreshes LTA list. Auto-update from GitHub runs silently at startup via `git pull --autostash origin main`.

---

## 10. Build / Distribution

```bash
python build_exe.py    # Runs PyInstaller → dist/MedAfrica.exe
```

Spec file: `MedAfrica.spec`  
Output: `build/MedAfrica/` + `dist/MedAfrica.exe`

The .exe bundles all Python dependencies. EDGE_PATH and DRIVER_PATH must still be user-configured via `.env` on the target machine.

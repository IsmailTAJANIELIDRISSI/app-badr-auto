# PROJECT.md — Architecture, Business Logic & Data Flow

Last updated: 2026-04-09

---

## 1. What This Project Does

**MedAfrica Logistics — BADR Automation** automates Morocco's customs declaration workflows on the BADR platform (`badr.douane.gov.ma:40444`).

Each air freight shipment arrives as an **LTA (Lettre de Transport Aérien — air waybill)** folder. The system:

1. Validates and prepares the raw LTA folder files (PDFs, Excel)
2. Creates an **Etat de Dépotage (ED)** on BADR — the unloading statement
3. Creates one or more **DUM declarations** — the customs declarations per LTA
4. Optionally performs **DS préapurement** — pre-clearance of the DS bond document

All of this is wrapped in a desktop GUI (`MedAfrica.exe`) that operators run locally on Windows.

---

## 2. Business Concepts

| Term             | Meaning                                                       |
| ---------------- | ------------------------------------------------------------- |
| **LTA**          | Air waybill — a shipment folder (e.g. `10eme LTA/`)           |
| **ED**           | Etat de Dépotage — unloading statement created in BADR        |
| **DUM**          | Déclaration Unique de Marchandises — customs declaration      |
| **DS**           | Bond document series (e.g. `9913 G`, `1510W`) used in ED lots |
| **Lot**          | A line inside an ED, linked to one DUM + one DS reference     |
| **MAWB**         | Master Air Waybill number (e.g. `607-50843822`)               |
| **Préapurement** | Pre-clearance step that verifies DS bond against DUM weights  |
| **Blocage**      | A held/blocked LTA requiring modified ED weights              |
| **Partial LTA**  | An LTA split across multiple flights / partial deliveries     |
| **DS MEAD**      | Document type 03 — used in exception partial smallest lot     |
| **DS Dépotage**  | Document type 05 — used in standard ED creation               |
| **Série**        | The DS series string (format: `NNNN X`)                       |
| **Clé**          | The DS key/checksum suffix                                    |

---

## 3. LTA Types and Their ED Behavior

### Standard LTA

- One ED covering all DUMs
- DS series from `*_shipper_name.txt` line 2
- Validated DS reference saved back to shipper txt line 4

### Blocage LTA

- Detected via `*_LTA.txt` blocage markers (line 5 / 12 / 13)
- Corrects weights before ED creation (`correct_blocage_weights`)
- Calls `modify_etat_depotage_for_blocage` instead of standard create

### Partial LTA

- Split into N partials configured in `*_partial_config.json`
- Each partial has its own DS series/cle/location/weight
- Calls `create_etat_depotage_partial` per partial
- Saves `ds_validated` per partial back into partial JSON

### Partial Exception Case

- Triggered when: `smallest_partial_weight < max(all DUM weights)`
- The **smallest partial** is treated as airport-cleared complement
- The **split DUM** straddles two partials:
  - Lot A → airport reference + DS MEAD (type 03)
  - Lot B → depotage-side split reference + DS Dépotage (type 05)
- Exception split DUM detected dynamically from `split_dums` keys (NOT hardcoded to DUM 1)

### Blocage + Partial LTA (both at once)

- The **largest partial** (by weight) is merged with the blocked waybills: its ED already
  exists (created by the agent with the blocked items), so it is **modified**, not created —
  via `process_blocage_partial_ed` → `modify_etat_depotage_for_blocage(..., partial_config, partial_number)`
  using that partial's DS série. LTA-ref lots deleted, `MA…` blocked lots kept, lots added
  for **only that partial's DUMs**.
- Its DUMs total `partial weight − blocked weight` (the partial dialog deducts it; Phase 1
  refuses a config that doesn't). Saved in config as `blocage_partial_number` / `blocage_weight`.
- Lot refs use the **partial** convention (`ref/N`, `ref/N/k`) so Phase 2 handles it like any
  other partial; `ds_validated` saved to the partial config = checkpoint.
- Blocage detected from line 5 of the LTA txt **or** from `blocage_weight` in the partial config
  (`resolve_blocage_for_partial`). DS of the blocage = DS of the largest partial (Préparation DS optional).
- **Quantités tab** (MED.5c): `Nombre total Contenant(s)` = **"Nb Articles Bloqués"** typed on
  the Préparation card (LTA txt **line 14**, 1 article = 1 contenant; required for blocage +
  partial) + the partial's DUM lot positions. The remaining lots of the ED (= blocked articles,
  `read_ed_lots_table`) are a cross-check / fallback when line 14 is empty; and
  `Poids brut total` = partial weight — corrected by `fix_ed_quantites` only if wrong.
  Safety net: on E2800435 at VALIDER, BADR's computed count is declared (if coherent) and the ED re-validated once.
- Other partials unchanged: exception (smallest) → DS MEAD, no ED; others → ED created.

---

## 4. Data Flow

```
┌─────────────────────────────────────────────────────────┐
│  GUI: PreparationScreen                                  │
│  → runs script_all_fuzy_match.py as subprocess          │
└──────────────────────┬──────────────────────────────────┘
                       ↓
┌─────────────────────────────────────────────────────────┐
│  script_all_fuzy_match.py (Preparation Phase)           │
│  Input: raw LTA folders with PDFs + Excel               │
│  Output: validated folder with:                         │
│    - *_shipper_name.txt  (shipper + DS info)            │
│    - Renamed MAWB PDFs (standardized names)             │
│    - generated_excel*.xlsx (Summary + Sheet N)          │
│    - Warning/error txt files if validation fails        │
└──────────────────────┬──────────────────────────────────┘
                       ↓
┌─────────────────────────────────────────────────────────┐
│  GUI: PreparationScreen (post-preparation)              │
│  - Editable table: shipper, DS serie, location, blocage │
│  - Partial config dialog (for split LTAs)               │
│  - Saves changes to *_shipper_name.txt and partial JSON │
└──────────────────────┬──────────────────────────────────┘
                       ↓
┌─────────────────────────────────────────────────────────┐
│  GUI: Phase1EDScreen                                     │
│  → runs badr_login_test.py 1 <indices>                  │
│  → creates ED per LTA in BADR                           │
│  Output: DS validated reference saved to shipper txt    │
│          or partial JSON (line 4 / ds_validated)        │
└──────────────────────┬──────────────────────────────────┘
                       ↓
┌─────────────────────────────────────────────────────────┐
│  GUI: Phase2DUMScreen                                    │
│  → runs badr_login_test.py 2 <indices>                  │
│  → creates DUM declarations in BADR                     │
│  → uploads supporting documents                         │
│  → performs DS préapurement lots                        │
│  Output: DUM references appended to result_LTAS.txt     │
│          DUM series saved to generated_excel             │
└─────────────────────────────────────────────────────────┘
```

---

## 5. File Contracts (CRITICAL — never break)

### `<X>eme LTA.txt` (primary LTA data file)

- Line 4: LTA reference (MAWB)
- Line 6: Shipper name
- Line 8: Signed series (if signed)
- Lines 5/12/13: Blocage metadata

### `<X>eme_LTA_shipper_name.txt` (shipper/DS data)

- Line 1: Shipper name
- Line 2: DS series (e.g. `9913 G`) — optional
- Line 3: Loading location — optional
- Line 4: DS validated reference (written by phase 1 after ED creation)
- Line 5: Validated LTA reference (written by phase 1)

### `<folder>_partial_config.json` (partial config)

```json
{
  "partial_type": "normal|exception",
  "partials": [
    {
      "partial_number": 1,
      "weight": 500.0,
      "positions": 10,
      "ds_serie": "9913 G",
      "ds_cle": "X",
      "loading_location": "CMN",
      "dums": [
        {
          "dum_number": 1,
          "weight": 200,
          "positions": 4,
          "is_split": false,
          "split_id": ""
        }
      ],
      "ds_validated": "1510W",
      "signed_series": "optional"
    }
  ],
  "split_dums": {
    "2": { "split_id": "A", "partial_1": 2, "partial_2": 1 }
  },
  "smallest_partial_number": 2,
  "smallest_partial_positions": 5,
  "smallest_partial_airport_reference": "607-12345678/2"
}
```

---

## 6. Key Functions in badr_login_test.py

| Function                              | Purpose                                                             |
| ------------------------------------- | ------------------------------------------------------------------- |
| `process_lta_folder_ed_only`          | Phase 1 dispatcher — detects type and calls appropriate ED function |
| `process_lta_folder_dum_only`         | Phase 2 dispatcher — creates DUMs for an LTA                        |
| `process_lta_folder`                  | Combined phase 1+2 (legacy)                                         |
| `create_etat_depotage`                | Standard ED creation                                                |
| `create_etat_depotage_partial`        | Partial ED creation (called per partial)                            |
| `modify_etat_depotage_for_blocage`    | Opens existing ED and edits for blocage case                        |
| `get_dum_preapurement_lots`           | Computes DS lot entries for DUM préapurement                        |
| `fill_declaration_form`               | Phase 2 — fills entire DUM form, uploads, validates                 |
| `detect_blocage_from_lta_file`        | Reads LTA txt to detect blocage flag                                |
| `read_shipper_from_txt`               | Reads shipper txt file (lines 1–5)                                  |
| `read_dum_data_from_summary`          | Reads generated_excel Summary sheet                                 |
| `save_ds_validated_to_partial_config` | Writes ds_validated to partial JSON                                 |
| `verify_uploaded_files`               | Checks BADR DOM table for uploaded doc presence                     |
| `return_to_home_after_error`          | Navigates browser back to BADR home                                 |
| `find_lta_folders`                    | Scans directory for LTA subfolders                                  |

---

## 7. Phase 1 ED Logic Detail

```
process_lta_folder_ed_only(driver, lta_folder_path, lta_name)
  ├── detect_blocage_from_lta_file() + get_lta_partial_info()
  ├── if blocage AND no partial_config → modify_etat_depotage_for_blocage()
  ├── elif partial_config exists (blocage or not):
  │     for each partial:
  │       exception (smallest) partial → skip (DS MEAD)
  │       ds_validated/signed_series set → skip (checkpoint)
  │       blocage AND partial == largest → process_blocage_partial_ed() (MODIFY existing ED)
  │       else → create_etat_depotage_partial(driver, path, config, partial_num)
  └── else (standard):
        MAX_ED_RESTARTS = 2 outer loop
        create_etat_depotage(driver, path, shipper_data)
        on failure → return_to_home_after_error() → retry
```

Note: `ED_CREATE_BUSINESS_ERROR = object()` sentinel — returned when BADR returns error E2800124 (already exists). Does NOT trigger ED restart.

---

## 8. Phase 2 Préapurement Lot Logic

`get_dum_preapurement_lots(dum_number, partial_config, validated_lta_reference)` returns a list of lots.

- **Standard**: 1 lot = `<validated_ref>/<dum_number>`
- **Partial non-split DUM**: 1 lot, DS from partial's `signed_series` (fallback to `ds_serie`)
- **Partial split DUM**: multiple lots (one per `split_dums` entry), each with own DS + reference
- **Exception split DUM**:
  - Lot 1: airport reference + `is_exception_smallest=True` → DS MEAD type
  - Lot 2: `airport_ref/N` + `is_exception_largest=True` → DS Dépotage type
  - Exception split DUM detected dynamically: `exception_split_dums = list(split_dums.keys())`

---

## 9. Error Handling Philosophy

- **Per-DUM granularity**: one DUM failure does not stop the whole LTA
- **Per-LTA granularity**: one LTA failure does not stop the batch
- **Explicit error files**: written near LTA folder when critical failure occurs
- **Selenium retries**: JS click fallback, iframe reset, home navigation fallback
- **Error log functions**: `save_dum_error_log`, `mark_dum_as_error_in_excel`
- **MAX_ED_RESTARTS = 2**: ED creation retried up to 2 times after failure (navigation reset between attempts)

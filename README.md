# Tdh Kenya Helpdesk Data Dashboard

An interactive **Streamlit** dashboard for exploring Terre des hommes (Tdh) Kenya
protection-helpdesk records. It loads a cleaned Excel workbook, derives demographic
and disability indicators (including Washington Group Short Set scoring), strips
personally identifiable information, and presents filterable charts, cross-tabs,
a GPS map, caseworker (CPV) productivity, and a searchable record viewer.

> **Developed by John Kul, MEAL Officer – Tdh.** · *ImpactLens Africa — Turning Data Into Human Impact.*

---

## Table of contents
- [Features](#features)
- [Screens / tabs](#screens--tabs)
- [Project structure](#project-structure)
- [Requirements](#requirements)
- [Installation](#installation)
- [Adding your data and logo](#adding-your-data-and-logo)
- [Running the app](#running-the-app)
- [Input data contract](#input-data-contract)
- [How the analysis works](#how-the-analysis-works)
- [Filters & session behaviour](#filters--session-behaviour)
- [Privacy & safeguarding](#privacy--safeguarding)
- [Configuration reference](#configuration-reference)
- [Troubleshooting](#troubleshooting)
- [Deployment](#deployment)
- [Known notes](#known-notes)
- [License & attribution](#license--attribution)

---

## Features
- **KPI header** — valid records, protection concerns, information requests, partner
  referrals, follow-ups required, and disability records, each with a rate delta.
- **Quick-insight cards** — busiest helpdesk, top request type, top protection
  concern, top impairment type.
- **Eight sections** (Overview, Disability, Concerns, Information, Referrals, Map,
  CPV Work, Records).
- **Gender-disaggregated** charts and pivot tables throughout, with consistent colour
  coding and totals.
- **Washington Group (WGQ) disability analytics** for adults plus a separate child
  disability path, person-level prevalence, multiplicity, severity, and a
  reconciliation table.
- **GPS validation** — coordinates are parsed, range-checked, and geofenced against
  expected Turkana West and Dadaab operational areas, with an outlier-review table.
- **Cascading sidebar filters** (date → camp → helpdesk → demographics → request) that
  prune stale selections automatically.
- **PII stripped** from everything the dashboard displays or exports.
- **CSV export** of the filtered (and optionally searched) records.
- **Custom theme** via injected CSS (Tdh green / gold palette).
- **Cached loading** keyed on the data file's modification time, with a sidebar
  "Refresh data" button.

## Screens / tabs
| Tab | What it shows |
|-----|---------------|
| **Overview** | Monthly request trend by gender, request-type donut + table, demographic and location breakdowns by gender. |
| **Disability** | Overall status, impairment types, adult WGQ prevalence/severity/multiplicity, adult specific-impairment breakdown, reconciliation table, child disability. |
| **Concerns** | Top protection concerns by gender and by disability status (configurable top‑N). |
| **Information** | Top general-information needs by gender and by disability status. |
| **Referrals** | Referral status, follow-up requirement, and top referral partners by gender. |
| **Map** | `st.map` of valid GPS points, area classification, mapped points, and GPS outliers for review. |
| **CPV Work** | Per-caseworker (CPV) record counts, referrals, follow-ups, disability records, mapped records, helpdesk coverage, and active date span. |
| **Records** | Column picker, free-text search, paginated preview (first 1,000 rows), CSV download of all matching rows, and the source KPI summary. |

---

## Project structure
```
tdh-kenya-helpdesk-dashboard/
├── app.py                  # the full Streamlit application
├── requirements.txt        # pinned dependencies
├── README.md               # this file
├── .gitignore              # excludes data/secrets/caches from version control
├── assets/
│   ├── styles.css          # all dashboard styling (edit here, not in app.py)
│   └── README.txt          # drop tdh-logo.png / developer-logo.png here
└── data/
    └── README.txt          # drop HELPDESK_DashboardData_Tdh_Kenya_D2.xlsx here
```
The app resolves these paths relative to `app.py`:
- Tdh header logo → `assets/tdh-logo.png` (optional; the banner renders without it).
- Developer logo → `assets/developer-logo.png` (optional; shown in the footer card.
  Falls back to text-only if absent).
- **Stylesheet → `assets/styles.css`** — all dashboard styling lives here. Edit colours,
  spacing, card design, etc. without touching the Python. If the file is missing the app
  still runs with a minimal inline fallback.
- Data → `data/HELPDESK_DashboardData_Tdh_Kenya_D2.xlsx` (**required**).

### Customising the look
- Theme colours are CSS variables at the top of `assets/styles.css` (`--tdh-green`,
  `--tdh-gold`, etc.) — change them once to re-skin the whole dashboard.
- The app/version label shown in the footer is set by `APP_VERSION` and
  `APP_VERSION_DATE` near the top of `app.py`.

---

## Requirements
- **Python 3.9 – 3.12**
- Packages pinned in `requirements.txt`:
  - `streamlit==1.40.2`, `pandas==2.2.3`, `altair==5.4.1`, `openpyxl==3.1.5`

## Installation
```bash
# 1. (recommended) create and activate a virtual environment
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate

# 2. install dependencies
pip install -r requirements.txt
```

## Adding your data and logo
The ZIP ships with **placeholders only** — add your real files before running:

1. Place the workbook at **`data/HELPDESK_DashboardData_Tdh_Kenya_D2.xlsx`**.
   The app reads the first sheet automatically (or a sheet named `cleaned_data` if
   one exists). A sheet named **`Column Mapping`** is optional and used only for
   friendly labels.
2. (Optional) Place the Tdh header logo at **`assets/tdh-logo.png`** and/or the
   developer logo at **`assets/developer-logo.png`** (the developer logo appears in
   the footer).

If the filename differs, either rename your file or edit `DATA_FILE_PATH` near the top
of `app.py`.

## Running the app
```bash
streamlit run app.py
```
Streamlit opens the dashboard in your browser (default `http://localhost:8501`).
Use the **sidebar** to filter and the **section selector** at the top to switch tabs.

---

## Input data contract
The app reads your data **directly from the workbook**. You do **not** need a sheet
named `cleaned_data`: if a sheet with that name exists it is used, otherwise the app
falls back to the **first sheet** in the workbook (so a single-sheet workbook just
works). A sheet named **`Column Mapping`** is used for friendly labels only if it is
present. The app tolerates missing optional columns (it backfills them as empty).
Columns it reads or derives from include:

**Identity / context**
`interview_date`, `camp_location`, `household_type`, `staff_name`,
`gps_latitude`, `gps_longitude` (free-text coordinates are parsed), and optional
`helpdesk_camp_location`, `helpdesk_village`, `_GPS Location_longitude`, `referral_date`.

**Demographics**
`information_seeker_age`, `information_seeker_type`, `information_seeker_gender`.
Life stage (Child/Adult) is derived from the age band; type and gender are reconciled
to the derived life stage (with correction flags retained).

**Request / action**
`request_type_protection_or_information`, `action_taken`, `follow_up_required`.

**Disability — adults (WGQ Short Set)** — any of:
`difficulty_seeing`, `difficulty_hearing`, `difficulty_walking_or_climbing`,
`difficulty_walking_or_climbing_steps`, `difficulty_remembering_or_concentrating`,
`difficulty_self_care`, `difficulty_communicating`, plus
`information_seeker_disability_type_other`.

**Disability — children**
`has_disability`, `child_disability_type`, `child_disability_type_other`.

**Multi-select question blocks** (wide → long via melt; value `1` = selected):
- Protection concerns → columns prefixed `concern_`
- General information needs → columns prefixed `info_`
- Referral partners → columns prefixed `ref_partner_`
Columns ending in `_specify` are ignored as labels.

**`Column Mapping` sheet (optional)** — columns `cleaned_column_name` and
`original_column_name`; used to produce human-readable labels for the `concern_`,
`info_`, and `ref_partner_` blocks. Without it, labels are generated from the codes.

**PII removed before display/export:** `information_seeker_name`,
`residence_neighborhood_compound_house`, `information_seeker_phone`,
`alternative_phone`, `information_seeker_individual_number`,
`information_seeker_ration_or_wristband_number`.

**Row validity:** a record is kept only if all of these are present —
`interview_date`, `information_seeker_type`, `camp_location`,
`information_seeker_gender`, `age_group`, `request_category`, `action_taken`.

---

## How the analysis works
- **Record IDs** are assigned as `HD-#####` based on the source Excel row number.
- **Life stage** is derived from `information_seeker_age` (child age bands vs adult
  age bands; numeric fallback uses age < 18).
- **Gender reconciliation** maps Girl/Boy ↔ Woman/Man to match the derived life stage,
  flagging any change.
- **Adult disability (WGQ):** each domain answer is scored 1–4
  (1 = no difficulty … 4 = cannot do at all). A domain counts as a disability at score
  **3 or 4** — the official Washington Group cutoff. Domains are standardized into
  impairment types; per-person results yield status, type (single vs *Multiple
  Impairments*), domain count, multiplicity, max-score severity
  (No Disability / Disability / Severe Disability), and an exclusion-risk flag
  (any score ≥ 2, or an additional reported category).
- **Child disability** uses the `has_disability` flag plus the child impairment-type
  fields.
- **Reconciliation table** explains the difference between *specific impairment
  mentions* (one person can contribute several) and *unique adults with impairment*.
- **GPS classification** parses coordinates, range-checks them, and geofences against
  Turkana West (incl. Kakuma/Kalobeyei) and Dadaab (Hagadera/Ifo/Ifo 2/Dagahaley);
  anything outside is flagged **GPS outlier** for review.
- **Referral status** is mapped from the cleaned `action_taken` value.

## Filters & session behaviour
- Filters cascade top-to-bottom: **Date range → Camp → Helpdesk → Seeker type →
  Gender → Age group → Request category.** You must pick a **camp** to unlock the
  helpdesk and demographic filters.
- Each filter's options are recomputed from the data already narrowed by the filters
  above it, and stale selections are pruned automatically.
- **Refresh data** clears the Streamlit cache and reloads the workbook.
  **Clear filters** resets selections without reloading.
- Caching is keyed on the file's modification time, so editing the workbook and
  pressing Refresh picks up changes.

## Privacy & safeguarding
This is sensitive humanitarian protection data. The app already drops direct PII
before display and export. Before sharing more widely, also consider:
- The **Map** plots exact coordinates and the **Records** tab exports full rows —
  both can enable re-identification. Restrict access, or aggregate/jitter points.
- Apply **small-cell suppression** (e.g. hide counts < 5) on disaggregated cross-tabs
  where required by your protection-reporting policy.
- Host behind authentication and avoid committing the workbook to version control
  (see `.gitignore`).

---

## Configuration reference
Constants near the top of `app.py` you may want to adjust:
| Constant | Purpose |
|----------|---------|
| `DATA_FILE_PATH`, `LOGO_PATH` | Source locations. |
| `PII_COLUMNS` | Columns removed before display/export. |
| `AGE_GROUP_ORDER`, `CHILD_AGE_GROUPS`, `ADULT_AGE_GROUPS` | Age banding & sort order. |
| `GENDER_ORDER`, `GENDER_COLORS` | Gender categories and chart colours. |
| `WGQ_DISABILITY_DOMAINS` | Maps WGQ columns → impairment types. |
| `DISABILITY_TYPE_STANDARD_MAP` | Normalizes free-text disability labels. |
| `RECORD_PREVIEW_LIMIT` | Max rows shown in the Records preview (default 1000). |

The geofence bounding boxes live in `classify_gps_operational_area()` — update them if
your operational sites change.

---

## Troubleshooting
| Symptom | Likely cause / fix |
|---------|--------------------|
| `File not found: .../HELPDESK_DashboardData_Tdh_Kenya_D2.xlsx` | Put the workbook in `data/` with the exact name, or edit `DATA_FILE_PATH`. |
| Wrong sheet is read | The app uses a sheet named `cleaned_data` if present, otherwise the first sheet. Make your data the first sheet, or name it `cleaned_data`. |
| "No valid dashboard records were found" | Some required columns (e.g. `interview_date`, `request_type_protection_or_information`, `action_taken`) are missing or empty for every row. |
| KPI counts look zero / labels odd | Category strings must match the values the code checks (see [Known notes](#known-notes)). Provide the `Column Mapping` sheet for friendly labels. |
| Logo not showing | Add `assets/tdh-logo.png`; otherwise the title renders alone (expected). |
| Charts/tables empty after filtering | Widen the date range or clear filters; remember you must select a camp to unlock deeper filters. |
| `openpyxl` not found | `pip install -r requirements.txt` (openpyxl is the Excel engine pandas needs). |

## Deployment
- **Streamlit Community Cloud:** push `app.py` + `requirements.txt` to a repo and point
  the app at `app.py`. Add the data file as a private artifact or via upload — do **not**
  commit sensitive data publicly.
- **Docker (sketch):**
  ```dockerfile
  FROM python:3.12-slim
  WORKDIR /app
  COPY requirements.txt .
  RUN pip install --no-cache-dir -r requirements.txt
  COPY . .
  EXPOSE 8501
  CMD ["streamlit", "run", "app.py", "--server.address=0.0.0.0"]
  ```

## Known notes
- **Referral mapping spelling:** the code matches the action value
  `"Case referrred to Tdh national staff"` (note the triple **r** in "referrred").
  This spelling is **kept exactly as in the source** so it matches the cleaned Excel
  data. If your `action_taken` value is spelled differently, that referral category
  will read as "No referral" until the strings match — adjust either the data or the
  `records.loc[...]` mapping block in `load_data()`.
- Derivations run **once per data load** and are cached; large workbooks take longer on
  the first load and on each Refresh.

## License & attribution
Internal tool for Terre des hommes (Tdh) Kenya. Developed by **John Kul, MEAL Officer –
Tdh** (ImpactLens Africa). Add a formal license here if you intend to distribute.

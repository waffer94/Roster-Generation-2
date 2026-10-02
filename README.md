# F.A.S.T. Rescue — Course Roster & Sign-Off Sheet Builder

Two tools in one app: generate print-ready **Word course rosters** and **TLMS sign-off sheets** (Excel) from Excel exports. Upload, click, download.

---

## Quick Start

```bash
pip install python-docx openpyxl streamlit
```

**Web app:**
```bash
streamlit run app.py
```

**CLI — rosters:**
```bash
python build_rosters.py EXPORT.xlsx --templates templates --out rosters
```

**CLI — sign-off sheets:**
```bash
python build_signoff.py EXPORT.xlsx MAPPING.xlsx --out signoff_sheets
```

---

## 1. Course Roster Generation

Generates `.docx` rosters from an Excel export, one per course.

### Templates (7)

| Template | Triggered by | Layout |
|---|---|---|
| In-Class First Aid | `in-class` / `in class` | Portrait |
| Blended First Aid | `blended` | Portrait (Online Part 1 col) |
| Recertification First Aid | `recert` (alone) | Portrait |
| Working At Heights | `working at heights` | Portrait (WAHF/WAHR col) |
| Lift Truck | `lift truck` / `forklift` | Landscape |
| EWP | `ewp` / `elevated work platform` / `elevating work platform` | Landscape |
| General | anything else | Portrait |

### What It Does

- Detects course type and picks the correct template automatically
- Extracts parent customer name from the Name column (handles multi-colon sub-customer patterns)
- Anchors on the 8-digit `YYYYMMDD` date to extract the course name; always appends `Training`
- Strips account-number prefixes/suffixes from company names
- Strips company/contact name prefixes from location
- Formats contact as `Name – phone / Onsite Contact – phone` when onsite details exist
- Formats dates: single, two-day (`May 28 & 29, 2026`), multi-session (`May 2, 3 & 5, 2026`)
- Parses participants in any format (per-line, tab-separated, multi-line with emails, bulleted)
- Sorts names A–Z, adds a **Sr. No.** column (or fills the existing one on landscape templates)
- Sizes the table to `max(20, participants + 5)` rows
- Row heights stay constant unless a name wraps to a second line
- Skips cancelled courses by default

### Working at Heights — 3 Variants

| Course name contains | Reads from | WAHF/WAHR column |
|---|---|---|
| `working at heights` (no recert) | Participant Names only | All **WAHF** |
| `working at heights` + `recert`/`refresher` (no `full`) | Recertification Participants only | All **WAHR** |
| `working at heights` + `full` + `recert` | Both columns | **WAHF** / **WAHR** by source |

### First Aid + Recert Combos

| Course type | Template | Recert participants marked in |
|---|---|---|
| Blended + Recert | Blended | **Online Part 1** = `Recert` |
| In-Class + Recert | In-Class | **Instructor Notes** = `Recert` |

### Instructor Blanking

If the instructor name contains **Noeline** or **Frank Keegan**, the Instructor field on the roster is left blank. The name still appears in the filename.

### Same-Day Duplicate Handling

When the same company has multiple courses on the same date:

| What differs | Suffix on filename |
|---|---|
| Time slot | `Session 1`, `Session 2`, … |
| Course name | Course name appended |
| Location | Location appended |

### Filename Format

```
Instructor - First date - Company [- Suffix].docx
```

Title-cased with minor-word exceptions (`of`, `the`, `and`). Trailing dots stripped.

### CLI Options

| Flag | Description |
|---|---|
| `--templates PATH` | Folder with blank template `.docx` files (default: `.`) |
| `--out PATH` | Output folder (default: `rosters`) |
| `--keep-order` | Preserve export order instead of A–Z |
| `--include-cancelled` | Include cancelled courses |

### Roster Export Columns

| Column | Description |
|---|---|
| Name | Full entry: account IDs, colons, YYYYMMDD date, course name, city |
| Customer | Company name (may include account prefix/suffix) |
| Course Location | Address |
| Training Contact | Contact name |
| Training Contact Phone | Phone number(s) |
| On-Site Contact | Onsite contact name (optional) |
| On-Site Phone | Onsite phone (optional) |
| Course Date | Start date |
| Course End Date | End date |
| Other Date | Middle session date (optional) |
| Time (from) | Start time |
| Time (To) | End time |
| Instructors | Instructor name |
| Participant Names | Newline-separated participant list |
| Recertification Participants | Same format |

---

## 2. Sign-Off Sheet Generation

Generates `.xlsx` TLMS sign-off sheets from a sign-off export + a keyword mapping workbook.

### How It Works

1. Each input row's course name is matched against the mapping table (longest keyword match first).
2. Participants are parsed from tab-separated (`First\tLast\temail`), plain, or bulleted formats.
3. An `.xlsx` is created with two sheets per course.

### Output Format

**Sheet1** — user import:

| Column | Value |
|---|---|
| Login | Participant email |
| Firstname | First name |
| Lastname | Last name |
| Email | Same as Login |
| Company | `{Company} - {YYYY-MM-DD}` |
| Instructor (User) | Instructor name |
| Branch | `fastrescue` |
| Group | *(empty)* |
| Course | TLMS course name from mapping |

**Sheet2** — course completion:

| Column | Value |
|---|---|
| Usertocourses | Participant email |
| course | TLMS course name |
| EnrolledOndate | Course End Date + 1 day |
| CompletionDate | Same as EnrolledOndate |
| Status | `completed` |

### Mapping Table

The mapping workbook must have a sheet named `List - Sign off sheets` with these columns:

| Column | Description |
|---|---|
| Keywords | Text to match in the course name |
| All Participants | TLMS course name for everyone (when one course covers all) |
| Participant names | TLMS course for Participant Names column (for combo courses) |
| Recertification Participants | TLMS course for Recertification column (for combo courses) |

When "All Participants" has a value, everyone gets that single course. When "Participant names" and "Recertification Participants" have separate values, each group gets their respective course (e.g. TIFAB for full participants, TIFAR for recert).

### Filename Format

```
TLMS Sign Off - YYYYMMDD - Company - CODE(s).xlsx
```

Codes are derived from the parenthetical in the TLMS course name: `(TIFA-B-2-ON)` → `TIFAB`. Multiple codes joined with `&`.

### CLI Usage

```bash
python build_signoff.py EXPORT.xlsx MAPPING.xlsx --out signoff_sheets
```

### Sign-Off Export Columns

| Column | Description |
|---|---|
| Name | Full entry with account IDs, date, course name |
| Customer | Company name |
| Participant Names | Tab or newline-separated list (may include emails) |
| Recertification Participants | Same format |
| Course Date | Start date |
| Course End Date | End date |
| Instructors | Instructor name |

---

## Repository Structure

```
roster-app/
├── app.py                       # Streamlit web UI (both tools)
├── build_rosters.py             # Roster generation (CLI + library)
├── build_signoff.py             # Sign-off sheet generation (CLI + library)
├── requirements.txt
├── README.md
├── templates/                   # Blank Word templates for rosters
│   ├── Blended_First_Aid_Template.docx
│   ├── In_Class_First_Aid_Template.docx
│   ├── Recert_First_Aid_Template.docx
│   ├── Working_At_Heights_Template.docx
│   ├── Lift_Truck_Template.docx
│   ├── EWP_Template.docx
│   └── General_Template.docx
├── assets/
│   └── logo.png
└── .streamlit/
    └── config.toml
```

---

## Deploying (Free)

1. Push this repo to GitHub (public for the free tier).
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in, deploy with main file `app.py`.
3. Optional: add `password = "..."` in Settings → Secrets.

The mapping workbook for sign-off sheets is uploaded per-use, not bundled — update it anytime without redeploying.

### Updating

Push changes to GitHub. Streamlit Cloud redeploys automatically.

---

## Adding a New Roster Template

1. Add the `.docx` to `templates/` — include `Template` in the filename.
2. Add a keyword check in `detect_course_type()` in `build_rosters.py`.
3. Add a filename match in `classify_template()`.
4. Pre-existing Sr. No. columns and multi-row headers are handled automatically.

---

## Dependencies

| Package | Purpose |
|---|---|
| `python-docx` | Word `.docx` read/write (rosters) |
| `openpyxl` | Excel `.xlsx` read/write (both tools) |
| `streamlit` | Web interface (optional) |

```bash
pip install python-docx openpyxl streamlit
```

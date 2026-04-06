# ServiceNow RITM Tools

Two scripts for pulling RITM data out of ServiceNow and uploading signed documents back to it. Designed to be run in sequence: `get_bookings.py` first, then `upload_signed_files.py` after the documents have been signed.

---

## Prerequisites

- Python 3.11+
- Both scripts self-install their dependencies into a local `.venv` on first run — no manual `pip install` needed
- A ServiceNow account with read access to `sc_req_item` / `sc_task` and write access to attachments

---

## Setup

Set two environment variables before running either script:

| Variable | Description | Example |
|---|---|---|
| `SN_INSTANCE` | Your ServiceNow hostname (no `https://`) | `company.service-now.com` |
| `SN_USERNAME` | Your ServiceNow username / email | `john.smith@company.com` |

**PowerShell:**
```powershell
[Environment]::SetEnvironmentVariable("SN_INSTANCE", "company.service-now.com", "User")
[Environment]::SetEnvironmentVariable("SN_USERNAME", "john.smith@company.com", "User")
```

**bash / zsh:**
```bash
echo "export SN_INSTANCE='company.service-now.com'" >> ~/.bashrc
echo "export SN_USERNAME='john.smith@company.com'" >> ~/.bashrc
source ~/.bashrc
```

Your password is never stored — you will be prompted for it each time a script runs.

---

## Script 1 — `get_bookings.py`

Reads a `BookingsReportingData.tsv` export, looks up each REQ/RITM number in ServiceNow, downloads the first page of each RITM as a PDF, and generates a picking list.

### Input

Place a `BookingsReportingData.tsv` file (exported from Bookings) in the same folder as the script. The script scans every line of the file for `REQxxxxx` or `RITMxxxxx` patterns — the ticket numbers can appear anywhere in the rows.

If a REQ number is found, all child RITMs under it are fetched automatically.

### Output

All files are written to an `output/` folder created next to the script:

| File | Description |
|---|---|
| `output/ServiceNow_RITM_Report.txt` | Human-readable picking list with RITM number, description, catalog item and quantity |
| `output/<RITM number>.pdf` | First page only of the ServiceNow PDF export for each RITM |
| `output/Signed RITM/` | Empty folder created ready for signed documents (see Script 2) |

If the TSV covers more than one day, the script will warn you and ask for confirmation before continuing — this is a safeguard against accidentally processing a large date range export.

### Usage

```bash
python get_bookings.py
```

Optional flag to also send all generated files to the system default printer:

```bash
python get_bookings.py --print
```

---

## Script 2 — `upload_signed_files.py`

Uploads signed PDFs back to their RITM records in ServiceNow and moves any open SCTASKs to **Closed-Completle**, assigned to you.

### Input

Place signed PDFs into `output/Signed RITM/`. Files must be named exactly after their RITM number:

```
output/
└── Signed RITM/
    ├── RITM0012345.pdf
    └── RITM0012346.pdf
```

The script only processes files matching the pattern `RITMxxxxxxx.pdf` — anything else in the folder is ignored.

### What it does for each RITM

1. Looks up the RITM record in ServiceNow
2. Prints the record details (description, catalog item, CI, quantity)
3. Lists any open/work in progress SCTASKs linked to the RITM
4. Uploads the PDF as an attachment to the `sc_req_item` record
5. Transitions any **Open** or **Work in Progress** SCTASKs to **Closed-Complete** and assigns them to your username

### Usage

```bash
python upload_signed_files.py
```

---

## Typical workflow

```
1. Export BookingsReportingData.tsv from Bookings
2. python get_bookings.py          ← downloads PDFs + creates picking list
3. Print / review the PDFs
4. Get them signed
5. Place signed PDFs into output/Signed RITM/
6. python upload_signed_files.py   ← uploads to ServiceNow + updates tasks
```

---

## File structure

```
.
├── get_bookings.py
├── upload_signed_files.py
├── BookingsReportingData.tsv       ← you provide this
├── .venv/                          ← auto-created on first run
└── output/
    ├── ServiceNow_RITM_Report.txt
    ├── RITM0012345.pdf
    ├── RITM0012346.pdf
    └── Signed RITM/
        ├── RITM0012345.pdf         ← signed copies go here
        └── RITM0012346.pdf
```

## Acknowledgements

Developed with assistance from Qwen 2.5 Coder 32B (Q4) https://ollama.com/library/qwen2.5-coder:32b-instruct-q4_K_M for code generation and debugging

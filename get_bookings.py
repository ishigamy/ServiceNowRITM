#!/usr/bin/env python3
# VENV BOOTSTRAPPER
import os
import sys
import subprocess
import venv as _venv


def _setup_venv() -> None:
    """Create / activate a project-local venv on first run."""
    if sys.prefix != sys.base_prefix:
        return

    script_dir = os.path.dirname(os.path.abspath(__file__))
    venv_dir = os.path.join(script_dir, ".venv")

    if not os.path.exists(venv_dir):
        print(f"First-time setup: creating virtualenv in {venv_dir} …")
        _venv.create(venv_dir, with_pip=True)

    python = (
        os.path.join(venv_dir, "Scripts", "python.exe")
        if os.name == "nt"
        else os.path.join(venv_dir, "bin", "python")
    )

    print("Checking dependencies …")
    for pkg in ("pypdf", "requests"):
        result = subprocess.run(
            [python, "-m", "pip", "show", "-q", pkg],
            capture_output=True,
        )
        if result.returncode != 0:
            subprocess.check_call([python, "-m", "pip", "install", "-q", pkg])

    print("Re-launching inside venv …\n")
    subprocess.call([python, os.path.abspath(__file__)] + sys.argv[1:])
    sys.exit(0)


_setup_venv()

# STANDARD IMPORTS
import argparse
import getpass
import io
import re
from datetime import datetime
from typing import Optional

import requests
from pypdf import PdfReader, PdfWriter
from requests.auth import HTTPBasicAuth

# How long (seconds) to wait for ServiceNow to respond before giving up
_HTTP_TIMEOUT = 30


# HELPERS
def _check_printer() -> bool:
    """
    Return True if a real printer appears to be available.
    On Windows, warns if the default printer is 'Microsoft Print to PDF'.
    """
    if sys.platform == "win32":
        try:
            out = subprocess.run(
                ["powershell", "-Command",
                 "(Get-CimInstance -ClassName Win32_Printer | Where-Object {$_.Default}).Name"],
                capture_output=True, text=True, timeout=10,
            )
            name = out.stdout.strip()
            if "print to pdf" in name.lower():
                print(
                    "[warn] Your default printer is 'Microsoft Print to PDF'.\n"
                    "       Install or set a physical printer as default before using -p / -pa.\n"
                )
                return False
        except Exception:
            pass  # Can't determine — let it try
    return True


def _print_file(filepath: str) -> None:
    """Send *filepath* to the OS default printer (best-effort)."""
    try:
        if sys.platform == "win32":
            os.startfile(filepath, "print")  # type: ignore[attr-defined]
        else:
            subprocess.run(["lp", filepath], check=False)
    except Exception as exc:  # noqa: BLE001
        print(f"[warn] Could not print {filepath}: {exc}")


def _str_val(field: object) -> str:
    """Extract a plain string from a ServiceNow display-value dict or raw value."""
    if isinstance(field, dict):
        return str(field.get("display_value") or field.get("value") or "")
    return str(field) if field is not None else ""


# SERVICE-NOW CLIENT
class ServiceNowClient:
    """
    Thin wrapper around the ServiceNow Table API.
    """

    # Fields fetched for every RITM record
    _RITM_FIELDS = "sys_id,number,short_description,cat_item,cmdb_ci,quantity"

    def __init__(self, instance: str, username: str, password: str) -> None:
        if not instance or "/" in instance or instance.startswith("http"):
            raise ValueError(
                "SN_INSTANCE must be a bare hostname, e.g. company.service-now.com"
            )
        self._base = f"https://{instance}"
        self._session = requests.Session()
        self._session.auth = HTTPBasicAuth(username, password)
        self._session.headers.update({"Accept": "application/json"})
        self._session.verify = True

    # Public helpers
    def close(self) -> None:
        """Release the underlying TCP connection pool."""
        self._session.close()

    def check_auth(self) -> bool:
        """
        Quick connectivity + auth test before the main loop.
        Returns True on success, prints the reason and returns False on failure.
        """
        url = f"{self._base}/api/now/table/sc_req_item"
        try:
            r = self._session.get(
                url,
                params={"sysparm_limit": "1"},
                timeout=_HTTP_TIMEOUT,
            )
            if r.status_code == 401:
                print(
                    "[error] Authentication failed (401 Unauthorized).\n"
                    "        Check your username and password.\n"
                )
                return False
            r.raise_for_status()
            return True
        except requests.RequestException as exc:
            print(f"[error] Could not reach ServiceNow: {exc}")
            return False

    def get_ritms_from_req(self, req_number: str, debug: bool = False) -> list:
        """Return all RITMs belonging to a REQ number."""
        return self._query_ritms(f"request.number={req_number.upper()}", debug=debug)

    def get_ritm_by_number(self, ritm_number: str, debug: bool = False) -> Optional[dict]:
        """Return a single RITM record, or None if not found."""
        results = self._query_ritms(f"number={ritm_number.upper()}", limit=1, debug=debug)
        return results[0] if results else None

    def download_pdf_first_page(self, sys_id: str) -> Optional[bytes]:
        """
        Download the raw PDF export of a sc_req_item record.
        Returns raw bytes or None on error.
        """
        url = f"{self._base}/sc_req_item.do"
        try:
            r = self._session.get(
                url,
                params={"PDF": "", "sys_id": sys_id},
                timeout=_HTTP_TIMEOUT,
            )
            r.raise_for_status()
            content = r.content
            if not content.startswith(b"%PDF"):
                return None
            return content
        except requests.RequestException as exc:
            print(f"    [warn] PDF download failed for sys_id {sys_id}: {exc}")
            return None

    # Private helpers
    def _query_ritms(self, query: str, limit: Optional[int] = None, debug: bool = False) -> list:
        url = f"{self._base}/api/now/table/sc_req_item"
        params: dict = {
            "sysparm_query": query,
            "sysparm_fields": self._RITM_FIELDS,
            "sysparm_display_value": "true",
        }
        if limit is not None:
            params["sysparm_limit"] = str(limit)

        try:
            r = self._session.get(url, params=params, timeout=_HTTP_TIMEOUT)
            if debug:
                print(f"    [debug] GET {r.url}")
                print(f"    [debug] Status: {r.status_code}")
                print(f"    [debug] Body (first 500 chars): {r.text[:500]}")
            r.raise_for_status()
            return r.json().get("result", [])
        except requests.HTTPError as exc:
            print(f"    [warn] HTTP error ({query}): {exc}")
            return []
        except requests.RequestException as exc:
            print(f"    [warn] Request failed ({query}): {exc}")
            return []

    @staticmethod
    def normalise_ritm(record: dict) -> dict:
        """Ensure 'number' and 'sys_id' are always plain strings in-place."""
        for key in ("number", "sys_id"):
            val = record.get(key)
            if isinstance(val, dict):
                record[key] = val.get("display_value" if key == "number" else "value", "")
            else:
                record[key] = val if val is not None else ""
            record[key] = str(record[key])
        return record


# TEXT REPORT GENERATOR
class TextReportGenerator:
    """Writes a human-readable picking-list .txt file, sorted by catalog item."""

    def __init__(self, output_dir: str, filename: str = "ServiceNow_RITM_Report.txt") -> None:
        self.filepath = os.path.join(output_dir, filename)

    @staticmethod
    def _trunc(s: str, maxlen: int) -> str:
        return (s[: maxlen - 1] + "…") if len(s) > maxlen else s

    def generate(self, ritm_list: list, send_to_printer: bool = False) -> None:
        """
        Write the picking list.  Columns are sized to fit within ~93 chars
        so the table prints cleanly on A4 portrait at 9pt monospace.

        Layout (~93 chars):
          RITM (14) | Description (22) | Catalog Item (22) | Config Item (18) | Qty (3)
        Config Item (serial number) is never truncated.
        """
        # Sort by catalog item name (case-insensitive)
        ritm_list = sorted(
            ritm_list, key=lambda r: _str_val(r.get("cat_item", "")).lower()
        )

        w_ritm, w_desc, w_item, w_ci = 14, 22, 22, 18
        header = (
            f"{'RITM':<{w_ritm}} | {'Description':<{w_desc}} "
            f"| {'Catalog Item':<{w_item}} | {'Config Item':<{w_ci}} | Qty"
        )
        sep = "-" * len(header)

        print(f"\nGenerating text report: {self.filepath}")
        try:
            with open(self.filepath, "w", encoding="utf-8") as fh:
                fh.write(f"--- Picking List ({len(ritm_list)} items) ---\n\n")
                fh.write(header + "\n")
                fh.write(sep + "\n")
                for rec in ritm_list:
                    num  = _str_val(rec.get("number"))
                    desc = self._trunc(_str_val(rec.get("short_description")), w_desc)
                    item = self._trunc(_str_val(rec.get("cat_item")), w_item)
                    ci   = _str_val(rec.get("cmdb_ci"))  # never truncated
                    qty  = _str_val(rec.get("quantity"))
                    fh.write(
                        f"{num:<{w_ritm}} | {desc:<{w_desc}} "
                        f"| {item:<{w_item}} | {ci:<{w_ci}} | {qty}\n"
                    )
        except OSError as exc:
            print(f"[error] Could not write report: {exc}")
            return

        if send_to_printer:
            _print_file(self.filepath)


# PDF PROCESSOR
class PDFProcessor:
    """Saves only the first page of a RITM PDF to disk."""

    def __init__(self, output_dir: str) -> None:
        self.output_dir = output_dir

    def save_first_page(
        self,
        ticket_number: str,
        raw_pdf_bytes: Optional[bytes],
        send_to_printer: bool = False,
    ) -> None:
        if not raw_pdf_bytes:
            return
        filepath = os.path.join(self.output_dir, f"{ticket_number}.pdf")
        try:
            reader = PdfReader(io.BytesIO(raw_pdf_bytes))
            if not reader.pages:
                return
            writer = PdfWriter()
            writer.add_page(reader.pages[0])
            with open(filepath, "wb") as fh:
                writer.write(fh)
        except Exception as exc:  # noqa: BLE001
            print(f"    [warn] PDF processing error for {ticket_number}: {exc}")
            return

        if send_to_printer:
            _print_file(filepath)


# TSV PARSER
_TICKET_RE = re.compile(r"\b(REQ\d+|RITM\d+)\b", re.IGNORECASE)


def get_tickets_from_tsv(tsv_path: str) -> list[str]:
    """
    Parse BookingsReportingData.tsv and return sorted, de-duplicated
    REQxxxxx / RITMxxxxx ticket numbers (upper-cased).
    """
    found: set[str] = set()
    try:
        with open(tsv_path, encoding="utf-8") as fh:
            for line in fh:
                for match in _TICKET_RE.findall(line):
                    found.add(match.upper())
    except OSError as exc:
        print(f"[warn] Could not read {tsv_path}: {exc}")
    return sorted(found)


def get_date_range_from_tsv(tsv_path: str) -> tuple[Optional[str], Optional[str]]:
    dates: set[str] = set()
    try:
        with open(tsv_path, encoding="utf-8") as fh:
            next(fh, None)  # skip header
            for line in fh:
                col = line.split("\t")[0].strip()
                for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S"):
                    try:
                        dates.add(datetime.strptime(col, fmt).strftime("%d/%m/%Y"))
                        break
                    except ValueError:
                        continue
    except OSError:
        pass
    if not dates:
        return None, None
    sorted_dates = sorted(dates, key=lambda d: d.split("/")[::-1])
    return sorted_dates[0], sorted_dates[-1]


# PROGRESS BAR (pacman-style)


def _progress_bar(current: int, total: int, width: int = 30, suffix: str = "") -> None:
    """Draw a pacman-style progress bar: C/c moves through dashes, eating o dots."""
    # Build the static track: dashes with 'o' every 4th position
    track = []
    for i in range(width):
        track.append("o" if (i + 1) % 4 == 0 else "-")

    # Pacman position (0 … width), stays put if % hasn't changed enough
    fraction = current / total if total else 1
    pos = min(int(width * fraction), width)
    pac = "C" if current % 2 == 0 else "c"

    # Compose the bar
    bar = ""
    for i in range(width):
        if i < pos:
            bar += "-"       # eaten trail
        elif i == pos:
            bar += pac       # pacman
        elif track[i] == "o":
            bar += "o"       # upcoming dot
        else:
            bar += " "       # space between dots (not yet eaten)
    if pos >= width:
        bar = "-" * width    # done

    pct = int(fraction * 100)
    line = f"\r  ({current}/{total}) [{bar}] {pct}%"
    if suffix:
        line += f"  {suffix}"
    sys.stdout.write(f"{line:<100}")
    sys.stdout.flush()
    if current == total:
        sys.stdout.write("\n")


# MAIN
def main() -> None:
    parser = argparse.ArgumentParser(
        description="Download RITM PDFs and generate a picking list from BookingsReportingData.tsv"
    )
    print_group = parser.add_mutually_exclusive_group()
    print_group.add_argument(
        "-p", action="store_true", dest="print_txt",
        help="Print only the text picking list",
    )
    print_group.add_argument(
        "-pa", action="store_true", dest="print_all",
        help="Print the text picking list AND all generated PDFs",
    )
    parser.add_argument(
        "-d", "--debug", action="store_true",
        help="Print debug info for the first ServiceNow query",
    )
    args = parser.parse_args()

    print_txt = args.print_txt or args.print_all
    print_pdfs = args.print_all

    # Early printer check when any print flag is active
    if print_txt and not _check_printer():
        sys.exit(1)

    # --- Credentials ---
    instance = os.environ.get("SN_INSTANCE", "").strip()
    username = os.environ.get("SN_USERNAME", "").strip()

    if not instance or not username:
        print("[error] Set SN_INSTANCE and SN_USERNAME environment variables first.")
        print('        Example (PowerShell): [Environment]::SetEnvironmentVariable("SN_INSTANCE", "company.service-now.com", "User")')
        print('        Example (bash):       echo "export SN_INSTANCE=\'company.service-now.com\'" >> ~/.bashrc')
        print("        Restart terminal after setting environment variables")
        sys.exit(1)

    password = getpass.getpass(f"ServiceNow password for {username}: ")
    if not password:
        print("[error] Password cannot be empty.")
        sys.exit(1)

    # --- Locate input file ---
    script_dir = os.path.dirname(os.path.abspath(__file__))
    tsv_path = os.path.join(script_dir, "BookingsReportingData.tsv")

    if not os.path.exists(tsv_path):
        print("[error] BookingsReportingData.tsv not found — export from Bookings and place it here.")
        sys.exit(1)

    # --- Check accidental large export range ---
    earliest, latest = get_date_range_from_tsv(tsv_path)
    if earliest and latest and earliest != latest:
        answer = input(
            f"Date range more than 1 day ({earliest} - {latest}). Continue? Y/n: "
        ).strip().lower()
        if answer not in ("", "y"):
            sys.exit(0)

    raw_tickets = get_tickets_from_tsv(tsv_path)
    if not raw_tickets:
        print("[info] No REQ or RITM numbers found in BookingsReportingData.tsv.")
        sys.exit(0)

    print(f"Found {len(raw_tickets)} ticket reference(s) in TSV.")

    # --- Output dir ---
    output_dir = os.path.join(script_dir, "output")
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(os.path.join(output_dir, "Signed RITM"), exist_ok=True)

    # --- Connect + auth check ---
    client = ServiceNowClient(instance, username, password)
    # Wipe the password string from local scope as soon as it is handed off.
    del password

    if not client.check_auth():
        client.close()
        sys.exit(1)

    # --- Fetch RITMs ---
    print("\nFetching RITM data from ServiceNow …")
    final_ritms: dict[str, dict] = {}
    first_query = True
    total = len(raw_tickets)

    for i, ticket in enumerate(raw_tickets, 1):
        _progress_bar(i, total)
        debug = args.debug and first_query
        if ticket.startswith("REQ"):
            children = client.get_ritms_from_req(ticket, debug=debug)
            if debug:
                first_query = False
            for child in children:
                ServiceNowClient.normalise_ritm(child)
                num = child.get("number", "")
                if num:
                    final_ritms[num] = child
        elif ticket.startswith("RITM"):
            result = client.get_ritm_by_number(ticket, debug=debug)
            if debug:
                first_query = False
            if result:
                ServiceNowClient.normalise_ritm(result)
                final_ritms[ticket] = result

    if not final_ritms:
        print("[info] No RITM data retrieved.")
        client.close()
        sys.exit(0)

    ritm_list = list(final_ritms.values())
    print(f"Retrieved {len(ritm_list)} RITM(s).")

    # --- Text report (sorted by catalog item inside generate()) ---
    TextReportGenerator(output_dir).generate(ritm_list, send_to_printer=print_txt)

    # --- PDFs ---
    print(f"\nDownloading and slicing {len(ritm_list)} PDF(s) …")
    pdf_proc = PDFProcessor(output_dir)
    pdf_total = len(ritm_list)

    for i, rec in enumerate(ritm_list, 1):
        num    = rec["number"]
        sys_id = rec["sys_id"]
        if not sys_id:
            _progress_bar(i, pdf_total, suffix=f"{num} — skipped")
            continue
        raw = client.download_pdf_first_page(sys_id)
        pdf_proc.save_first_page(num, raw, send_to_printer=print_pdfs)
        _progress_bar(i, pdf_total, suffix=f"{num}.pdf")

    client.close()
    print(f"\nDone. Files are in: {output_dir}")


if __name__ == "__main__":
    main()

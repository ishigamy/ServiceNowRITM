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
    for pkg in ("requests",):
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
import getpass
import re
from typing import Optional

import requests
from requests.auth import HTTPBasicAuth

_HTTP_TIMEOUT = 30

# ServiceNow state values for sc_task
_STATE_OPEN             = "1"
_STATE_CLOSED_COMPLETE  = "3"


# SERVICE-NOW CLIENT
class ServiceNowClient:
    _RITM_FIELDS = "sys_id,number,short_description,cat_item,cmdb_ci,quantity"
    # States 3=Closed Complete, 4=Closed Incomplete, 6=Closed Skipped, 7=Cancelled
    _SCTASK_OPEN_FILTER = "state!=3^state!=4^state!=6^state!=7"
    _SCTASK_FIELDS = "sys_id,number,short_description,state,assigned_to"

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

    def close(self) -> None:
        self._session.close()

    def check_auth(self) -> bool:
        url = f"{self._base}/api/now/table/sc_req_item"
        try:
            r = self._session.get(url, params={"sysparm_limit": "1"}, timeout=_HTTP_TIMEOUT)
            if r.status_code == 401:
                print("[error] Authentication failed (401 Unauthorized).\n        Check your username and password.\n")
                return False
            r.raise_for_status()
            return True
        except requests.RequestException as exc:
            print(f"[error] Could not reach ServiceNow: {exc}")
            return False

    def get_ritm_by_number(self, ritm_number: str) -> Optional[dict]:
        results = self._query_ritms(f"number={ritm_number.upper()}", limit=1)
        return results[0] if results else None

    def get_open_sctasks_for_ritm(self, ritm_sys_id: str) -> list:
        """Return all open SCTASKs linked to a RITM sys_id."""
        url = f"{self._base}/api/now/table/sc_task"
        params = {
            "sysparm_query": f"request_item={ritm_sys_id}^{self._SCTASK_OPEN_FILTER}",
            "sysparm_fields": self._SCTASK_FIELDS,
            "sysparm_display_value": "true",
        }
        try:
            r = self._session.get(url, params=params, timeout=_HTTP_TIMEOUT)
            r.raise_for_status()
            return r.json().get("result", [])
        except requests.HTTPError as exc:
            print(f"    [warn] HTTP error fetching SCTASKs for {ritm_sys_id}: {exc}")
            return []
        except requests.RequestException as exc:
            print(f"    [warn] Request failed fetching SCTASKs for {ritm_sys_id}: {exc}")
            return []

    def upload_attachment(self, table_sys_id: str, filename: str, pdf_bytes: bytes) -> bool:
        """Upload a PDF as an attachment to a sc_req_item record."""
        url = f"{self._base}/api/now/attachment/file"
        params = {
            "table_name": "sc_req_item",
            "table_sys_id": table_sys_id,
            "file_name": filename,
        }
        headers = {"Content-Type": "application/pdf"}
        try:
            r = self._session.post(
                url,
                params=params,
                headers=headers,
                data=pdf_bytes,
                timeout=_HTTP_TIMEOUT,
            )
            r.raise_for_status()
            return True
        except requests.HTTPError as exc:
            print(f"    [warn] Attachment upload failed (HTTP {exc.response.status_code}): {exc}")
            return False
        except requests.RequestException as exc:
            print(f"    [warn] Attachment upload failed: {exc}")
            return False

    def update_sctask(self, sctask_sys_id: str, state: str, assigned_to_email: str) -> bool:
        """
        PATCH a single sc_task: update state and assign to a user by email.
        Uses sysparm_input_display_value=true so ServiceNow resolves the email
        to a sys_user record automatically.
        """
        url = f"{self._base}/api/now/table/sc_task/{sctask_sys_id}"
        params = {"sysparm_input_display_value": "true"}
        payload: dict = {"state": state}
        if assigned_to_email:
            payload["assigned_to"] = assigned_to_email
        try:
            r = self._session.patch(
                url,
                params=params,
                json=payload,
                timeout=_HTTP_TIMEOUT,
            )
            r.raise_for_status()
            return True
        except requests.HTTPError as exc:
            print(f"    [warn] SCTASK update failed (HTTP {exc.response.status_code}): {exc}")
            return False
        except requests.RequestException as exc:
            print(f"    [warn] SCTASK update failed: {exc}")
            return False

    def _query_ritms(self, query: str, limit: Optional[int] = None) -> list:
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
        for key in ("number", "sys_id"):
            val = record.get(key)
            if isinstance(val, dict):
                record[key] = val.get("display_value" if key == "number" else "value", "")
            else:
                record[key] = val if val is not None else ""
            record[key] = str(record[key])
        return record

    @staticmethod
    def _str_val(field: object) -> str:
        if isinstance(field, dict):
            return str(field.get("display_value") or field.get("value") or "")
        return str(field) if field is not None else ""


# OUTPUT FOLDER SCANNER
_RITM_FILENAME_RE = re.compile(r"^(RITM\d+)\.pdf$", re.IGNORECASE)


def get_ritm_numbers_from_output(output_dir: str) -> list[str]:
    """Return sorted RITM numbers extracted from PDF filenames in output_dir."""
    found = []
    try:
        for fname in sorted(os.listdir(output_dir)):
            match = _RITM_FILENAME_RE.match(fname)
            if match:
                found.append(match.group(1).upper())
    except OSError as exc:
        print(f"[warn] Could not read output directory: {exc}")
    return found


def process_ritm(client: ServiceNowClient, ritm_number: str, output_dir: str, assigned_to_email: str) -> None:
    """Look up a single RITM, upload its PDF, and move Open SCTASKs to Work in Progress."""
    print(f"\n{'─' * 60}")
    print(f"  {ritm_number}")
    print(f"{'─' * 60}")

    # --- Fetch RITM ---
    record = client.get_ritm_by_number(ritm_number)
    if not record:
        print(f"  [skip] No record found in ServiceNow.")
        return

    ServiceNowClient.normalise_ritm(record)
    sys_id = record.get("sys_id", "")

    print(f"  Short description: {record.get('short_description')}")
    print(f"  Catalog item:      {record.get('cat_item')}")
    print(f"  CMDB CI:           {record.get('cmdb_ci')}")
    print(f"  Quantity:          {record.get('quantity')}")

    if not sys_id:
        print(f"  [skip] No sys_id on record — cannot continue.")
        return

    # --- Fetch open SCTASKs ---
    tasks = client.get_open_sctasks_for_ritm(sys_id)
    if tasks:
        print(f"\n  Open SCTASKs ({len(tasks)}):")
        for task in tasks:
            num   = ServiceNowClient._str_val(task.get("number"))
            desc  = ServiceNowClient._str_val(task.get("short_description"))
            state = ServiceNowClient._str_val(task.get("state"))
            who   = ServiceNowClient._str_val(task.get("assigned_to"))
            print(f"    {num}  [{state}]  {desc}  →  {who or '(unassigned)'}")
    else:
        print(f"\n  No open SCTASKs found.")

    # --- Upload PDF attachment ---
    pdf_filename = f"{ritm_number}.pdf"
    pdf_path = os.path.join(output_dir, pdf_filename)

    if not os.path.isfile(pdf_path):
        print(f"\n  [skip] PDF not found locally: {pdf_filename}")
        return

    print(f"\n  Uploading {pdf_filename} …", end=" ", flush=True)
    with open(pdf_path, "rb") as fh:
        pdf_bytes = fh.read()
    upload_ok = client.upload_attachment(sys_id, pdf_filename, pdf_bytes)
    print("OK" if upload_ok else "FAILED")

    if not upload_ok:
        return

    # --- Transition Open SCTASKs → Work in Progress ---
    if not tasks:
        return

    print(f"\n  Updating Open SCTASKs to Closed Complete …")
    for task in tasks:
        task_sys_id = ServiceNowClient._str_val(task.get("sys_id"))
        task_num    = ServiceNowClient._str_val(task.get("number"))
        task_state  = ServiceNowClient._str_val(task.get("state"))

        if task_state.lower() not in ("open", "work in progress"):
            print(f"    {task_num}  [{task_state}] — skipped (not Open or Work in Progress)")
            continue

        print(f"    {task_num}  {task_state} → Closed Complete, assigned to {assigned_to_email} …", end=" ", flush=True)
        ok = client.update_sctask(task_sys_id, _STATE_CLOSED_COMPLETE, assigned_to_email)
        print("OK" if ok else "FAILED")


# MAIN
def main() -> None:
    # --- Credentials ---
    instance = os.environ.get("SN_INSTANCE", "").strip()
    username = os.environ.get("SN_USERNAME", "").strip()

    if not instance or not username:
        print("[error] Set SN_INSTANCE and SN_USERNAME environment variables first.")
        sys.exit(1)

    user_email = username

    password = getpass.getpass(f"ServiceNow password for {username}: ")
    if not password:
        print("[error] Password cannot be empty.")
        sys.exit(1)

    # --- Locate output folder ---
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, "output", "Signed RITM")

    if not os.path.isdir(output_dir):
        print(f"[error] Output folder not found: {output_dir}")
        sys.exit(1)

    # --- Scan for RITM PDFs ---
    ritm_numbers = get_ritm_numbers_from_output(output_dir)

    if not ritm_numbers:
        print("[info] No RITM PDF files found in output folder.")
        sys.exit(0)

    print(f"Found {len(ritm_numbers)} RITM PDF(s) in output folder.")
    print(f"Tasks will be assigned to: {user_email}")

    # --- Connect + auth check ---
    client = ServiceNowClient(instance, username, password)
    del password

    if not client.check_auth():
        client.close()
        sys.exit(1)

    # --- Process each RITM ---
    for ritm_number in ritm_numbers:
        process_ritm(client, ritm_number, output_dir, user_email)

    client.close()
    print(f"\n{'─' * 60}")
    print(f"Done. Processed {len(ritm_numbers)} RITM(s).")


if __name__ == "__main__":
    main()
"""
Direct extraction from QuickBooks Desktop via the QBXMLRP2 COM SDK.

Platform: Windows-only (requires the QuickBooks SDK and pywin32).
On non-Windows systems, every method raises QBDUnavailable — the UI
should fall back to manual Excel upload.

QBXMLRP2 reference:
  https://developer.intuit.com/app/developer/qbdesktop/docs/develop/explore-the-quickbooks-sdk

Typical flow:
    ex = QBDExtractor(company_file=None)  # uses currently-open file
    ex.open()
    ex.extract_all(output_dir=Path("qbd_exports"))
    ex.close()
"""

from __future__ import annotations

import logging
import platform
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import pandas as pd

log = logging.getLogger("qbd.extract")

IS_WINDOWS = platform.system() == "Windows"


class QBDUnavailable(RuntimeError):
    """Raised when QBD direct-connect isn't available on this platform."""


# ─── XML helpers ──────────────────────────────────────────────────

def _qbxml_envelope(inner: str, version: str = "16.0") -> str:
    """Wrap a request in the standard QBXML header."""
    return (
        f'<?xml version="1.0"?>'
        f'<?qbxml version="{version}"?>'
        f'<QBXML><QBXMLMsgsRq onError="continueOnError">'
        f'{inner}'
        f'</QBXMLMsgsRq></QBXML>'
    )


def _text(elem: ET.Element | None, tag: str, default: str = "") -> str:
    if elem is None:
        return default
    child = elem.find(tag)
    return child.text if child is not None and child.text is not None else default


def _bool(elem: ET.Element | None, tag: str) -> bool:
    val = _text(elem, tag, "").strip().lower()
    return val in ("true", "1", "yes")


# ─── Extractor ────────────────────────────────────────────────────

@dataclass
class QBDConnectionInfo:
    app_name: str = "QBD-QBO Migration"
    app_id: str = ""
    company_file: str = ""           # empty = use currently-open file
    connection_mode: int = 1          # 1 = localQBD, 2 = localQBDLaunchUI


class QBDExtractor:
    """Wraps a QBXMLRP2 COM session and extracts QBD data to DataFrames/Excel."""

    def __init__(self, info: QBDConnectionInfo | None = None):
        self.info = info or QBDConnectionInfo()
        self._rp = None           # QBXMLRP2.RequestProcessor
        self._ticket: str | None = None
        self._co_initialized = False
        if not IS_WINDOWS:
            log.info("QBD direct extractor instantiated on non-Windows — open() will raise.")

    # ── Connection lifecycle ──────────────────────────────────

    def open(self) -> dict:
        """Open a QBXMLRP2 session. Returns a small metadata dict on success.

        IMPORTANT: QBXMLRP2 is an STA COM object. It must be opened, used, and
        closed on the *same* thread. Call this inside whichever thread will do
        the extraction — do not open on one thread and call extract_* from
        another, or the process will crash (RPC_E_WRONG_THREAD).
        """
        if not IS_WINDOWS:
            raise QBDUnavailable(
                "QuickBooks Desktop direct connection requires Windows and the QB SDK. "
                "On macOS/Linux, export Excel files from QBD and use the Upload option."
            )
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
        except ImportError as e:
            raise QBDUnavailable(
                f"pywin32 is not installed: {e}. Run: pip install pywin32"
            )

        # Initialize COM for this thread. If it's already initialized (e.g.
        # the main thread), CoInitialize returns S_FALSE — still safe.
        try:
            pythoncom.CoInitialize()
            self._co_initialized = True
        except Exception as e:
            log.warning("CoInitialize returned: %s", e)
            self._co_initialized = False

        try:
            self._rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
            self._rp.OpenConnection2(self.info.app_id, self.info.app_name, self.info.connection_mode)
            self._ticket = self._rp.BeginSession(self.info.company_file, 0)  # 0 = dontCare
            log.info("QBD session opened (ticket=%s)", self._ticket)
            return {
                "connected": True,
                "company_file": self.info.company_file or "(currently open)",
                "ticket": self._ticket,
            }
        except Exception as e:
            self._safe_close()
            raise QBDUnavailable(f"Failed to open QBD session: {e}") from e

    def close(self) -> None:
        self._safe_close()

    def _safe_close(self) -> None:
        try:
            if self._rp and self._ticket:
                self._rp.EndSession(self._ticket)
        except Exception as e:
            log.warning("EndSession failed: %s", e)
        try:
            if self._rp:
                self._rp.CloseConnection()
        except Exception as e:
            log.warning("CloseConnection failed: %s", e)
        self._rp = None
        self._ticket = None
        # Balance CoInitialize with CoUninitialize on the same thread
        if getattr(self, "_co_initialized", False):
            try:
                import pythoncom  # type: ignore
                pythoncom.CoUninitialize()
            except Exception as e:
                log.warning("CoUninitialize failed: %s", e)
            self._co_initialized = False

    # ── Request dispatcher ────────────────────────────────────

    def _send(self, qbxml: str) -> ET.Element:
        """Send a QBXML request and return the parsed XML root."""
        if not (self._rp and self._ticket):
            raise QBDUnavailable("QBD session is not open — call open() first")
        response_xml = self._rp.ProcessRequest(self._ticket, qbxml)
        return ET.fromstring(response_xml)

    # ── Entity extractors ─────────────────────────────────────

    def extract_accounts(self) -> pd.DataFrame:
        log.info("Extracting Chart of Accounts")
        root = self._send(_qbxml_envelope(
            '<AccountQueryRq requestID="1"><ActiveStatus>All</ActiveStatus></AccountQueryRq>'
        ))
        rows: list[dict[str, Any]] = []
        for a in root.iter("AccountRet"):
            rows.append({
                "Account": _text(a, "FullName") or _text(a, "Name"),
                "Type": _text(a, "AccountType"),
                "Number": _text(a, "AccountNumber"),
                "Description": _text(a, "Desc"),
                "Balance": _text(a, "Balance"),
                "Active": _bool(a, "IsActive"),
            })
        return pd.DataFrame(rows)

    def extract_customers(self) -> pd.DataFrame:
        """Returns a single DataFrame; jobs (Parent:Child) come through naturally
        in the FullName field and are handled by the downstream transformer."""
        log.info("Extracting Customers + Jobs")
        root = self._send(_qbxml_envelope(
            '<CustomerQueryRq requestID="1"><ActiveStatus>All</ActiveStatus></CustomerQueryRq>'
        ))
        rows: list[dict[str, Any]] = []
        for c in root.iter("CustomerRet"):
            bill = c.find("BillAddress")
            rows.append({
                "Name": _text(c, "FullName"),
                "Company Name": _text(c, "CompanyName"),
                "Main Email": _text(c, "Email"),
                "Main Phone": _text(c, "Phone"),
                "Terms": _text(c.find("TermsRef"), "FullName"),
                "Job Status": _text(c, "JobStatus"),
                "Bill To 1": _text(bill, "Addr1"),
                "Bill To City": _text(bill, "City"),
                "Bill To State": _text(bill, "State"),
                "Bill To Zip": _text(bill, "PostalCode"),
                "Active": _bool(c, "IsActive"),
            })
        return pd.DataFrame(rows)

    def extract_vendors(self) -> pd.DataFrame:
        log.info("Extracting Vendors")
        root = self._send(_qbxml_envelope(
            '<VendorQueryRq requestID="1"><ActiveStatus>All</ActiveStatus></VendorQueryRq>'
        ))
        rows = []
        for v in root.iter("VendorRet"):
            rows.append({
                "Name": _text(v, "Name"),
                "Company Name": _text(v, "CompanyName"),
                "Main Email": _text(v, "Email"),
                "Main Phone": _text(v, "Phone"),
                "1099": "Yes" if _bool(v, "IsVendorEligibleFor1099") else "No",
                "Tax ID": _text(v, "VendorTaxIdent"),
                "Active": _bool(v, "IsActive"),
            })
        return pd.DataFrame(rows)

    def extract_items(self) -> pd.DataFrame:
        """Queries Service, Inventory, and Non-Inventory items."""
        log.info("Extracting Items (Service, Inventory, Non-Inventory)")
        queries = [
            ("ItemServiceQueryRq", "ItemServiceRet", "Service"),
            ("ItemInventoryQueryRq", "ItemInventoryRet", "Inventory Part"),
            ("ItemNonInventoryQueryRq", "ItemNonInventoryRet", "Non-inventory Part"),
            ("ItemOtherChargeQueryRq", "ItemOtherChargeRet", "Other Charge"),
        ]
        rows = []
        for rq, ret, type_label in queries:
            try:
                root = self._send(_qbxml_envelope(f'<{rq} requestID="1"/>'))
                for i in root.iter(ret):
                    # Different item types store price in different fields
                    price = (
                        _text(i.find("SalesOrPurchase"), "Price")
                        or _text(i.find("SalesAndPurchase"), "SalesPrice")
                        or _text(i, "SalesPrice")
                    )
                    rows.append({
                        "Item": _text(i, "FullName") or _text(i, "Name"),
                        "Type": type_label,
                        "Price": price,
                        "Description": (
                            _text(i.find("SalesOrPurchase"), "Desc")
                            or _text(i.find("SalesAndPurchase"), "SalesDesc")
                            or _text(i, "SalesDesc")
                        ),
                        "Active": _bool(i, "IsActive"),
                    })
            except Exception as e:
                log.warning("Item query %s failed: %s", rq, e)
        return pd.DataFrame(rows)

    def extract_employees(self) -> pd.DataFrame:
        log.info("Extracting Employees")
        root = self._send(_qbxml_envelope(
            '<EmployeeQueryRq requestID="1"><ActiveStatus>All</ActiveStatus></EmployeeQueryRq>'
        ))
        rows = []
        for e in root.iter("EmployeeRet"):
            rows.append({
                "Name": _text(e, "Name"),
                "SSN": _text(e, "SSN"),
                "Email": _text(e, "Email"),
                "Phone": _text(e, "Phone"),
                "Active": _bool(e, "IsActive"),
            })
        return pd.DataFrame(rows)

    def extract_open_invoices(self) -> pd.DataFrame:
        """Open (unpaid) AR invoices."""
        log.info("Extracting Open AR Invoices")
        root = self._send(_qbxml_envelope(
            '<InvoiceQueryRq requestID="1">'
            '<PaidStatus>NotPaidOnly</PaidStatus>'
            '</InvoiceQueryRq>'
        ))
        rows = []
        for inv in root.iter("InvoiceRet"):
            rows.append({
                "Customer": _text(inv.find("CustomerRef"), "FullName"),
                "Num": _text(inv, "RefNumber"),
                "Date": _text(inv, "TxnDate"),
                "Due Date": _text(inv, "DueDate"),
                "Amount": _text(inv, "Subtotal") or _text(inv, "BalanceRemaining"),
                "Open Balance": _text(inv, "BalanceRemaining"),
                "Terms": _text(inv.find("TermsRef"), "FullName"),
            })
        return pd.DataFrame(rows)

    def extract_open_bills(self) -> pd.DataFrame:
        """Open (unpaid) AP bills."""
        log.info("Extracting Open AP Bills")
        root = self._send(_qbxml_envelope(
            '<BillQueryRq requestID="1">'
            '<PaidStatus>NotPaidOnly</PaidStatus>'
            '</BillQueryRq>'
        ))
        rows = []
        for b in root.iter("BillRet"):
            rows.append({
                "Vendor": _text(b.find("VendorRef"), "FullName"),
                "Ref Number": _text(b, "RefNumber"),
                "Date": _text(b, "TxnDate"),
                "Due Date": _text(b, "DueDate"),
                "Amount": _text(b, "AmountDue"),
                "Open Balance": _text(b, "OpenAmount") or _text(b, "AmountDue"),
            })
        return pd.DataFrame(rows)

    def extract_trial_balance(self) -> pd.DataFrame:
        """Trial balance report. Returns a flat Account/Debit/Credit frame."""
        log.info("Extracting Trial Balance")
        try:
            root = self._send(_qbxml_envelope(
                '<GeneralSummaryReportQueryRq requestID="1">'
                '<GeneralSummaryReportType>TrialBalance</GeneralSummaryReportType>'
                '</GeneralSummaryReportQueryRq>'
            ))
        except Exception as e:
            log.warning("Trial balance query failed: %s", e)
            return pd.DataFrame()

        rows = []
        # Report structure is deeply nested; we walk DataRow elements and take
        # the account name + first two numeric column values as Debit/Credit.
        for dr in root.iter("DataRow"):
            cols = [c.text or "" for c in dr.findall("ColData")]
            if len(cols) >= 3:
                try:
                    debit = float(cols[1].replace(",", "") or 0)
                except ValueError:
                    debit = 0
                try:
                    credit = float(cols[2].replace(",", "") or 0)
                except ValueError:
                    credit = 0
                rows.append({"Account": cols[0], "Debit": debit, "Credit": credit})
        return pd.DataFrame(rows)

    # ── Convenience: run everything to Excel ──────────────────

    def extract_all(
        self,
        output_dir: Path,
        progress: Callable[[str, str], None] | None = None,
    ) -> dict[str, int]:
        """Extract every supported entity and write the standard Excel files.

        Returns a map of entity → row count. `progress(name, status)` is called
        before and after each entity so the UI can stream updates.
        """
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        jobs: list[tuple[str, str, Callable[[], pd.DataFrame]]] = [
            ("accounts",      "QBD_ChartOfAccounts.xlsx", self.extract_accounts),
            ("customers",     "QBD_Customers.xlsx",       self.extract_customers),
            ("vendors",       "QBD_Vendors.xlsx",         self.extract_vendors),
            ("items",         "QBD_Items.xlsx",           self.extract_items),
            ("employees",     "QBD_Employees.xlsx",       self.extract_employees),
            ("open_invoices", "QBD_OpenInvoices.xlsx",    self.extract_open_invoices),
            ("open_bills",    "QBD_OpenBills.xlsx",       self.extract_open_bills),
            ("trial_balance", "QBD_TrialBalance.xlsx",    self.extract_trial_balance),
        ]
        counts: dict[str, int] = {}
        for name, filename, fn in jobs:
            if progress:
                progress(name, "running")
            try:
                df = fn()
                path = output_dir / filename
                df.to_excel(path, index=False)
                counts[name] = len(df)
                log.info("Wrote %s (%d rows) → %s", name, len(df), path)
                if progress:
                    progress(name, f"done:{len(df)}")
            except Exception as e:
                log.exception("Failed to extract %s", name)
                counts[name] = -1
                if progress:
                    progress(name, f"error:{e}")
        return counts

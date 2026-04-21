"""
QBD → QBO migration agent.

Reads QBD Excel exports, maps them to QBO REST v3 entities, and loads them in
dependency order: Accounts → Terms → Vendors → Customers → Jobs → Items →
Invoices → Bills.

Designed to run end-to-end from the Flask UI in app.py, but the phase classes
(FileExtractor, Transformer, QBOLoader, Validator) can also be called directly.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd

from qbo_client import QBOClient, QBOError

log = logging.getLogger("qbo.migration")


# ─── QBD → QBO type mappings ──────────────────────────────────────

# QBD account "Type" string → (QBO AccountType, AccountSubType)
ACCOUNT_TYPE_MAP: dict[str, tuple[str, str]] = {
    "Bank": ("Bank", "Checking"),
    "Accounts Receivable": ("Accounts Receivable", "AccountsReceivable"),
    "Other Current Asset": ("Other Current Asset", "OtherCurrentAssets"),
    "Fixed Asset": ("Fixed Asset", "Machinery and Equipment"),
    "Other Asset": ("Other Asset", "OtherLongTermAssets"),
    "Accounts Payable": ("Accounts Payable", "AccountsPayable"),
    "Credit Card": ("Credit Card", "CreditCard"),
    "Other Current Liability": ("Other Current Liability", "OtherCurrentLiabilities"),
    "Long Term Liability": ("Long Term Liability", "OtherLongTermLiabilities"),
    "Equity": ("Equity", "OpeningBalanceEquity"),
    "Income": ("Income", "SalesOfProductIncome"),
    "Other Income": ("Other Income", "OtherPrimaryIncome"),
    "Cost of Goods Sold": ("Cost of Goods Sold", "SuppliesMaterialsCogs"),
    "Expense": ("Expense", "OtherBusinessExpenses"),
    "Other Expense": ("Other Expense", "OtherMiscellaneousExpense"),
}

ITEM_TYPE_MAP: dict[str, str] = {
    "Service": "Service",
    "Inventory Part": "Inventory",
    "Non-inventory Part": "NonInventory",
    "Other Charge": "NonInventory",
}


# ─── Domain objects ───────────────────────────────────────────────

@dataclass
class Customer:
    qbd_name: str
    display_name: str
    company: str = ""
    email: str = ""
    phone: str = ""
    billing_addr: dict = field(default_factory=dict)
    terms_name: str = ""
    qbo_id: str = ""  # filled after create


@dataclass
class Job:
    qbd_name: str           # "Parent:Child"
    parent_qbd_name: str    # "Parent"
    child_name: str         # "Child"
    status: str = ""
    qbo_id: str = ""


# ─── Phase 1: Extract from QBD Excel ──────────────────────────────

class FileExtractor:
    """Reads QBD Excel export files into pandas DataFrames."""

    def extract_accounts(self, path: str) -> pd.DataFrame:
        log.info("Reading accounts from %s", path)
        df = pd.read_excel(path)
        return df.fillna("")

    def extract_customers(self, path: str) -> tuple[list[Customer], list[Job]]:
        log.info("Reading customers from %s", path)
        df = pd.read_excel(path).fillna("")
        customers: dict[str, Customer] = {}
        jobs: list[Job] = []

        for _, row in df.iterrows():
            name = str(row.get("Name", "")).strip()
            if not name:
                continue
            if ":" in name:
                parent, child = name.split(":", 1)
                jobs.append(Job(
                    qbd_name=name,
                    parent_qbd_name=parent.strip(),
                    child_name=child.strip(),
                    status=str(row.get("Job Status", "")),
                ))
            else:
                customers[name] = Customer(
                    qbd_name=name,
                    display_name=name,
                    company=str(row.get("Company Name", "")),
                    email=str(row.get("Main Email", "")),
                    phone=str(row.get("Main Phone", "")),
                    terms_name=str(row.get("Terms", "")),
                    billing_addr=self._parse_address(row),
                )
        return list(customers.values()), jobs

    def extract_vendors(self, path: str) -> pd.DataFrame:
        log.info("Reading vendors from %s", path)
        return pd.read_excel(path).fillna("")

    def extract_items(self, path: str) -> pd.DataFrame:
        log.info("Reading items from %s", path)
        return pd.read_excel(path).fillna("")

    def extract_invoices(self, path: str) -> pd.DataFrame:
        log.info("Reading open invoices from %s", path)
        return pd.read_excel(path).fillna("")

    def extract_bills(self, path: str) -> pd.DataFrame:
        log.info("Reading open bills from %s", path)
        return pd.read_excel(path).fillna("")

    def _parse_address(self, row) -> dict:
        return {
            "Line1": str(row.get("Bill To 1", "")),
            "City": str(row.get("Bill To City", "")),
            "CountrySubDivisionCode": str(row.get("Bill To State", "")),
            "PostalCode": str(row.get("Bill To Zip", "")),
        }


# ─── Phase 2: Transform ───────────────────────────────────────────

class Transformer:
    """Maps QBD concepts to QBO shapes and warns on missing data."""

    def __init__(self):
        self.warnings: list[str] = []
        self.unique_terms: set[str] = set()

    def collect_terms(self, customers: list[Customer]) -> list[str]:
        for c in customers:
            if c.terms_name:
                self.unique_terms.add(c.terms_name)
        return sorted(self.unique_terms)

    def account_to_qbo(self, row: pd.Series) -> dict:
        qbd_type = str(row.get("Type", "")).strip()
        name = str(row.get("Account", row.get("Name", ""))).strip()
        acct_type, sub_type = ACCOUNT_TYPE_MAP.get(qbd_type, ("Other Current Asset", "OtherCurrentAssets"))
        if qbd_type not in ACCOUNT_TYPE_MAP:
            self.warnings.append(f"Unknown QBD account type '{qbd_type}' for '{name}' — defaulted to OtherCurrentAssets")
        return {
            "Name": name[:100],
            "AccountType": acct_type,
            "AccountSubType": sub_type,
        }

    def customer_to_qbo(self, c: Customer, term_id_map: dict[str, str]) -> dict:
        body: dict[str, Any] = {
            "DisplayName": c.display_name[:100],
            "CompanyName": c.company[:50] if c.company else "",
        }
        if c.email:
            body["PrimaryEmailAddr"] = {"Address": c.email}
        if c.phone:
            body["PrimaryPhone"] = {"FreeFormNumber": c.phone}
        if c.billing_addr.get("Line1"):
            body["BillAddr"] = c.billing_addr
        if c.terms_name and c.terms_name in term_id_map:
            body["SalesTermRef"] = {"value": term_id_map[c.terms_name]}
        return body

    def job_to_qbo(self, j: Job, parent_id_map: dict[str, str]) -> dict | None:
        parent_id = parent_id_map.get(j.parent_qbd_name)
        if not parent_id:
            self.warnings.append(f"Job '{j.qbd_name}' — parent customer '{j.parent_qbd_name}' not found, skipping")
            return None
        return {
            "DisplayName": j.child_name[:100],
            "Job": True,
            "ParentRef": {"value": parent_id},
        }

    def vendor_to_qbo(self, row: pd.Series) -> dict:
        name = str(row.get("Name", "")).strip()
        body = {"DisplayName": name[:100]}
        company = str(row.get("Company Name", "")).strip()
        if company:
            body["CompanyName"] = company[:50]
        email = str(row.get("Main Email", "")).strip()
        if email:
            body["PrimaryEmailAddr"] = {"Address": email}
        phone = str(row.get("Main Phone", "")).strip()
        if phone:
            body["PrimaryPhone"] = {"FreeFormNumber": phone}
        if str(row.get("1099", "")).strip().lower() in ("yes", "true", "1"):
            body["Vendor1099"] = True
        return body

    def item_to_qbo(self, row: pd.Series, income_account_id: str) -> dict:
        name = str(row.get("Item", row.get("Name", ""))).strip()
        qbd_type = str(row.get("Type", "Service")).strip()
        qbo_type = ITEM_TYPE_MAP.get(qbd_type, "Service")
        body = {
            "Name": name[:100],
            "Type": qbo_type,
            "IncomeAccountRef": {"value": income_account_id},
        }
        price = row.get("Price")
        if price not in (None, "", "nan"):
            try:
                body["UnitPrice"] = float(price)
            except (ValueError, TypeError):
                pass
        return body


# ─── Phase 3: Load to QBO ─────────────────────────────────────────

class QBOLoader:
    """Creates entities in QBO via REST API v3."""

    def __init__(self, client: QBOClient, dry_run: bool = False):
        self.client = client
        self.dry_run = dry_run
        self.summary: dict[str, int] = {}
        self.errors: list[str] = []

    def create_accounts(self, accounts_df: pd.DataFrame, transformer: Transformer) -> dict[str, str]:
        """Returns map of account Name → QBO Id."""
        log.info("Creating %d accounts in QBO%s", len(accounts_df), " (dry run)" if self.dry_run else "")
        id_map: dict[str, str] = {}
        count = 0
        for _, row in accounts_df.iterrows():
            body = transformer.account_to_qbo(row)
            if not body["Name"]:
                continue
            try:
                if self.dry_run:
                    id_map[body["Name"]] = f"dry-{count}"
                else:
                    resp = self.client.create("account", body)
                    created = resp.get("Account", {})
                    id_map[body["Name"]] = created.get("Id", "")
                count += 1
            except QBOError as e:
                self.errors.append(f"Account '{body['Name']}': {e}")
                log.warning("Account create failed: %s", e)
        self.summary["accounts"] = count
        return id_map

    def create_terms(self, terms: list[str]) -> dict[str, str]:
        log.info("Creating %d terms in QBO%s", len(terms), " (dry run)" if self.dry_run else "")
        id_map: dict[str, str] = {}
        for i, t in enumerate(terms):
            body = self._term_body(t)
            try:
                if self.dry_run:
                    id_map[t] = f"dry-term-{i}"
                else:
                    resp = self.client.create("term", body)
                    id_map[t] = resp.get("Term", {}).get("Id", "")
            except QBOError as e:
                self.errors.append(f"Term '{t}': {e}")
        self.summary["terms"] = len(id_map)
        return id_map

    def create_customers(self, customers: list[Customer], transformer: Transformer,
                         term_id_map: dict[str, str]) -> dict[str, str]:
        log.info("Creating %d customers in QBO%s", len(customers), " (dry run)" if self.dry_run else "")
        id_map: dict[str, str] = {}
        for i, c in enumerate(customers):
            body = transformer.customer_to_qbo(c, term_id_map)
            try:
                if self.dry_run:
                    c.qbo_id = f"dry-cust-{i}"
                else:
                    resp = self.client.create("customer", body)
                    c.qbo_id = resp.get("Customer", {}).get("Id", "")
                id_map[c.qbd_name] = c.qbo_id
            except QBOError as e:
                self.errors.append(f"Customer '{c.display_name}': {e}")
        self.summary["customers"] = len(id_map)
        return id_map

    def create_jobs(self, jobs: list[Job], transformer: Transformer,
                    customer_id_map: dict[str, str]) -> dict[str, str]:
        log.info("Creating %d jobs (sub-customers) in QBO%s", len(jobs), " (dry run)" if self.dry_run else "")
        id_map: dict[str, str] = {}
        for i, j in enumerate(jobs):
            body = transformer.job_to_qbo(j, customer_id_map)
            if not body:
                continue
            try:
                if self.dry_run:
                    j.qbo_id = f"dry-job-{i}"
                else:
                    resp = self.client.create("customer", body)
                    j.qbo_id = resp.get("Customer", {}).get("Id", "")
                id_map[j.qbd_name] = j.qbo_id
            except QBOError as e:
                self.errors.append(f"Job '{j.qbd_name}': {e}")
        self.summary["jobs"] = len(id_map)
        return id_map

    def create_vendors(self, vendors_df: pd.DataFrame, transformer: Transformer) -> dict[str, str]:
        log.info("Creating %d vendors in QBO%s", len(vendors_df), " (dry run)" if self.dry_run else "")
        id_map: dict[str, str] = {}
        for i, row in vendors_df.iterrows():
            body = transformer.vendor_to_qbo(row)
            if not body["DisplayName"]:
                continue
            try:
                if self.dry_run:
                    id_map[body["DisplayName"]] = f"dry-vend-{i}"
                else:
                    resp = self.client.create("vendor", body)
                    id_map[body["DisplayName"]] = resp.get("Vendor", {}).get("Id", "")
            except QBOError as e:
                self.errors.append(f"Vendor '{body['DisplayName']}': {e}")
        self.summary["vendors"] = len(id_map)
        return id_map

    def create_items(self, items_df: pd.DataFrame, transformer: Transformer,
                     income_account_id: str) -> dict[str, str]:
        log.info("Creating %d items in QBO%s", len(items_df), " (dry run)" if self.dry_run else "")
        id_map: dict[str, str] = {}
        if not income_account_id:
            log.warning("No income account available — skipping item creation")
            return id_map
        for i, row in items_df.iterrows():
            body = transformer.item_to_qbo(row, income_account_id)
            if not body["Name"]:
                continue
            try:
                if self.dry_run:
                    id_map[body["Name"]] = f"dry-item-{i}"
                else:
                    resp = self.client.create("item", body)
                    id_map[body["Name"]] = resp.get("Item", {}).get("Id", "")
            except QBOError as e:
                self.errors.append(f"Item '{body['Name']}': {e}")
        self.summary["items"] = len(id_map)
        return id_map

    def create_invoices(self, invoices_df: pd.DataFrame, customer_id_map: dict[str, str],
                        item_id_map: dict[str, str]) -> int:
        log.info("Creating %d invoices in QBO%s", len(invoices_df), " (dry run)" if self.dry_run else "")
        count = 0
        for _, row in invoices_df.iterrows():
            cust_name = str(row.get("Customer", "")).strip()
            cust_id = customer_id_map.get(cust_name)
            if not cust_id:
                self.errors.append(f"Invoice skipped — customer '{cust_name}' not in QBO")
                continue
            amount = float(row.get("Amount", 0) or 0)
            if amount <= 0:
                continue
            body = {
                "CustomerRef": {"value": cust_id},
                "Line": [{
                    "DetailType": "SalesItemLineDetail",
                    "Amount": amount,
                    "SalesItemLineDetail": {
                        "ItemRef": {"value": next(iter(item_id_map.values()), "1")},
                    },
                }],
            }
            doc_num = str(row.get("Num", "")).strip()
            if doc_num:
                body["DocNumber"] = doc_num[:21]
            try:
                if not self.dry_run:
                    self.client.create("invoice", body)
                count += 1
            except QBOError as e:
                self.errors.append(f"Invoice '{doc_num}': {e}")
        self.summary["invoices"] = count
        return count

    def create_bills(self, bills_df: pd.DataFrame, vendor_id_map: dict[str, str],
                     expense_account_id: str) -> int:
        log.info("Creating %d bills in QBO%s", len(bills_df), " (dry run)" if self.dry_run else "")
        count = 0
        if not expense_account_id:
            log.warning("No expense account available — skipping bills")
            return 0
        for _, row in bills_df.iterrows():
            vend_name = str(row.get("Vendor", "")).strip()
            vend_id = vendor_id_map.get(vend_name)
            if not vend_id:
                self.errors.append(f"Bill skipped — vendor '{vend_name}' not in QBO")
                continue
            amount = float(row.get("Open Balance", row.get("Amount", 0)) or 0)
            if amount <= 0:
                continue
            body = {
                "VendorRef": {"value": vend_id},
                "Line": [{
                    "DetailType": "AccountBasedExpenseLineDetail",
                    "Amount": amount,
                    "AccountBasedExpenseLineDetail": {
                        "AccountRef": {"value": expense_account_id},
                    },
                }],
            }
            try:
                if not self.dry_run:
                    self.client.create("bill", body)
                count += 1
            except QBOError as e:
                self.errors.append(f"Bill for '{vend_name}': {e}")
        self.summary["bills"] = count
        return count

    def get_summary(self) -> dict:
        return {"counts": dict(self.summary), "errors": list(self.errors)}

    def _term_body(self, term_name: str) -> dict:
        """Parse 'Net 30', '2% 10 Net 30', etc. into a QBO Term."""
        body = {"Name": term_name[:31], "Type": "STANDARD"}
        m = re.search(r"Net\s+(\d+)", term_name, re.IGNORECASE)
        if m:
            body["DueDays"] = int(m.group(1))
        disc = re.search(r"(\d+(?:\.\d+)?)\s*%\s*(\d+)", term_name)
        if disc:
            body["DiscountPercent"] = float(disc.group(1))
            body["DiscountDays"] = int(disc.group(2))
        return body


# ─── Phase 4: Validate ────────────────────────────────────────────

class Validator:
    """Compares QBO entity counts against the source data."""

    def __init__(self, client: QBOClient):
        self.client = client

    def validate(self, expected: dict[str, int]) -> dict[str, dict]:
        results: dict[str, dict] = {}
        entity_map = {
            "accounts": "Account",
            "customers": "Customer",
            "vendors": "Vendor",
            "items": "Item",
            "invoices": "Invoice",
            "bills": "Bill",
        }
        for key, entity in entity_map.items():
            if key not in expected:
                continue
            try:
                actual = self.client.count(entity)
                results[key] = {
                    "expected": expected[key],
                    "actual": actual,
                    "ok": actual >= expected[key],
                }
            except QBOError as e:
                results[key] = {"expected": expected[key], "actual": 0, "error": str(e), "ok": False}
        return results

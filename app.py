"""
QBD → QuickBooks Online Migration Connector.

Flask UI wrapping the qbo_migration agent. Handles OAuth 2.0 with Intuit,
QBD Excel file uploads, dry-run validation, and live migrations with
streaming logs over SSE.

Run: python app.py
"""

from __future__ import annotations

import json
import logging
import os
import queue
import secrets
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

from flask import (
    Flask,
    Response,
    jsonify,
    redirect,
    render_template,
    request,
    session,
    url_for,
)

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from qbo_client import QBOClient, QBOError
from qbd_extractor import QBDExtractor, QBDConnectionInfo, QBDUnavailable, IS_WINDOWS
import qbo_migration as qm

# ─── Config ──────────────────────────────────────────────────────

BASE_DIR = Path(__file__).parent
EXPORT_DIR = BASE_DIR / "qbd_exports"
WORK_DIR = BASE_DIR / "migration_workspace"
EXPORT_DIR.mkdir(exist_ok=True)
WORK_DIR.mkdir(exist_ok=True)

CLIENT_ID = os.environ.get("QBO_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("QBO_CLIENT_SECRET", "")
REDIRECT_URI = os.environ.get("QBO_REDIRECT_URI", "http://localhost:5050/callback")
ENVIRONMENT = os.environ.get("QBO_ENVIRONMENT", "sandbox")
SCOPES = os.environ.get("QBO_SCOPES", "com.intuit.quickbooks.accounting")
DEMO_MODE = not (CLIENT_ID and CLIENT_SECRET)

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", secrets.token_hex(16))

# Singleton QBO client (token file persists across restarts)
qbo = QBOClient(
    client_id=CLIENT_ID,
    client_secret=CLIENT_SECRET,
    redirect_uri=REDIRECT_URI,
    environment=ENVIRONMENT,
    token_store=BASE_DIR / "tokens.json",
)

# ─── SSE log queue ────────────────────────────────────────────────

log_queue: queue.Queue = queue.Queue()
migration_status = {
    "running": False,
    "phase": "",
    "summary": None,
    "error": None,
    "started_at": None,
}


class QueueLogHandler(logging.Handler):
    def emit(self, record):
        msg = self.format(record)
        log_queue.put({"type": "log", "level": record.levelname, "message": msg})


# ─── Routes: UI ───────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template(
        "index.html",
        demo_mode=DEMO_MODE,
        environment=ENVIRONMENT,
    )


# ─── Routes: OAuth ────────────────────────────────────────────────

@app.route("/connect")
def connect():
    if DEMO_MODE:
        # Fake a connection for presentation demos
        qbo.tokens.access_token = "demo-token"
        qbo.tokens.refresh_token = "demo-refresh"
        qbo.tokens.realm_id = "1234567890"
        qbo.tokens.expires_at = time.time() + 3600
        qbo._save_tokens()
        return redirect(url_for("index") + "?connected=demo")

    state = secrets.token_urlsafe(16)
    session["oauth_state"] = state
    return redirect(qbo.authorize_url(state=state, scopes=SCOPES))


@app.route("/callback")
def callback():
    state = request.args.get("state", "")
    code = request.args.get("code", "")
    realm_id = request.args.get("realmId", "")

    if state != session.get("oauth_state"):
        return "State mismatch — possible CSRF", 400
    if not code or not realm_id:
        return f"Missing code or realmId. Error: {request.args.get('error', 'unknown')}", 400

    try:
        qbo.exchange_code(code, realm_id)
    except QBOError as e:
        return f"Token exchange failed: {e}", 500

    return redirect(url_for("index") + "?connected=1")


@app.route("/api/disconnect", methods=["POST"])
def disconnect():
    qbo.disconnect()
    return jsonify({"success": True})


@app.route("/api/status")
def status():
    connected = qbo.is_connected()
    company = ""
    if connected and not DEMO_MODE:
        try:
            info = qbo.company_info()
            company = info.get("CompanyName", "")
        except QBOError:
            pass
    elif DEMO_MODE and connected:
        company = "Demo Company (Presentation Mode)"

    return jsonify({
        "connected": connected,
        "demo_mode": DEMO_MODE,
        "environment": ENVIRONMENT,
        "realm_id": qbo.tokens.realm_id,
        "company": company,
    })


# ─── Routes: QBD files ───────────────────────────────────────────

EXPECTED_FILES = [
    "QBD_ChartOfAccounts.xlsx",
    "QBD_Customers.xlsx",
    "QBD_Vendors.xlsx",
    "QBD_Items.xlsx",
    "QBD_Employees.xlsx",
    "QBD_OpenInvoices.xlsx",
    "QBD_OpenBills.xlsx",
    "QBD_TrialBalance.xlsx",
]


@app.route("/api/check-qbd")
def check_qbd():
    files = {}
    for fname in EXPECTED_FILES:
        p = EXPORT_DIR / fname
        if p.exists():
            st = p.stat()
            files[fname] = {
                "exists": True,
                "size": st.st_size,
                "modified": datetime.fromtimestamp(st.st_mtime).isoformat(),
            }
        else:
            files[fname] = {"exists": False}
    found = sum(1 for f in files.values() if f["exists"])
    return jsonify({"files": files, "found": found, "total": len(EXPECTED_FILES)})


@app.route("/api/upload", methods=["POST"])
def upload():
    if "files" not in request.files:
        return jsonify({"success": False, "message": "No files provided"}), 400
    uploaded = []
    for f in request.files.getlist("files"):
        if f.filename and f.filename.endswith(".xlsx"):
            f.save(str(EXPORT_DIR / f.filename))
            uploaded.append(f.filename)
    return jsonify({"success": True, "files": uploaded,
                    "message": f"Uploaded {len(uploaded)} file(s)"})


# ─── Routes: direct QBD connection ───────────────────────────────

# Module-level state for the QBD session (singleton — one company open at a time)
_qbd_state: dict = {
    "extractor": None,
    "connected": False,
    "company_file": "",
    "error": "",
}

# Extraction progress, streamed back over the same SSE bus as migrations
_qbd_progress: dict = {"running": False, "entities": {}}


@app.route("/api/qbd/platform")
def qbd_platform():
    """Tell the UI whether direct QBD connection is even possible on this host."""
    return jsonify({
        "windows": IS_WINDOWS,
        "supported": IS_WINDOWS,
        "message": (
            "Direct QBD connection available" if IS_WINDOWS
            else "Direct QBD connection requires Windows + QuickBooks SDK. "
                 "On this host, use the Upload Files option instead."
        ),
    })


@app.route("/api/qbd/status")
def qbd_status():
    return jsonify({
        "connected": _qbd_state["connected"],
        "company_file": _qbd_state["company_file"],
        "windows": IS_WINDOWS,
        "error": _qbd_state["error"],
    })


@app.route("/api/qbd/connect", methods=["POST"])
def qbd_connect():
    """Open a QBXMLRP2 session to the currently-open QBD company file
    (or to a file path if provided)."""
    data = request.get_json(silent=True) or {}
    company_file = (data.get("company_file") or "").strip()

    # Close any stale session first
    if _qbd_state["extractor"]:
        try:
            _qbd_state["extractor"].close()
        except Exception:
            pass

    info = QBDConnectionInfo(
        app_name="QBD-QBO Migration",
        app_id="",
        company_file=company_file,
        connection_mode=1,
    )
    ex = QBDExtractor(info)
    try:
        meta = ex.open()
        _qbd_state.update({
            "extractor": ex,
            "connected": True,
            "company_file": meta.get("company_file", ""),
            "error": "",
        })
        return jsonify({"success": True, **meta})
    except QBDUnavailable as e:
        _qbd_state.update({"extractor": None, "connected": False, "error": str(e)})
        return jsonify({"success": False, "message": str(e), "windows": IS_WINDOWS}), 400
    except Exception as e:
        log.exception("QBD connect failed")
        _qbd_state.update({"extractor": None, "connected": False, "error": str(e)})
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/api/qbd/disconnect", methods=["POST"])
def qbd_disconnect():
    ex = _qbd_state.get("extractor")
    if ex:
        try:
            ex.close()
        except Exception:
            pass
    _qbd_state.update({"extractor": None, "connected": False, "company_file": "", "error": ""})
    return jsonify({"success": True})


@app.route("/api/qbd/extract", methods=["POST"])
def qbd_extract():
    """Stream the full extraction to Excel files in qbd_exports/.
    Progress is pushed into the same log_queue the migration uses, so the
    existing log console in the UI shows it in real time."""
    if _qbd_progress["running"]:
        return jsonify({"success": False, "message": "Extraction already running"}), 409
    if not _qbd_state["connected"] or not _qbd_state["extractor"]:
        return jsonify({"success": False, "message": "Connect to QBD first"}), 400

    t = threading.Thread(target=_run_qbd_extract, daemon=True)
    t.start()
    return jsonify({"success": True})


def _run_qbd_extract():
    _qbd_progress["running"] = True
    _qbd_progress["entities"] = {}

    # Drain queue so the console starts clean
    while not log_queue.empty():
        try:
            log_queue.get_nowait()
        except queue.Empty:
            break

    handler = QueueLogHandler()
    handler.setFormatter(logging.Formatter("%(message)s"))
    from qbd_extractor import log as qbd_log
    qbd_log.addHandler(handler)
    qbd_log.setLevel(logging.INFO)

    _emit("phase", "Extracting directly from QuickBooks Desktop...", phase="qbd-extract")

    def progress(name: str, status: str):
        _qbd_progress["entities"][name] = status
        _emit("qbd-progress", f"{name}: {status}", name=name, status=status)

    try:
        ex = _qbd_state["extractor"]
        counts = ex.extract_all(EXPORT_DIR, progress=progress)
        _emit(
            "qbd-complete",
            f"Extracted {sum(c for c in counts.values() if c >= 0)} total rows across "
            f"{sum(1 for c in counts.values() if c >= 0)} entities.",
            counts=counts,
        )
    except Exception as e:
        log.exception("QBD extraction failed")
        _emit("error", f"QBD extraction failed: {e}")
    finally:
        _qbd_progress["running"] = False
        qbd_log.removeHandler(handler)


@app.route("/api/generate-sample", methods=["POST"])
def generate_sample():
    """Write small dummy QBD Excel files for demo presentations."""
    try:
        import pandas as pd

        pd.DataFrame([
            {"Account": "Checking", "Type": "Bank"},
            {"Account": "Accounts Receivable", "Type": "Accounts Receivable"},
            {"Account": "Accounts Payable", "Type": "Accounts Payable"},
            {"Account": "Sales Income", "Type": "Income"},
            {"Account": "Job Materials", "Type": "Cost of Goods Sold"},
            {"Account": "Office Expense", "Type": "Expense"},
            {"Account": "Opening Balance Equity", "Type": "Equity"},
        ]).to_excel(EXPORT_DIR / "QBD_ChartOfAccounts.xlsx", index=False)

        pd.DataFrame([
            {"Name": "Acme Construction", "Company Name": "Acme Inc", "Main Email": "ap@acme.com",
             "Main Phone": "555-0101", "Terms": "Net 30"},
            {"Name": "Acme Construction:Office Remodel", "Job Status": "In Progress"},
            {"Name": "Ridgewood Partners", "Company Name": "Ridgewood LLC",
             "Main Email": "billing@ridgewood.com", "Terms": "Net 15"},
        ]).to_excel(EXPORT_DIR / "QBD_Customers.xlsx", index=False)

        pd.DataFrame([
            {"Name": "Home Depot", "Company Name": "The Home Depot", "1099": "No"},
            {"Name": "Bob the Electrician", "1099": "Yes"},
        ]).to_excel(EXPORT_DIR / "QBD_Vendors.xlsx", index=False)

        pd.DataFrame([
            {"Item": "Labor", "Type": "Service", "Price": 85.00},
            {"Item": "2x4 Lumber", "Type": "Non-inventory Part", "Price": 4.25},
        ]).to_excel(EXPORT_DIR / "QBD_Items.xlsx", index=False)

        pd.DataFrame([
            {"Name": "Jane Smith", "SSN": "XXX-XX-1234"},
        ]).to_excel(EXPORT_DIR / "QBD_Employees.xlsx", index=False)

        pd.DataFrame([
            {"Customer": "Acme Construction", "Num": "INV-1001", "Amount": 2500.00},
            {"Customer": "Ridgewood Partners", "Num": "INV-1002", "Amount": 875.50},
        ]).to_excel(EXPORT_DIR / "QBD_OpenInvoices.xlsx", index=False)

        pd.DataFrame([
            {"Vendor": "Home Depot", "Open Balance": 340.15},
            {"Vendor": "Bob the Electrician", "Open Balance": 1200.00},
        ]).to_excel(EXPORT_DIR / "QBD_OpenBills.xlsx", index=False)

        pd.DataFrame([
            {"Account": "Checking", "Debit": 15000, "Credit": 0},
            {"Account": "Accounts Receivable", "Debit": 3375.50, "Credit": 0},
            {"Account": "Accounts Payable", "Debit": 0, "Credit": 1540.15},
            {"Account": "Opening Balance Equity", "Debit": 0, "Credit": 16835.35},
        ]).to_excel(EXPORT_DIR / "QBD_TrialBalance.xlsx", index=False)

        return jsonify({"success": True, "message": "Sample QBD data generated"})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


# ─── Routes: migration ───────────────────────────────────────────

@app.route("/api/migrate", methods=["POST"])
def start_migration():
    if migration_status["running"]:
        return jsonify({"success": False, "message": "Migration already running"}), 409
    if not qbo.is_connected():
        return jsonify({"success": False, "message": "Connect to QuickBooks Online first"}), 400

    dry_run = bool(request.json.get("dry_run", True)) if request.is_json else True
    t = threading.Thread(target=_run_migration, args=(dry_run,), daemon=True)
    t.start()
    return jsonify({"success": True})


@app.route("/api/migrate/stream")
def migration_stream():
    def generate():
        # replay any pending items first
        while True:
            try:
                item = log_queue.get(timeout=30)
                yield f"data: {json.dumps(item)}\n\n"
                if item.get("type") in ("complete", "error"):
                    break
            except queue.Empty:
                yield f"data: {json.dumps({'type': 'heartbeat'})}\n\n"

    return Response(
        generate(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/api/migrate/status")
def migration_state():
    return jsonify(migration_status)


def _emit(type_: str, message: str, **extra):
    log_queue.put({"type": type_, "message": message, **extra})


def _run_migration(dry_run: bool):
    """Background worker: runs the full 4-phase migration."""
    migration_status.update({
        "running": True,
        "phase": "starting",
        "summary": None,
        "error": None,
        "started_at": datetime.utcnow().isoformat(),
    })

    # Clear stale queue
    while not log_queue.empty():
        try:
            log_queue.get_nowait()
        except queue.Empty:
            break

    handler = QueueLogHandler()
    handler.setFormatter(logging.Formatter("%(message)s"))
    qm.log.addHandler(handler)
    qm.log.setLevel(logging.INFO)

    try:
        # ── Phase 1: Extract ──
        _emit("phase", "Phase 1: Extracting from QBD Excel exports...", phase="extract")
        migration_status["phase"] = "extract"

        extractor = qm.FileExtractor()

        cust_file = EXPORT_DIR / "QBD_Customers.xlsx"
        if not cust_file.exists():
            raise FileNotFoundError("QBD_Customers.xlsx not found — upload files or generate sample data")

        accounts_df = extractor.extract_accounts(str(EXPORT_DIR / "QBD_ChartOfAccounts.xlsx"))
        customers, jobs = extractor.extract_customers(str(cust_file))
        vendors_df = extractor.extract_vendors(str(EXPORT_DIR / "QBD_Vendors.xlsx"))

        items_file = EXPORT_DIR / "QBD_Items.xlsx"
        items_df = extractor.extract_items(str(items_file)) if items_file.exists() else None

        invoices_file = EXPORT_DIR / "QBD_OpenInvoices.xlsx"
        invoices_df = extractor.extract_invoices(str(invoices_file)) if invoices_file.exists() else None

        bills_file = EXPORT_DIR / "QBD_OpenBills.xlsx"
        bills_df = extractor.extract_bills(str(bills_file)) if bills_file.exists() else None

        # ── Phase 2: Transform ──
        _emit("phase", "Phase 2: Transforming QBD → QBO shape...", phase="transform")
        migration_status["phase"] = "transform"
        transformer = qm.Transformer()
        terms = transformer.collect_terms(customers)

        # ── Phase 3: Load ──
        _emit("phase", f"Phase 3: Loading to QBO{' (dry run)' if dry_run else ''}...", phase="load")
        migration_status["phase"] = "load"

        loader = qm.QBOLoader(qbo, dry_run=dry_run)
        account_ids = loader.create_accounts(accounts_df, transformer)
        term_ids = loader.create_terms(terms)
        customer_ids = loader.create_customers(customers, transformer, term_ids)
        loader.create_jobs(jobs, transformer, customer_ids)
        vendor_ids = loader.create_vendors(vendors_df, transformer)

        # Find a reasonable default income account for item creation
        income_id = next(
            (aid for name, aid in account_ids.items() if "income" in name.lower() or "sales" in name.lower()),
            next(iter(account_ids.values()), ""),
        )
        expense_id = next(
            (aid for name, aid in account_ids.items() if "expense" in name.lower() or "material" in name.lower()),
            "",
        )

        item_ids = {}
        if items_df is not None:
            item_ids = loader.create_items(items_df, transformer, income_id)

        if invoices_df is not None:
            loader.create_invoices(invoices_df, customer_ids, item_ids)

        if bills_df is not None:
            loader.create_bills(bills_df, vendor_ids, expense_id)

        # ── Phase 4: Validate ──
        if not dry_run:
            _emit("phase", "Phase 4: Validating against QBO...", phase="validate")
            migration_status["phase"] = "validate"
            validator = qm.Validator(qbo)
            validation = validator.validate(loader.summary)
        else:
            validation = {}

        summary = loader.get_summary()
        summary["warnings"] = transformer.warnings
        summary["validation"] = validation

        migration_status["summary"] = summary
        migration_status["phase"] = "complete"
        _emit("complete", "Migration complete.", summary=summary, warnings=transformer.warnings)

    except Exception as e:
        log.exception("Migration failed")
        migration_status["error"] = str(e)
        migration_status["phase"] = "error"
        _emit("error", f"Migration failed: {e}")

    finally:
        migration_status["running"] = False
        qm.log.removeHandler(handler)


# ─── Entry point ──────────────────────────────────────────────────

log = logging.getLogger("qbd-qbo")
logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(levelname)-7s  %(message)s")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    banner = "=" * 60
    print(f"\n{banner}")
    print("  QBD → QuickBooks Online Migration Connector")
    print(f"  Mode:  {'DEMO (no Intuit app configured)' if DEMO_MODE else ENVIRONMENT.upper()}")
    print(f"  Open:  http://localhost:{port}")
    print(f"{banner}\n")
    app.run(host="0.0.0.0", port=port, debug=True, threaded=True)

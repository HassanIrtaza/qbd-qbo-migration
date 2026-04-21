"""
QuickBooks Online REST API v3 client.

Handles OAuth 2.0 authorization code flow, automatic token refresh,
and rate-limit (429) retry with exponential backoff.

Docs:
  https://developer.intuit.com/app/developer/qbo/docs/api/accounting/all-entities/account
  https://developer.intuit.com/app/developer/qbo/docs/develop/authentication-and-authorization/oauth-2.0
"""

from __future__ import annotations

import base64
import json
import logging
import time
import urllib.parse
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import requests

log = logging.getLogger("qbo")

AUTH_URL = "https://appcenter.intuit.com/connect/oauth2"
TOKEN_URL = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer"
REVOKE_URL = "https://developer.api.intuit.com/v2/oauth2/tokens/revoke"

API_BASE_SANDBOX = "https://sandbox-quickbooks.api.intuit.com/v3/company"
API_BASE_PROD = "https://quickbooks.api.intuit.com/v3/company"


@dataclass
class QBOTokens:
    access_token: str = ""
    refresh_token: str = ""
    realm_id: str = ""
    expires_at: float = 0.0

    def is_expired(self) -> bool:
        # refresh 60s before actual expiry
        return time.time() > (self.expires_at - 60)

    def to_dict(self) -> dict:
        return {
            "access_token": self.access_token,
            "refresh_token": self.refresh_token,
            "realm_id": self.realm_id,
            "expires_at": self.expires_at,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "QBOTokens":
        return cls(**d)


class QBOClient:
    """Thin wrapper over the QBO REST v3 API."""

    def __init__(
        self,
        client_id: str,
        client_secret: str,
        redirect_uri: str,
        environment: str = "sandbox",
        token_store: Path | None = None,
    ):
        self.client_id = client_id
        self.client_secret = client_secret
        self.redirect_uri = redirect_uri
        self.environment = environment
        self.token_store = token_store or Path("tokens.json")
        self.tokens = QBOTokens()
        self._load_tokens()

    # ── OAuth ─────────────────────────────────────────────────────

    def authorize_url(self, state: str, scopes: str = "com.intuit.quickbooks.accounting") -> str:
        params = {
            "client_id": self.client_id,
            "response_type": "code",
            "scope": scopes,
            "redirect_uri": self.redirect_uri,
            "state": state,
        }
        return f"{AUTH_URL}?{urllib.parse.urlencode(params)}"

    def exchange_code(self, code: str, realm_id: str) -> QBOTokens:
        body = {
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": self.redirect_uri,
        }
        resp = self._token_request(body)
        self.tokens = QBOTokens(
            access_token=resp["access_token"],
            refresh_token=resp["refresh_token"],
            realm_id=realm_id,
            expires_at=time.time() + int(resp["expires_in"]),
        )
        self._save_tokens()
        return self.tokens

    def refresh(self) -> QBOTokens:
        if not self.tokens.refresh_token:
            raise RuntimeError("No refresh token — reauthorize required")
        body = {
            "grant_type": "refresh_token",
            "refresh_token": self.tokens.refresh_token,
        }
        resp = self._token_request(body)
        self.tokens.access_token = resp["access_token"]
        self.tokens.refresh_token = resp.get("refresh_token", self.tokens.refresh_token)
        self.tokens.expires_at = time.time() + int(resp["expires_in"])
        self._save_tokens()
        return self.tokens

    def disconnect(self) -> None:
        if not self.tokens.refresh_token:
            return
        auth = self._basic_auth()
        try:
            requests.post(
                REVOKE_URL,
                headers={"Authorization": auth, "Content-Type": "application/json"},
                json={"token": self.tokens.refresh_token},
                timeout=10,
            )
        except Exception as e:
            log.warning("Revoke failed: %s", e)
        self.tokens = QBOTokens()
        self._save_tokens()

    def is_connected(self) -> bool:
        return bool(self.tokens.access_token and self.tokens.realm_id)

    # ── API ───────────────────────────────────────────────────────

    def _api_base(self) -> str:
        return API_BASE_SANDBOX if self.environment == "sandbox" else API_BASE_PROD

    def request(self, method: str, path: str, **kwargs) -> dict:
        """Send an authenticated request. Handles 401 refresh and 429 backoff."""
        if self.tokens.is_expired():
            self.refresh()

        url = f"{self._api_base()}/{self.tokens.realm_id}/{path.lstrip('/')}"
        headers = kwargs.pop("headers", {})
        headers["Authorization"] = f"Bearer {self.tokens.access_token}"
        headers.setdefault("Accept", "application/json")
        if method.upper() in ("POST", "PUT") and "json" in kwargs:
            headers["Content-Type"] = "application/json"

        for attempt in range(4):
            resp = requests.request(method, url, headers=headers, timeout=30, **kwargs)

            if resp.status_code == 401 and attempt == 0:
                log.info("401 — refreshing token and retrying")
                self.refresh()
                headers["Authorization"] = f"Bearer {self.tokens.access_token}"
                continue

            if resp.status_code == 429:
                wait = 2 ** attempt
                log.warning("429 rate limit — sleeping %ss", wait)
                time.sleep(wait)
                continue

            if resp.status_code >= 400:
                raise QBOError(resp.status_code, resp.text)

            return resp.json() if resp.text else {}

        raise QBOError(0, "Exhausted retries")

    def query(self, sql: str) -> list[dict]:
        """Run a QBO SQL-like query. Returns list of entities."""
        encoded = urllib.parse.quote(sql)
        resp = self.request("GET", f"query?query={encoded}")
        qr = resp.get("QueryResponse", {})
        # first non-meta key holds the list
        for k, v in qr.items():
            if isinstance(v, list):
                return v
        return []

    def create(self, entity: str, body: dict) -> dict:
        return self.request("POST", entity.lower(), json=body)

    def count(self, entity: str) -> int:
        sql = f"SELECT COUNT(*) FROM {entity}"
        encoded = urllib.parse.quote(sql)
        resp = self.request("GET", f"query?query={encoded}")
        return int(resp.get("QueryResponse", {}).get("totalCount", 0))

    def company_info(self) -> dict:
        resp = self.request("GET", f"companyinfo/{self.tokens.realm_id}")
        return resp.get("CompanyInfo", {})

    # ── Internal ──────────────────────────────────────────────────

    def _token_request(self, body: dict) -> dict:
        resp = requests.post(
            TOKEN_URL,
            headers={
                "Authorization": self._basic_auth(),
                "Content-Type": "application/x-www-form-urlencoded",
                "Accept": "application/json",
            },
            data=body,
            timeout=15,
        )
        if resp.status_code >= 400:
            raise QBOError(resp.status_code, resp.text)
        return resp.json()

    def _basic_auth(self) -> str:
        raw = f"{self.client_id}:{self.client_secret}".encode()
        return "Basic " + base64.b64encode(raw).decode()

    def _load_tokens(self) -> None:
        if self.token_store.exists():
            try:
                self.tokens = QBOTokens.from_dict(json.loads(self.token_store.read_text()))
            except Exception as e:
                log.warning("Failed to load tokens: %s", e)

    def _save_tokens(self) -> None:
        try:
            self.token_store.write_text(json.dumps(self.tokens.to_dict(), indent=2))
        except Exception as e:
            log.warning("Failed to save tokens: %s", e)


class QBOError(Exception):
    def __init__(self, status: int, body: str):
        super().__init__(f"QBO API error {status}: {body[:500]}")
        self.status = status
        self.body = body

from __future__ import annotations

import os
from functools import lru_cache
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

PROJECT_ROOT = Path(__file__).resolve().parent
MSAL_CACHE_PATH = PROJECT_ROOT / "temp" / "msal_cache.json"
REGISTRY_LOCAL_PATH = PROJECT_ROOT / "temp" / "registry.xlsx"
ALERT_STATE_PATH = PROJECT_ROOT / "temp" / "alert_state.json"
DEFAULT_TIMEZONE = "Europe/Warsaw"
DEFAULT_GRAPH_SCOPE = "https://graph.microsoft.com/.default"
DEFAULT_GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
DEFAULT_SCAN_INTERVAL_MINUTES = 30


def _require(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise ValueError(f"Missing required environment variable: {name}")
    return value


def _optional(name: str) -> str | None:
    value = os.getenv(name, "").strip()
    return value or None


def _parse_bool(name: str, *, default: bool) -> bool:
    raw = os.getenv(name)
    if raw is None or not raw.strip():
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


def _parse_email_list(raw: str) -> list[str]:
    return [address.strip().lower() for address in raw.split(",") if address.strip()]


@lru_cache
def AZURE_TENANT_ID() -> str:
    return _require("AZURE_TENANT_ID")


@lru_cache
def AZURE_CLIENT_ID() -> str:
    return _require("AZURE_CLIENT_ID")


@lru_cache
def GRAPH_SCOPE() -> str:
    return os.getenv("GRAPH_SCOPE", DEFAULT_GRAPH_SCOPE).strip()


@lru_cache
def GRAPH_BASE_URL() -> str:
    return os.getenv("GRAPH_BASE_URL", DEFAULT_GRAPH_BASE_URL).strip().rstrip("/")


@lru_cache
def MAILBOX_USER() -> str:
    return _require("MAILBOX_USER")


@lru_cache
def CERT_PRIVATE_KEY_PATH() -> Path:
    return Path(_require("CERT_PRIVATE_KEY_PATH")).expanduser()


@lru_cache
def CERT_PUBLIC_PATH() -> Path | None:
    raw = _optional("CERT_PUBLIC_PATH")
    return Path(raw).expanduser() if raw else None


@lru_cache
def CERT_THUMBPRINT() -> str | None:
    return _optional("CERT_THUMBPRINT")


@lru_cache
def TIMEZONE() -> str:
    return os.getenv("TIMEZONE", DEFAULT_TIMEZONE).strip()


@lru_cache
def ONEDRIVE_REGISTRY_PATH() -> str:
    return _require("ONEDRIVE_REGISTRY_PATH")


@lru_cache
def DUPLICATE_ALERT_RECIPIENTS() -> list[str]:
    raw = _require("DUPLICATE_ALERT_RECIPIENTS")
    return [address.strip() for address in raw.split(",") if address.strip()]


@lru_cache
def VALIDATE_PERIOD() -> bool:
    return _parse_bool("VALIDATE_PERIOD", default=False)


@lru_cache
def SCAN_ALL_SENDERS() -> bool:
    return _parse_bool("SCAN_ALL_SENDERS", default=True)


@lru_cache
def ALLOWED_SENDERS() -> list[str]:
    raw = os.getenv("ALLOWED_SENDERS", "")
    return _parse_email_list(raw)


@lru_cache
def scan_allowed_senders() -> frozenset[str]:
    """Return allowed sender addresses when sender filtering is enabled."""
    if SCAN_ALL_SENDERS():
        return frozenset()
    allowed = ALLOWED_SENDERS()
    if not allowed:
        raise ValueError(
            "ALLOWED_SENDERS must contain at least one address when "
            "SCAN_ALL_SENDERS=false"
        )
    return frozenset(allowed)


@lru_cache
def SCAN_INTERVAL_MINUTES() -> int:
    raw = os.getenv("SCAN_INTERVAL_MINUTES", str(DEFAULT_SCAN_INTERVAL_MINUTES)).strip()
    try:
        value = int(raw)
    except ValueError as exc:
        raise ValueError(
            f"SCAN_INTERVAL_MINUTES must be an integer, got: {raw!r}"
        ) from exc
    if value <= 0:
        raise ValueError("SCAN_INTERVAL_MINUTES must be greater than zero")
    return value

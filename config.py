from __future__ import annotations

import os
from functools import lru_cache
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

PROJECT_ROOT = Path(__file__).resolve().parent
MSAL_CACHE_PATH = PROJECT_ROOT / "temp" / "msal_cache.json"
DEFAULT_TIMEZONE = "Europe/Warsaw"
DEFAULT_GRAPH_SCOPE = "https://graph.microsoft.com/.default"
DEFAULT_GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


def _require(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise ValueError(f"Missing required environment variable: {name}")
    return value


def _optional(name: str) -> str | None:
    value = os.getenv(name, "").strip()
    return value or None


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

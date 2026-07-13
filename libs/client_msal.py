from __future__ import annotations

from pathlib import Path

import msal
from msal import SerializableTokenCache

import config


class GraphClient:
    """MSAL client for Microsoft Graph with certificate authentication."""

    def __init__(self) -> None:
        self._tenant_id = config.AZURE_TENANT_ID()
        self._client_id = config.AZURE_CLIENT_ID()
        self._scope = [config.GRAPH_SCOPE()]
        self._cache = SerializableTokenCache()
        self._load_cache()
        self._app = self._build_app()
        self._token: str | None = None

    def _load_cache(self) -> None:
        cache_path = config.MSAL_CACHE_PATH
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        if cache_path.exists():
            self._cache.deserialize(cache_path.read_text(encoding="utf-8"))

    def _save_cache(self) -> None:
        if self._cache.has_state_changed:
            config.MSAL_CACHE_PATH.write_text(
                self._cache.serialize(), encoding="utf-8"
            )

    def _read_pem(self, path: Path) -> str:
        if not path.exists():
            raise FileNotFoundError(f"Certificate file not found: {path}")
        return path.read_text(encoding="utf-8")

    def _build_app(self) -> msal.ConfidentialClientApplication:
        private_key_path = config.CERT_PRIVATE_KEY_PATH()
        private_key = self._read_pem(private_key_path)

        credential: dict[str, str] = {"private_key": private_key}

        public_path = config.CERT_PUBLIC_PATH()
        thumbprint = config.CERT_THUMBPRINT()

        if public_path is not None:
            credential["public_certificate"] = self._read_pem(public_path)
        if thumbprint:
            credential["thumbprint"] = thumbprint
        elif public_path is None:
            raise ValueError(
                "CERT_THUMBPRINT is required when the public certificate file "
                "is not available locally. Copy the thumbprint from Azure Portal "
                "(App registration → Certificates & secrets)."
            )

        return msal.ConfidentialClientApplication(
            client_id=self._client_id,
            authority=f"https://login.microsoftonline.com/{self._tenant_id}",
            client_credential=credential,
            token_cache=self._cache,
        )

    def acquire_token(self, *, force_refresh: bool = False) -> str:
        if self._token and not force_refresh:
            return self._token

        result = self._app.acquire_token_silent(self._scope, account=None)
        if not result or force_refresh:
            result = self._app.acquire_token_for_client(scopes=self._scope)

        self._save_cache()

        if not result or "access_token" not in result:
            error = (result or {}).get("error_description", "Unknown MSAL error")
            raise RuntimeError(f"Failed to acquire token: {error}")

        self._token = result["access_token"]
        return self._token

    def request(
        self,
        method: str,
        url: str,
        *,
        params: dict | None = None,
        json: dict | None = None,
        content: bytes | None = None,
        headers: dict | None = None,
        timeout: float = 60.0,
    ):
        import httpx

        token = self.acquire_token()
        request_headers = {
            "Authorization": f"Bearer {token}",
            **(headers or {}),
        }

        with httpx.Client(timeout=timeout) as client:
            response = client.request(
                method,
                url,
                params=params,
                json=json,
                content=content,
                headers=request_headers,
            )

            if response.status_code == 401:
                token = self.acquire_token(force_refresh=True)
                request_headers["Authorization"] = f"Bearer {token}"
                response = client.request(
                    method,
                    url,
                    params=params,
                    json=json,
                    content=content,
                    headers=request_headers,
                )

            return response

    @property
    def mailbox_user(self) -> str:
        return config.MAILBOX_USER()

    def user_url(self, path: str) -> str:
        base = f"{config.GRAPH_BASE_URL}/users/{self.mailbox_user}"
        if path.startswith("/"):
            return f"{base}{path}"
        return f"{base}/{path}"

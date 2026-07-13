from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import datetime, timedelta
from email.utils import parseaddr
from pathlib import Path
from urllib.parse import quote
from zoneinfo import ZoneInfo

import httpx
import msal
from msal import SerializableTokenCache

import config

logger = logging.getLogger(__name__)

CHUNK_SIZE = 3 * 1024 * 1024
SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024


@dataclass(frozen=True)
class MailAttachment:
    id: str
    name: str
    content_type: str | None = None
    size: int | None = None


@dataclass(frozen=True)
class MailMessage:
    id: str
    received_at: datetime
    sender: str
    subject: str
    attachments: list[MailAttachment]


class GraphClient:
    """Microsoft Graph client: auth, mailbox, OneDrive, outbound mail."""

    def __init__(self) -> None:
        self._tenant_id = config.AZURE_TENANT_ID()
        self._client_id = config.AZURE_CLIENT_ID()
        self._scope = [config.GRAPH_SCOPE()]
        self._timezone = ZoneInfo(config.TIMEZONE())
        self._cache = SerializableTokenCache()
        self._load_cache()
        self._app = self._build_app()
        self._token: str | None = None

    @property
    def mail_scan_user(self) -> str:
        return config.MAIL_SCAN_USER()

    @property
    def onedrive_user(self) -> str:
        return config.ONEDRIVE_USER()

    @property
    def mail_send_user(self) -> str:
        return config.MAIL_SEND_USER()

    def user_url(self, user: str, path: str) -> str:
        base = f"{config.GRAPH_BASE_URL()}/users/{user}"
        if path.startswith("/"):
            return f"{base}{path}"
        return f"{base}/{path}"

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
    ) -> httpx.Response:
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

    def fetch_inbox_messages_today(self) -> list[MailMessage]:
        """Return inbox messages received today (local timezone) that have attachments."""
        start_utc, end_utc = self._today_utc_range()
        url = self.user_url(self.mail_scan_user, "/mailFolders/inbox/messages")
        params = {
            "$filter": (
                f"receivedDateTime ge {start_utc} and receivedDateTime lt {end_utc} "
                "and hasAttachments eq true"
            ),
            "$select": "id,receivedDateTime,from,subject,hasAttachments",
            "$expand": "attachments($select=id,name,contentType,size)",
            "$orderby": "receivedDateTime asc",
            "$top": "50",
        }

        messages: list[MailMessage] = []
        next_url: str | None = url
        next_params: dict | None = params

        while next_url:
            response = self.request("GET", next_url, params=next_params)
            if response.status_code != 200:
                raise RuntimeError(
                    f"Failed to fetch inbox messages: "
                    f"{response.status_code} {response.text}"
                )

            payload = response.json()
            for item in payload.get("value", []):
                parsed = self._parse_message(item)
                if parsed.attachments:
                    messages.append(parsed)

            next_url = payload.get("@odata.nextLink")
            next_params = None

        logger.info("Fetched %d inbox message(s) with attachments for today", len(messages))
        return messages

    def download_drive_file(self, remote_path: str, local_path: Path) -> bool:
        """Download a file from OneDrive. Returns False when the remote file does not exist."""
        local_path.parent.mkdir(parents=True, exist_ok=True)
        url = f"{self._drive_item_path(remote_path)}/content"
        response = self.request("GET", url)

        if response.status_code == 404:
            logger.info("OneDrive file not found: %s", remote_path)
            return False
        if response.status_code != 200:
            raise RuntimeError(
                f"Failed to download OneDrive file: "
                f"{response.status_code} {response.text}"
            )

        local_path.write_bytes(response.content)
        logger.info("Downloaded OneDrive file to %s", local_path)
        return True

    def upload_drive_file(self, content: bytes, remote_path: str) -> str:
        """Upload bytes to OneDrive and return the item web URL."""
        if len(content) <= SIMPLE_UPLOAD_LIMIT:
            return self._simple_upload(content, remote_path)
        return self._session_upload(content, remote_path)

    def upload_drive_file_from_path(self, local_path: Path, remote_path: str) -> str:
        return self.upload_drive_file(local_path.read_bytes(), remote_path)

    def send_mail(
        self,
        recipients: list[str],
        subject: str,
        body: str,
        *,
        content_type: str = "Text",
        save_to_sent_items: bool = True,
    ) -> None:
        if not recipients:
            raise ValueError("At least one recipient is required")

        payload = {
            "message": {
                "subject": subject,
                "body": {"contentType": content_type, "content": body},
                "toRecipients": [
                    {"emailAddress": {"address": address.strip()}}
                    for address in recipients
                    if address.strip()
                ],
            },
            "saveToSentItems": save_to_sent_items,
        }
        response = self.request(
            "POST",
            self.user_url(self.mail_send_user, "/sendMail"),
            json=payload,
        )
        if response.status_code != 202:
            raise RuntimeError(
                f"Failed to send mail: {response.status_code} {response.text}"
            )
        logger.info("Sent mail to %s", ", ".join(recipients))

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
        private_key = self._read_pem(config.CERT_PRIVATE_KEY_PATH())
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
                "(App registration -> Certificates & secrets)."
            )

        return msal.ConfidentialClientApplication(
            client_id=self._client_id,
            authority=f"https://login.microsoftonline.com/{self._tenant_id}",
            client_credential=credential,
            token_cache=self._cache,
        )

    def _today_utc_range(self) -> tuple[str, str]:
        now_local = datetime.now(self._timezone)
        start_local = now_local.replace(hour=0, minute=0, second=0, microsecond=0)
        end_local = start_local + timedelta(days=1)
        start_utc = start_local.astimezone(ZoneInfo("UTC")).strftime("%Y-%m-%dT%H:%M:%SZ")
        end_utc = end_local.astimezone(ZoneInfo("UTC")).strftime("%Y-%m-%dT%H:%M:%SZ")
        return start_utc, end_utc

    def _parse_message(self, item: dict) -> MailMessage:
        received_raw = item.get("receivedDateTime", "")
        received_at = datetime.fromisoformat(received_raw.replace("Z", "+00:00")).astimezone(
            self._timezone
        )

        sender_info = item.get("from", {}).get("emailAddress", {})
        sender = sender_info.get("address") or sender_info.get("name") or ""
        sender = parseaddr(sender)[1] or sender

        attachments: list[MailAttachment] = []
        for attachment in item.get("attachments", []):
            if attachment.get("@odata.type") != "#microsoft.graph.fileAttachment":
                continue
            name = attachment.get("name")
            attachment_id = attachment.get("id")
            if not name or not attachment_id:
                continue
            attachments.append(
                MailAttachment(
                    id=attachment_id,
                    name=name,
                    content_type=attachment.get("contentType"),
                    size=attachment.get("size"),
                )
            )

        return MailMessage(
            id=item.get("id", ""),
            received_at=received_at,
            sender=sender,
            subject=item.get("subject", "") or "",
            attachments=attachments,
        )

    def _drive_item_path(self, remote_path: str) -> str:
        normalized = remote_path.strip("/")
        encoded = "/".join(quote(part, safe="") for part in normalized.split("/"))
        return f"{self.user_url(self.onedrive_user, '/drive/root:')}:{encoded}:"

    def _simple_upload(self, content: bytes, remote_path: str) -> str:
        url = f"{self._drive_item_path(remote_path)}/content"
        response = self.request(
            "PUT",
            url,
            content=content,
            headers={"Content-Type": "application/octet-stream"},
            timeout=120.0,
        )
        if response.status_code not in (200, 201):
            raise RuntimeError(
                f"Failed to upload OneDrive file: "
                f"{response.status_code} {response.text}"
            )
        web_url = response.json().get("webUrl", "")
        logger.info("Uploaded %s (%d bytes)", remote_path, len(content))
        return web_url

    def _session_upload(self, content: bytes, remote_path: str) -> str:
        session_url = f"{self._drive_item_path(remote_path)}:/createUploadSession"
        session_response = self.request(
            "POST",
            session_url,
            json={
                "item": {
                    "@microsoft.graph.conflictBehavior": "replace",
                    "name": Path(remote_path).name,
                }
            },
        )
        if session_response.status_code not in (200, 201):
            raise RuntimeError(
                f"Failed to create OneDrive upload session: "
                f"{session_response.status_code} {session_response.text}"
            )

        upload_url = session_response.json()["uploadUrl"]
        total = len(content)
        web_url = ""

        with httpx.Client(timeout=120.0) as http_client:
            for start in range(0, total, CHUNK_SIZE):
                end = min(start + CHUNK_SIZE, total) - 1
                chunk = content[start : end + 1]
                headers = {
                    "Content-Length": str(len(chunk)),
                    "Content-Range": f"bytes {start}-{end}/{total}",
                }
                response = http_client.put(upload_url, content=chunk, headers=headers)
                if response.status_code in (200, 201):
                    web_url = response.json().get("webUrl", web_url)
                elif response.status_code not in (202,):
                    raise RuntimeError(
                        f"OneDrive chunk upload failed: "
                        f"{response.status_code} {response.text}"
                    )

        if not web_url:
            meta_response = self.request("GET", self._drive_item_path(remote_path))
            if meta_response.status_code == 200:
                web_url = meta_response.json().get("webUrl", "")

        logger.info("Uploaded %s via session (%d bytes)", remote_path, total)
        return web_url

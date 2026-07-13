from __future__ import annotations

from pathlib import Path
from urllib.parse import quote

from libs.msal_client import GraphClient
from logs.app import get_app_logger
from logs.err import get_err_logger

logger = get_app_logger()
err_logger = get_err_logger()

CHUNK_SIZE = 3 * 1024 * 1024  # 3 MB
SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024  # 4 MB


class OneDriveUploader:
    def __init__(self, client: GraphClient) -> None:
        self._client = client
        self._attachments_folder = None
        self._registry_path = None

    @property
    def attachments_folder(self) -> str:
        if self._attachments_folder is None:
            import config

            self._attachments_folder = config.ONEDRIVE_ATTACHMENTS_FOLDER()
        return self._attachments_folder

    @property
    def registry_path(self) -> str:
        if self._registry_path is None:
            import config

            self._registry_path = config.ONEDRIVE_REGISTRY_PATH()
        return self._registry_path

    def _item_path(self, remote_path: str) -> str:
        normalized = remote_path.strip("/")
        encoded = "/".join(quote(part, safe="") for part in normalized.split("/"))
        return f"{self._client.user_url('/drive/root:')}:{encoded}:"

    def download_registry(self, local_path: Path) -> bool:
        """Download registry XLSX from OneDrive. Returns False if file does not exist."""
        local_path.parent.mkdir(parents=True, exist_ok=True)
        url = f"{self._item_path(self.registry_path)}/content"
        response = self._client.request("GET", url)

        if response.status_code == 404:
            logger.info("Registry file not found on OneDrive; will create a new one")
            return False
        if response.status_code != 200:
            raise RuntimeError(
                f"Failed to download registry: {response.status_code} {response.text}"
            )

        local_path.write_bytes(response.content)
        logger.info("Downloaded registry to %s", local_path)
        return True

    def upload_registry(self, local_path: Path) -> None:
        self._upload_bytes(local_path.read_bytes(), self.registry_path)
        logger.info("Uploaded registry to OneDrive: %s", self.registry_path)

    def upload_file(
        self, local_path: Path, remote_folder: str, filename: str | None = None
    ) -> str:
        name = filename or local_path.name
        remote_path = f"{remote_folder.strip('/')}/{name}"
        content = local_path.read_bytes()
        return self._upload_bytes(content, remote_path)

    def _upload_bytes(self, content: bytes, remote_path: str) -> str:
        if len(content) <= SIMPLE_UPLOAD_LIMIT:
            return self._simple_upload(content, remote_path)
        return self._session_upload(content, remote_path)

    def _simple_upload(self, content: bytes, remote_path: str) -> str:
        url = f"{self._item_path(remote_path)}/content"
        response = self._client.request(
            "PUT",
            url,
            content=content,
            headers={"Content-Type": "application/octet-stream"},
            timeout=120.0,
        )
        if response.status_code not in (200, 201):
            raise RuntimeError(
                f"Failed to upload file: {response.status_code} {response.text}"
            )
        payload = response.json()
        web_url = payload.get("webUrl", "")
        logger.info("Uploaded %s (%d bytes)", remote_path, len(content))
        return web_url

    def _session_upload(self, content: bytes, remote_path: str) -> str:
        session_url = f"{self._item_path(remote_path)}:/createUploadSession"
        session_response = self._client.request(
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
                f"Failed to create upload session: "
                f"{session_response.status_code} {session_response.text}"
            )

        upload_url = session_response.json()["uploadUrl"]
        total = len(content)
        web_url = ""

        import httpx

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
                        f"Chunk upload failed: {response.status_code} {response.text}"
                    )

        if not web_url:
            meta_response = self._client.request("GET", self._item_path(remote_path))
            if meta_response.status_code == 200:
                web_url = meta_response.json().get("webUrl", "")

        logger.info("Uploaded %s via session (%d bytes)", remote_path, total)
        return web_url

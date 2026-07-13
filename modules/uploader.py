from __future__ import annotations

from pathlib import Path

import config
from libs.msal_client import GraphClient
from logs.app import get_app_logger

logger = get_app_logger()


class RegistryUploader:
    """Upload the local registry workbook to OneDrive."""

    def __init__(self, client: GraphClient) -> None:
        self._client = client

    def upload(self, local_path: Path) -> None:
        remote_path = config.ONEDRIVE_REGISTRY_PATH()
        self._client.upload_drive_file_from_path(local_path, remote_path)
        logger.info("Uploaded registry to OneDrive: %s", remote_path)

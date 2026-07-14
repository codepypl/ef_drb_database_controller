from __future__ import annotations

from pathlib import PurePosixPath

import config
from libs.msal_client import GraphClient
from logs.app import get_app_logger
from modules.registry import Registry
from modules.uploader import RegistryUploader

logger = get_app_logger()


class OneDriveStructure:
    """Ensure registry folders and file exist on ONEDRIVE_USER's OneDrive."""

    def __init__(self, client: GraphClient) -> None:
        self._client = client
        self._remote_path = config.ONEDRIVE_REGISTRY_PATH().strip("/")

    def ensure_registry_structure(self) -> None:
        folder_path, _filename = self._split_remote_path()

        if folder_path:
            logger.info("Ensuring OneDrive folder exists: %s", folder_path)
            self._client.ensure_drive_folders(folder_path)

        if self._client.drive_item_exists(self._remote_path):
            logger.info("Registry file exists on OneDrive: %s", self._remote_path)
            self._client.download_drive_file(self._remote_path, config.REGISTRY_LOCAL_PATH)
            return

        logger.info("Creating registry file on OneDrive: %s", self._remote_path)
        registry = Registry()
        registry.save()
        RegistryUploader(self._client).upload(registry.path)

    def _split_remote_path(self) -> tuple[str, str]:
        path = PurePosixPath(self._remote_path)
        folder = "" if path.parent == PurePosixPath(".") else str(path.parent)
        return folder, path.name

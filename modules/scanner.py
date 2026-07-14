from __future__ import annotations

import re
from datetime import date, datetime
from zoneinfo import ZoneInfo

import config
from libs.msal_client import GraphClient, MailAttachment, MailMessage
from logs.app import get_app_logger
from logs.err import get_err_logger
from modules.mailer import DuplicateMailer
from modules.registry import Registry
from modules.structure import OneDriveStructure
from modules.uploader import RegistryUploader

logger = get_app_logger()
err_logger = get_err_logger()

ATTACHMENT_PATTERN = re.compile(
    r"^[A-Z][A-Z0-9_]*_(?:(?:0[1-9]|1[0-2])|Q[1-4])_\d{4}\.(?:xlsx|csv)$",
    re.IGNORECASE,
)
PERIOD_PATTERN = re.compile(
    r"^[A-Z][A-Z0-9_]*_((?:0[1-9]|1[0-2])|Q[1-4])_(\d{4})\.(?:xlsx|csv)$",
    re.IGNORECASE,
)


class MailboxScanner:
    def __init__(self, client: GraphClient | None = None) -> None:
        self._client = client or GraphClient()
        self._timezone = ZoneInfo(config.TIMEZONE())
        self._validate_period = config.VALIDATE_PERIOD()
        self._scan_all_senders = config.SCAN_ALL_SENDERS()
        self._allowed_senders = config.scan_allowed_senders()

    def run_once(self) -> None:
        scanned_at = datetime.now(self._timezone)
        today = scanned_at.date()
        registry = Registry()

        OneDriveStructure(self._client).ensure_registry_structure()

        registry.load()
        mailer = DuplicateMailer(self._client)
        mailer.prune_old_state()

        changed = False
        messages = self._client.fetch_inbox_messages_today()
        logger.info("Processing %d message(s) from today's inbox", len(messages))

        for message in messages:
            if not self._is_allowed_sender(message.sender):
                logger.debug("Skipping message %s from %s", message.id, message.sender)
                continue

            for attachment in message.attachments:
                try:
                    if self._process_attachment(
                        attachment=attachment,
                        message=message,
                        registry=registry,
                        mailer=mailer,
                        scanned_at=scanned_at,
                        today=today,
                    ):
                        changed = True
                except Exception as exc:
                    err_logger.error(
                        "Failed to process attachment %s from message %s: %s",
                        attachment.name,
                        message.id,
                        exc,
                        exc_info=True,
                    )

        if changed:
            registry.save()
            RegistryUploader(self._client).upload(registry.path)
            logger.info("Registry updated and uploaded")
        else:
            logger.info("No registry changes detected")

    def _process_attachment(
        self,
        *,
        attachment: MailAttachment,
        message: MailMessage,
        registry: Registry,
        mailer: DuplicateMailer,
        scanned_at: datetime,
        today: date,
    ) -> bool:
        filename = attachment.name.strip()
        if not is_matching_attachment(filename):
            return False

        if self._validate_period and not is_expected_period(filename, today):
            logger.warning("Attachment %s failed period validation", filename)
            return False

        existing = registry.find_by_filename(filename)
        if existing is None:
            registry.add_entry(
                Registry.build_entry(message, attachment, scanned_at)
            )
            logger.info("Added registry entry for %s", filename)
            return True

        if existing.scan_date == today:
            logger.debug("Skipping %s; already registered today", filename)
            return False

        mailer.send_duplicate_alert(existing, message, attachment)
        logger.warning("Duplicate attachment detected: %s", filename)
        return False

    def _is_allowed_sender(self, sender: str) -> bool:
        if self._scan_all_senders:
            return True
        return sender.strip().lower() in self._allowed_senders


def is_matching_attachment(filename: str) -> bool:
    return bool(ATTACHMENT_PATTERN.match(filename.strip()))


def is_expected_period(filename: str, today: date) -> bool:
    match = PERIOD_PATTERN.match(filename.strip())
    if not match:
        return False

    period, year_text = match.groups()
    year = int(year_text)
    period_upper = period.upper()

    if period_upper.startswith("Q"):
        expected_period, expected_year = _expected_quarterly_period(today)
        return period_upper == expected_period and year == expected_year

    expected_period, expected_year = _expected_monthly_period(today)
    return period == expected_period and year == expected_year


def _expected_monthly_period(today: date) -> tuple[str, int]:
    if today.month == 1:
        return "12", today.year - 1
    return f"{today.month - 1:02d}", today.year


def _expected_quarterly_period(today: date) -> tuple[str, int]:
    current_quarter = (today.month - 1) // 3 + 1
    if current_quarter == 1:
        return "Q4", today.year - 1
    return f"Q{current_quarter - 1}", today.year

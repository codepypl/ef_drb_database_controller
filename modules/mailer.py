from __future__ import annotations

import json
from datetime import date, datetime
from zoneinfo import ZoneInfo

import config
from libs.msal_client import GraphClient, MailAttachment, MailMessage
from logs.app import get_app_logger
from modules.registry import RegistryEntry

logger = get_app_logger()


class DuplicateMailer:
    def __init__(self, client: GraphClient) -> None:
        self._client = client
        self._timezone = ZoneInfo(config.TIMEZONE())
        self._recipients = config.DUPLICATE_ALERT_RECIPIENTS()
        self._state_path = config.ALERT_STATE_PATH
        self._state = self._load_state()

    def send_duplicate_alert(
        self,
        existing: RegistryEntry,
        message: MailMessage,
        attachment: MailAttachment,
    ) -> None:
        today = datetime.now(self._timezone).date().isoformat()
        filename = attachment.name.strip()
        alerted_today = self._state.setdefault(today, [])

        if filename in alerted_today:
            logger.info("Duplicate alert already sent today for %s", filename)
            return

        subject = f"[DRB] Duplikat załącznika: {filename}"
        body = (
            "Wykryto ponowne wystąpienie załącznika w skrzynce pocztowej.\n\n"
            f"Nazwa pliku: {filename}\n"
            f"Pierwszy wpis w rejestrze: {existing.scan_date.isoformat()} "
            f"{existing.scan_time.strftime('%H:%M:%S')}\n"
            f"Pierwotny nadawca: {existing.sender}\n"
            f"Pierwotny temat: {existing.subject}\n\n"
            f"Nowy mail od: {message.sender}\n"
            f"Nowy temat: {message.subject}\n"
            f"Godzina dostarczenia: {message.received_at.strftime('%H:%M:%S')}\n"
        )

        self._client.send_mail(self._recipients, subject, body)
        alerted_today.append(filename)
        self._save_state()
        logger.info("Sent duplicate alert for %s", filename)

    def _load_state(self) -> dict[str, list[str]]:
        if not self._state_path.exists():
            return {}
        payload = json.loads(self._state_path.read_text(encoding="utf-8"))
        return {
            day: list(filenames)
            for day, filenames in payload.items()
            if isinstance(filenames, list)
        }

    def _save_state(self) -> None:
        self._state_path.parent.mkdir(parents=True, exist_ok=True)
        self._state_path.write_text(
            json.dumps(self._state, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def prune_old_state(self, *, keep_days: int = 30) -> None:
        cutoff = datetime.now(self._timezone).date().toordinal() - keep_days
        kept: dict[str, list[str]] = {}
        for day, filenames in self._state.items():
            try:
                day_date = date.fromisoformat(day)
            except ValueError:
                continue
            if day_date.toordinal() >= cutoff:
                kept[day] = filenames
        if kept != self._state:
            self._state = kept
            self._save_state()

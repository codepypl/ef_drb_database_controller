from __future__ import annotations

import sys

from logs.app import get_app_logger
from logs.err import get_err_logger
from modules.scanner import MailboxScanner


def main() -> int:
    logger = get_app_logger()
    err_logger = get_err_logger()

    try:
        logger.info("Starting mailbox scan")
        MailboxScanner().run_once()
        logger.info("Mailbox scan finished")
        return 0
    except Exception as exc:
        err_logger.error("Mailbox scan failed: %s", exc, exc_info=True)
        return 1


if __name__ == "__main__":
    sys.exit(main())

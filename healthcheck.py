#!/usr/bin/env python3
"""Container health check: open an SMTP session and verify the 220 banner.

Exits 0 if the relay answers with a 2xx greeting, 1 otherwise. Used by the
Dockerfile HEALTHCHECK so an orchestrator can restart a locked-up container.
"""

import os
import smtplib
import sys


def main() -> int:
    port = int(os.environ.get("Smtp_Port") or os.environ.get("Smtp__Port") or 25)
    host = "127.0.0.1"
    try:
        with smtplib.SMTP(host, port, timeout=5) as smtp:
            code, _ = smtp.ehlo()
            if 200 <= code < 400:
                return 0
            print(f"unhealthy: EHLO returned {code}", file=sys.stderr)
            return 1
    except Exception as exc:  # noqa: BLE001
        print(f"unhealthy: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())

#!/usr/bin/env python3
"""Bare-bones SMTP relay that forwards mail through Microsoft 365 / Graph.

Listens as an (internal) open SMTP relay and sends each received message via
the Microsoft Graph `sendMail` endpoint using app-only (client credentials)
authentication. Intended to sit on a private Docker network so other
containers can send mail without dealing with OAuth themselves.
"""

import asyncio
import base64
import logging
import os
import signal
import sys
from email import message_from_bytes
from email.utils import getaddresses
from logging.handlers import RotatingFileHandler

import msal
import requests
from aiosmtpd.controller import Controller

GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

log = logging.getLogger("smtp-relay")
events = logging.getLogger("smtp-relay.events")


def env(*names, default=None, required=False):
    """Return the first environment variable that is set among *names.

    Lets us accept both single- and double-underscore spellings (e.g.
    ``Graph_ClientSecret`` and ``Graph__ClientSecret``) without fuss.
    """
    for name in names:
        value = os.environ.get(name)
        if value not in (None, ""):
            return value
    if required:
        sys.exit(f"Missing required environment variable: {' / '.join(names)}")
    return default


class Config:
    def __init__(self):
        self.tenant_id = env("Graph_TenantId", "Graph__TenantId", required=True)
        self.client_id = env("Graph_ClientId", "Graph__ClientId", required=True)
        self.client_secret = env(
            "Graph_ClientSecret", "Graph__ClientSecret", required=True
        )
        self.send_from = env("SendFrom", required=True)
        self.log_level = env("LogLevel", default="INFO").upper()
        self.log_path = env("LogPath", "LogFile")
        self.smtp_host = env("Smtp_Host", "Smtp__Host", default="0.0.0.0")
        self.smtp_port = int(env("Smtp_Port", "Smtp__Port", default="25"))


class GraphMailer:
    """Acquires app-only tokens (cached by MSAL) and sends mail via Graph."""

    def __init__(self, cfg: Config):
        self.cfg = cfg
        self.app = msal.ConfidentialClientApplication(
            client_id=cfg.client_id,
            client_credential=cfg.client_secret,
            authority=f"https://login.microsoftonline.com/{cfg.tenant_id}",
        )

    def _token(self) -> str:
        result = self.app.acquire_token_for_client(scopes=GRAPH_SCOPE)
        if "access_token" not in result:
            raise RuntimeError(
                "Failed to acquire Graph token: "
                f"{result.get('error')}: {result.get('error_description')}"
            )
        return result["access_token"]

    def send(self, sender: str, recipients: list[str], raw: bytes) -> None:
        parsed = message_from_bytes(raw)
        subject = parsed.get("Subject", "")
        body, content_type, attachments = _extract_body(parsed)

        message = {
            "subject": subject,
            "body": {"contentType": content_type, "content": body},
            "toRecipients": [
                {"emailAddress": {"address": addr}} for addr in recipients
            ],
        }
        if attachments:
            message["attachments"] = attachments

        url = f"{GRAPH_BASE}/users/{sender}/sendMail"
        resp = requests.post(
            url,
            headers={
                "Authorization": f"Bearer {self._token()}",
                "Content-Type": "application/json",
            },
            json={"message": message, "saveToSentItems": False},
            timeout=30,
        )
        if resp.status_code not in (200, 202):
            raise RuntimeError(
                f"Graph sendMail failed ({resp.status_code}): {resp.text}"
            )


def _extract_body(parsed):
    """Return (body, contentType, attachments) from a parsed email message."""
    text_body = ""
    html_body = ""
    attachments = []

    if parsed.is_multipart():
        for part in parsed.walk():
            if part.is_multipart():
                continue
            ctype = part.get_content_type()
            disposition = (part.get("Content-Disposition") or "").lower()
            if "attachment" in disposition or part.get_filename():
                payload = part.get_payload(decode=True) or b""
                attachments.append(
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": part.get_filename() or "attachment",
                        "contentType": ctype,
                        "contentBytes": base64.b64encode(payload).decode("ascii"),
                    }
                )
            elif ctype == "text/plain":
                text_body += _decode_part(part)
            elif ctype == "text/html":
                html_body += _decode_part(part)
    else:
        if parsed.get_content_type() == "text/html":
            html_body = _decode_part(parsed)
        else:
            text_body = _decode_part(parsed)

    if html_body:
        return html_body, "HTML", attachments
    return text_body, "Text", attachments


def _decode_part(part) -> str:
    payload = part.get_payload(decode=True)
    if payload is None:
        return ""
    charset = part.get_content_charset() or "utf-8"
    try:
        return payload.decode(charset, errors="replace")
    except (LookupError, TypeError):
        return payload.decode("utf-8", errors="replace")


class RelayHandler:
    def __init__(self, mailer: GraphMailer, cfg: Config):
        self.mailer = mailer
        self.cfg = cfg

    async def handle_DATA(self, server, session, envelope):
        parsed = message_from_bytes(envelope.content)
        peer_ip = session.peer[0] if session.peer else "unknown"

        recipients = list(envelope.rcpt_tos)
        # Fall back to recipients found in the message headers if the envelope
        # somehow has none.
        if not recipients:
            recipients = [
                addr
                for _, addr in getaddresses(
                    parsed.get_all("To", []) + parsed.get_all("Cc", [])
                )
                if addr
            ]

        cc = [addr for _, addr in getaddresses(parsed.get_all("Cc", [])) if addr]
        subject = parsed.get("Subject", "")
        # Header From / envelope MAIL FROM are logged for visibility, but the
        # message is always *sent* from SendFrom (a Graph constraint).
        from_header = parsed.get("From", "")

        self._log_event(
            peer_ip=peer_ip,
            mail_from=envelope.mail_from or "",
            from_header=from_header,
            recipients=recipients,
            cc=cc,
            subject=subject,
            size=len(envelope.content),
        )

        if not recipients:
            events.warning("REJECTED ip=%s reason=no-recipients", peer_ip)
            return "550 No recipients"

        sender = self.cfg.send_from
        log.info("Relaying message from %s to %s", sender, ", ".join(recipients))
        try:
            await asyncio.to_thread(
                self.mailer.send, sender, recipients, envelope.content
            )
        except Exception as exc:  # noqa: BLE001 - report any failure back to client
            log.error("Failed to relay message: %s", exc)
            events.error("FAILED ip=%s to=%s error=%s", peer_ip, ",".join(recipients), exc)
            return f"451 Relay failed: {exc}"
        events.info("SENT ip=%s to=%s", peer_ip, ",".join(recipients))
        return "250 Message accepted for delivery"

    @staticmethod
    def _log_event(*, peer_ip, mail_from, from_header, recipients, cc, subject, size):
        events.info(
            "RECEIVED ip=%s mail_from=%s from=%r to=%s cc=%s subject=%r size=%d",
            peer_ip,
            mail_from or "-",
            from_header or "-",
            ",".join(recipients) or "-",
            ",".join(cc) or "-",
            subject or "-",
            size,
        )


def setup_logging(cfg: Config):
    level = getattr(logging, cfg.log_level, logging.INFO)
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")

    root = logging.getLogger()
    root.setLevel(level)

    # aiosmtpd's internal logger is very chatty at INFO; keep its protocol
    # noise out of our logs unless the user explicitly asks for DEBUG.
    if level > logging.DEBUG:
        logging.getLogger("mail.log").setLevel(logging.WARNING)

    console = logging.StreamHandler()
    console.setFormatter(fmt)
    root.addHandler(console)

    if cfg.log_path:
        # If a directory is given, write to relay.log inside it.
        path = cfg.log_path
        if os.path.isdir(path) or path.endswith(os.sep):
            path = os.path.join(path, "relay.log")
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        file_handler = RotatingFileHandler(
            path, maxBytes=10 * 1024 * 1024, backupCount=5, encoding="utf-8"
        )
        file_handler.setFormatter(fmt)
        root.addHandler(file_handler)
        log.info("Writing logs to %s", path)


def main():
    cfg = Config()
    setup_logging(cfg)
    log.info(
        "Starting SMTP relay on %s:%s -> Graph sendMail as %s",
        cfg.smtp_host,
        cfg.smtp_port,
        cfg.send_from,
    )

    mailer = GraphMailer(cfg)
    controller = Controller(
        RelayHandler(mailer, cfg),
        hostname=cfg.smtp_host,
        port=cfg.smtp_port,
    )
    controller.start()

    stop = asyncio.Event()
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    for sig in (signal.SIGINT, signal.SIGTERM):
        loop.add_signal_handler(sig, stop.set)
    try:
        loop.run_until_complete(stop.wait())
    finally:
        log.info("Shutting down")
        controller.stop()
        loop.close()


if __name__ == "__main__":
    main()

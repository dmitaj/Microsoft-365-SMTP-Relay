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

    def send(
        self,
        sender: str,
        recipients: list[str],
        raw: bytes,
        from_name: str = "",
    ) -> str:
        """Send the message and return the From address actually used."""
        parsed = message_from_bytes(raw)
        subject = parsed.get("Subject", "")
        body, content_type, attachments = _extract_body(parsed)

        def build(from_addr: str, name: str) -> dict:
            emailAddress = {"address": from_addr}
            if name:
                emailAddress["name"] = name
            msg = {
                "subject": subject,
                "body": {"contentType": content_type, "content": body},
                "from": {"emailAddress": emailAddress},
                "toRecipients": [
                    {"emailAddress": {"address": addr}} for addr in recipients
                ],
            }
            if attachments:
                msg["attachments"] = attachments
            return msg

        fallback = self.cfg.send_from
        distinct = sender.lower() != fallback.lower()

        # 1) Send directly as the From address (user mailboxes and aliases).
        status, text = self._post(sender, build(sender, from_name))
        if status in (200, 202):
            return sender

        # 2) The From isn't a user mailbox (e.g. a Microsoft 365 Group, which
        #    Graph rejects with 404 ErrorInvalidUser): route through SendFrom
        #    but keep the original From. Requires SendFrom to have Send As /
        #    Send on Behalf on that address.
        if distinct and _is_invalid_user(status, text):
            log.warning(
                "Sender %s is not a user mailbox (%s); retrying via %s keeping From",
                sender,
                status,
                fallback,
            )
            status, text = self._post(fallback, build(sender, from_name))
            if status in (200, 202):
                return sender

        # 3) The From still can't be used (non-existent address, or no Send As
        #    grant): send plainly as SendFrom so the mail still goes out.
        if distinct and _is_unusable_sender(status, text):
            log.warning(
                "Cannot send as %s (%s); falling back to From=%s",
                sender,
                status,
                fallback,
            )
            status, text = self._post(fallback, build(fallback, ""))
            if status in (200, 202):
                return fallback

        raise RuntimeError(f"Graph sendMail failed ({status}): {text}")

    def _post(self, mailbox: str, message: dict) -> tuple[int, str]:
        resp = requests.post(
            f"{GRAPH_BASE}/users/{mailbox}/sendMail",
            headers={
                "Authorization": f"Bearer {self._token()}",
                "Content-Type": "application/json",
            },
            json={"message": message, "saveToSentItems": False},
            timeout=30,
        )
        return resp.status_code, resp.text


def _is_invalid_user(status: int, text: str) -> bool:
    """True when Graph rejected the sending mailbox as not a valid user.

    This is what comes back when the From address belongs to something that
    isn't a user mailbox (a Microsoft 365 Group, a contact, a typo, etc.).
    """
    return status == 404 and (
        "ErrorInvalidUser" in text or "ResourceNotFound" in text
    )


def _is_unusable_sender(status: int, text: str) -> bool:
    """True when the From address can't be used as a sender at all.

    Covers both "not a user mailbox" (404 ErrorInvalidUser) and "not allowed to
    send as this address" (403 ErrorSendAsDenied) - e.g. a non-existent From, or
    a group with no Send As grant. In these cases we send plainly as SendFrom.
    """
    return _is_invalid_user(status, text) or (
        status == 403 and "ErrorSendAsDenied" in text
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
        from_header = parsed.get("From", "")

        # Send as the message's From address when present, then the envelope
        # MAIL FROM, and finally fall back to the configured SendFrom mailbox.
        from_pairs = getaddresses(parsed.get_all("From", []))
        from_name, from_addr = from_pairs[0] if from_pairs else ("", "")
        sender = from_addr or envelope.mail_from or self.cfg.send_from

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

        log.info("Relaying message from %s to %s", sender, ", ".join(recipients))
        try:
            sent_from = await asyncio.to_thread(
                self.mailer.send, sender, recipients, envelope.content, from_name
            )
        except Exception as exc:  # noqa: BLE001 - report any failure back to client
            log.error("Failed to relay message: %s", exc)
            events.error("FAILED ip=%s to=%s error=%s", peer_ip, ",".join(recipients), exc)
            return f"451 Relay failed: {exc}"
        events.info(
            "SENT ip=%s from=%s to=%s", peer_ip, sent_from, ",".join(recipients)
        )
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

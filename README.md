# Microsoft 365 SMTP Relay

A bare-bones SMTP relay in a single Docker container. It accepts plain SMTP
from other containers on your private network and forwards each message through
the **Microsoft Graph `sendMail` API** using app-only (client credentials)
authentication.

No web portal, no database, no TLS termination, no user accounts — just an
internal open relay so your apps can send mail through Microsoft 365 without
implementing OAuth themselves.

> ⚠️ This is an **open relay** with no authentication. Only run it on a trusted,
> private Docker network. Do **not** expose it to the public internet.

## How it works

```
your-app ──SMTP──► smtp-relay ──HTTPS / Graph sendMail──► Microsoft 365
```

The relay listens for SMTP connections, takes the envelope recipients and the
message body/attachments, and calls
`POST /v1.0/users/{SendFrom}/sendMail`. All mail is sent from the single
`SendFrom` mailbox.

## Environment variables

| Variable             | Required | Default   | Description                                                   |
| -------------------- | -------- | --------- | ------------------------------------------------------------ |
| `Graph_TenantId`     | yes      | —         | Entra ID (Azure AD) tenant ID.                               |
| `Graph_ClientId`     | yes      | —         | App registration (client) ID.                                |
| `Graph_ClientSecret` | yes      | —         | App registration client secret.                              |
| `SendFrom`           | yes      | —         | Mailbox/UPN that mail is sent from (e.g. `noreply@…`).       |
| `LogLevel`           | no       | `INFO`    | Python log level: `DEBUG`, `INFO`, `WARNING`, `ERROR`.       |
| `LogPath`            | no       | —         | File or directory for a persistent event log (see below).    |
| `Smtp_Host`          | no       | `0.0.0.0` | Address the SMTP server binds to.                            |
| `Smtp_Port`          | no       | `25`      | Port the SMTP server listens on.                             |

Both single- and double-underscore spellings are accepted for the Graph and
SMTP variables (e.g. `Graph_ClientSecret` and `Graph__ClientSecret`,
`Smtp_Port` and `Smtp__Port`), so you can match whatever convention you already
use.

## Logging

Logs always go to stdout (visible via `docker logs`). If you set `LogPath`,
the relay **also** writes to a rotating file (10 MB × 5 backups). Point it at a
directory (a `relay.log` file is created inside) or a specific file path, and
mount a volume there for persistence:

```yaml
environment:
  LogPath: "/var/log/smtp-relay"
volumes:
  - ./logs:/var/log/smtp-relay
```

Each received message produces a structured event line:

```
2026-06-29 12:00:00 INFO smtp-relay.events: RECEIVED ip=172.18.0.5 mail_from=app@example.com from='App <app@example.com>' to=alice@example.com,bob@example.com cc=carol@example.com subject='Nightly report' size=2048
2026-06-29 12:00:01 INFO smtp-relay.events: SENT ip=172.18.0.5 to=alice@example.com,bob@example.com
```

`RECEIVED` is logged for every message with the peer IP, envelope `MAIL FROM`,
the `From`/`To`/`Cc` headers, subject, and size; `SENT`, `FAILED`, or
`REJECTED` records the outcome.

## Health check

The image ships with a Docker `HEALTHCHECK` that opens an SMTP session against
the listener every 30s and verifies the greeting. If the relay locks up, the
container is marked `unhealthy` so your orchestrator (or
`restart: unless-stopped` plus a watchdog) can restart it. Check status with
`docker inspect --format '{{.State.Health.Status}}' smtp-relay`.

## Azure / Entra ID setup

1. Register an application in **Entra ID → App registrations**.
2. Under **Certificates & secrets**, create a **client secret** → use it as
   `Graph_ClientSecret`.
3. Under **API permissions**, add the **Application** permission
   `Mail.Send` (Microsoft Graph) and grant admin consent.
4. (Recommended) Scope the app to a single mailbox with an
   [application access policy](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access)
   so it can only send from `SendFrom`.

## Usage

### Docker Compose

```yaml
services:
  smtp-relay:
    image: ghcr.io/dmitaj/microsoft-365-smtp-relay:latest
    restart: unless-stopped
    environment:
      Graph_TenantId: "00000000-0000-0000-0000-000000000000"
      Graph_ClientId: "00000000-0000-0000-0000-000000000000"
      Graph__ClientSecret: "your-app-client-secret"
      SendFrom: "noreply@yourdomain.com"

  your-app:
    image: your/app
    environment:
      SMTP_HOST: smtp-relay   # other containers reach it by service name
      SMTP_PORT: 25
```

Put the relay and your apps on the same Docker network and point them at
`smtp-relay:25`. No need to publish a host port unless you want to reach it
from the host.

### Updating / redeploying (Unraid host)

`deploy.sh` updates the local clone, builds the image, pushes it, and recreates
the running container with the new image — run it on the Docker host:

```bash
cd /mnt/user/github/Microsoft-365-SMTP-Relay
./deploy.sh                 # pull, build, push, recreate
./deploy.sh --no-push       # local build only, then recreate
./deploy.sh --no-restart    # build + push, leave the container running
```

It recreates the container by snapshotting its existing `docker run` command
(via the `runlike` helper) so all your Unraid template settings — env vars,
the appdata log volume, ports, restart policy — are preserved. Override
`IMAGE`, `CONTAINER`, `BRANCH`, or `REPO_DIR` as environment variables if your
names differ. Pushing requires being logged in to the registry
(`docker login ghcr.io`).

### Plain Docker

```bash
docker build -t microsoft-365-smtp-relay .

docker run -d --name smtp-relay \
  -e Graph_TenantId=... \
  -e Graph_ClientId=... \
  -e Graph_ClientSecret=... \
  -e SendFrom=noreply@yourdomain.com \
  -e LogPath=/var/log/smtp-relay \
  -v "$PWD/logs:/var/log/smtp-relay" \
  -p 2525:25 \
  microsoft-365-smtp-relay
```

### Test it

```bash
python3 - <<'EOF'
import smtplib
from email.message import EmailMessage

msg = EmailMessage()
msg["From"] = "noreply@yourdomain.com"
msg["To"] = "you@yourdomain.com"
msg["Subject"] = "SMTP relay test"
msg.set_content("It works!")

with smtplib.SMTP("localhost", 2525) as s:
    s.send_message(msg)
EOF
```

## Notes & limitations

- Mail is sent **from the message's `From:` address** when present, falling
  back to the SMTP `MAIL FROM` and finally to the configured `SendFrom`
  mailbox. The sending app registration must have permission to send as that
  address — with `Mail.Send` scoped by an
  [application access policy](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access),
  only allowed mailboxes will succeed; others get rejected by Graph. Use
  `SendFrom` as the safe default for apps that don't set a `From:`.
- **Non-user senders (e.g. Microsoft 365 Groups):** a group address isn't a
  user mailbox, so Graph rejects a direct send with `404 ErrorInvalidUser`.
  When that happens, the relay automatically retries through the `SendFrom`
  mailbox while keeping the original address as the `From:`. For this to
  succeed, grant `SendFrom` **Send As** rights on the group in Exchange Online:
  ```powershell
  $g = Get-Recipient -RecipientTypeDetails GroupMailbox -Identity "<groupAlias>"
  Add-RecipientPermission -Identity $g.Name -Trustee "<SendFrom mailbox>" -AccessRights SendAs
  ```
  The same fallback covers shared mailboxes and other Send-As addresses that
  aren't directly addressable as a user.
- Recipients come from the SMTP envelope (`RCPT TO`); the relay falls back to
  the `To`/`Cc` headers only if the envelope has none.
- Plain-text, HTML, and file attachments are supported. The relay does not do
  TLS, AUTH, or queuing/retries — if a send fails it returns a `451` so the
  client can retry.
- `saveToSentItems` is disabled, so relayed mail does not clutter the
  `SendFrom` mailbox's Sent folder.

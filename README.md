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

| Variable              | Required | Default     | Description                                            |
| --------------------- | -------- | ----------- | ------------------------------------------------------ |
| `Graph_TenantId`      | yes      | —           | Entra ID (Azure AD) tenant ID.                         |
| `Graph_ClientId`      | yes      | —           | App registration (client) ID.                          |
| `Graph__ClientSecret` | yes      | —           | App registration client secret.                        |
| `SendFrom`            | yes      | —           | Mailbox/UPN that mail is sent from (e.g. `noreply@…`). |
| `LogLevel`            | no       | `INFO`      | Python log level: `DEBUG`, `INFO`, `WARNING`, `ERROR`. |
| `Smtp_Host`           | no       | `0.0.0.0`   | Address the SMTP server binds to.                      |
| `Smtp_Port`           | no       | `25`        | Port the SMTP server listens on.                       |

Both single- and double-underscore spellings are accepted for the Graph and
SMTP variables (e.g. `Graph_ClientSecret` and `Graph__ClientSecret`,
`Smtp_Port` and `Smtp__Port`), so you can match whatever convention you already
use.

## Azure / Entra ID setup

1. Register an application in **Entra ID → App registrations**.
2. Under **Certificates & secrets**, create a **client secret** → use it as
   `Graph__ClientSecret`.
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

### Plain Docker

```bash
docker build -t microsoft-365-smtp-relay .

docker run -d --name smtp-relay \
  -e Graph_TenantId=... \
  -e Graph_ClientId=... \
  -e Graph__ClientSecret=... \
  -e SendFrom=noreply@yourdomain.com \
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

- Mail is always sent **from `SendFrom`** regardless of the SMTP `MAIL FROM` /
  `From:` header (Graph sends from the mailbox in the API path).
- Recipients come from the SMTP envelope (`RCPT TO`); the relay falls back to
  the `To`/`Cc` headers only if the envelope has none.
- Plain-text, HTML, and file attachments are supported. The relay does not do
  TLS, AUTH, or queuing/retries — if a send fails it returns a `451` so the
  client can retry.
- `saveToSentItems` is disabled, so relayed mail does not clutter the
  `SendFrom` mailbox's Sent folder.

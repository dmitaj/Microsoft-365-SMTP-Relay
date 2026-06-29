FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY relay.py healthcheck.py ./

# Default SMTP port (override with Smtp_Port). Runs as root so it can bind
# the privileged default port 25; this is an internal relay meant to live on
# a private Docker network.
EXPOSE 25

HEALTHCHECK --interval=30s --timeout=10s --start-period=10s --retries=3 \
    CMD python healthcheck.py

CMD ["python", "relay.py"]

# SeePlaces Tools — web deployment image (Flask + headless Chrome for Selenium)
FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1

# System dependencies + Google Chrome (Selenium drives it to log in to SeePlaces)
RUN apt-get update && apt-get install -y --no-install-recommends \
        wget gnupg ca-certificates fonts-liberation \
    && wget -q -O /tmp/chrome.deb https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb \
    && apt-get install -y --no-install-recommends /tmp/chrome.deb \
    && rm -f /tmp/chrome.deb \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt

COPY . .

# Render injects $PORT at runtime; default for local docker runs.
ENV PORT=10000
EXPOSE 10000

# Single worker (in-memory job queue is shared across threads), many threads for SSE.
CMD ["sh", "-c", "gunicorn -w 1 -k gthread --threads 8 --timeout 300 -b 0.0.0.0:${PORT} app:app"]

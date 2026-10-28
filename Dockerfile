FROM python:3.13-slim AS build

COPY --from=ghcr.io/astral-sh/uv:0.9.5 /uv /uvx /bin/
WORKDIR /app
COPY pyproject.toml uv.lock README.md ./
COPY microsoft365mcp ./microsoft365mcp
COPY main.py ./
RUN uv sync --frozen --no-dev --no-editable

FROM python:3.13-slim

RUN useradd --create-home --uid 10001 --shell /usr/sbin/nologin app \
    && mkdir -p /data/tokens \
    && chown app:app /data/tokens
WORKDIR /app
COPY --from=build --chown=app:app /app/.venv /app/.venv
COPY docker/entrypoint.sh /entrypoint.sh
RUN chmod 755 /entrypoint.sh

ENV PATH="/app/.venv/bin:$PATH" \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    M365_TRANSPORT=http \
    M365_HOST=0.0.0.0 \
    M365_PORT=8000 \
    M365_TOKEN_DIR=/data/tokens

EXPOSE 8000
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD ["python", "-c", "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8000/health', timeout=3)"]
ENTRYPOINT ["/entrypoint.sh"]
CMD ["python", "-m", "microsoft365mcp"]

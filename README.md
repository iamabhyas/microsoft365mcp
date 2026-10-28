# Microsoft 365 MCP (read-only)

This service proxies the Claude Microsoft 365 connector at `https://microsoft365.mcp.claude.com/mcp`. Only read and search tools are exposed. Write tools are disabled.

The signed-in user's Microsoft 365 permissions still apply. Optional comma-separated limits:

- `M365_REACHABLE_MAILBOXES` — mailboxes this process may read
- `M365_REACHABLE_FOLDERS` — SharePoint or OneDrive folders this process may read

A mailbox or folder outside those lists is rejected before the connector is called.

## Local use

Cursor starts the stdio server from `.cursor/mcp.json`. These commands run it directly:

```bash
uv run fastmcp version
uv run fastmcp dev inspector main.py
uv run python -m microsoft365mcp
```

The first upstream call opens a browser to sign in. Set `M365_TOKEN_ENCRYPTION_KEY` and `M365_TOKEN_DIR` when that sign-in should survive a restart.

## HTTP deployment

Copy the example environment and fill in the two secrets:

```bash
cp .env.example .env
python -c "import secrets; print(secrets.token_urlsafe(32))"
python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"
```

`M365_API_TOKEN` is the bearer token clients send to this service. `M365_TOKEN_ENCRYPTION_KEY` encrypts the Microsoft 365 sign-in on disk. `M365_HTTP_ALLOWED_HOSTS` lists the Host values clients use when the process listens on `0.0.0.0`.

Sign in once on a machine with a browser, using the same key and token directory the service will mount:

```bash
mkdir -p data/tokens
set -a && source .env && set +a
M365_TOKEN_DIR=./data/tokens uv run python -m microsoft365mcp login
```

Then start the service:

```bash
docker compose up -d --build
```

The MCP endpoint is `http://localhost:8000/mcp`. Send `Authorization: Bearer $M365_API_TOKEN`. `GET /health` and `GET /ready` stay open for process checks.

The container runs the server as uid 10001, keeps the filesystem read-only, and stores tokens only under `/data/tokens`.

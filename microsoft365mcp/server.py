"""Read-only proxy for the Claude Microsoft 365 connector."""

from __future__ import annotations

import logging
import sys
from typing import Any

import anyio
import mcp_types
from fastmcp import Client, FastMCP
from fastmcp.client.auth.oauth import OAuth
from fastmcp.exceptions import ToolError
from fastmcp.server.middleware import CallNext, Middleware, MiddlewareContext
from fastmcp.server.transforms.visibility import Visibility
from fastmcp.tools.base import Tool, ToolResult
from starlette.requests import Request
from starlette.responses import JSONResponse, Response

from microsoft365mcp.auth import DeploymentTokenVerifier
from microsoft365mcp.catalog import READ_TOOLS
from microsoft365mcp.config import ConfigError, Settings, _valid_fernet_key
from microsoft365mcp.scope import ReachabilityPolicy
from microsoft365mcp.tokens import build_token_store

logger = logging.getLogger(__name__)

_INSTRUCTIONS = (
    "Read-only Microsoft 365 connector. "
    "Search and read mail, calendar, Teams, SharePoint, and OneDrive. "
    "Write tools are disabled."
)


def create_readonly_server(
    upstream: FastMCP | str | None = None,
    *,
    policy: ReachabilityPolicy | None = None,
    settings: Settings | None = None,
) -> FastMCP:
    """Expose the read-only connector without contacting Microsoft until a tool runs."""
    resolved = settings or Settings.from_env()
    if resolved.serves_http:
        resolved.validate_for_serve()
    server = FastMCP(
        name="microsoft365",
        instructions=_INSTRUCTIONS,
        auth=(
            DeploymentTokenVerifier(resolved.api_token)
            if resolved.api_token is not None
            else None
        ),
        mask_error_details=resolved.serves_http,
    )
    for name in sorted(READ_TOOLS):
        server.add_tool(ForwardingTool(name=name, upstream=upstream, settings=resolved))
    server.add_transform(Visibility(False, components={"tool"}))
    server.add_transform(Visibility(True, names=set(READ_TOOLS)))
    server.add_middleware(
        ReadOnlyScopeMiddleware(
            policy if policy is not None else resolved.policy,
        )
    )
    _install_health(server)
    return server


def main(argv: list[str] | None = None) -> None:
    args = sys.argv[1:] if argv is None else argv
    try:
        if args[:1] == ["login"]:
            login()
            return
        if args:
            raise ConfigError(
                [f"Unknown command {args[0]!r}. Serve with no command, or use login."]
            )
        settings = Settings.from_env()
        settings.validate_for_serve()
        _configure_logging(settings.log_level)
        server = create_readonly_server(settings=settings)
        if settings.serves_http:
            logger.info(
                "Serving Microsoft 365 read-only MCP on %s:%s",
                settings.host,
                settings.port,
            )
            server.run(
                transport="http",
                host=settings.host,
                port=settings.port,
                path="/mcp",
                show_banner=False,
                log_level=settings.log_level,
                stateless_http=True,
                host_origin_protection=True,
                allowed_hosts=list(settings.allowed_hosts),
            )
            return
        server.run(
            transport="stdio",
            show_banner=False,
            log_level=settings.log_level,
        )
    except ConfigError as exc:
        print(str(exc), file=sys.stderr)
        raise SystemExit(2) from exc


def login() -> None:
    """Sign in once and store the connector tokens for later server starts."""
    settings = Settings.from_env()
    if not settings.token_encryption_key or not _valid_fernet_key(
        settings.token_encryption_key
    ):
        raise ConfigError(
            [
                "Set M365_TOKEN_ENCRYPTION_KEY to a url-safe Fernet key "
                "before saving a Microsoft 365 sign-in."
            ]
        )
    _configure_logging(settings.log_level)
    client = _oauth_client(settings)

    async def _run() -> None:
        async with client:
            await client.list_tools()

    anyio.run(_run)
    print("Microsoft 365 sign-in saved.", file=sys.stderr)


class ForwardingTool(Tool):
    """A read tool whose schema is local and whose call goes to the connector."""

    def __init__(
        self,
        *,
        name: str,
        upstream: FastMCP | str | None,
        settings: Settings,
    ) -> None:
        super().__init__(
            name=name,
            description=f"Read-only Microsoft 365 {name.replace('_', ' ')}.",
            parameters={
                "type": "object",
                "additionalProperties": True,
                "properties": {
                    "query": {"type": "string"},
                    "mailbox": {"type": "string"},
                    "folder": {"type": "string"},
                    "uri": {"type": "string"},
                },
            },
        )
        self._upstream = upstream
        self._settings = settings

    async def run(self, arguments: dict[str, Any]) -> ToolResult:
        client = _client_for(self._upstream, self._settings)
        async with client:
            result = await client.call_tool_mcp(self.name, arguments or {})
        if result.is_error:
            text = "The Microsoft 365 connector returned an error."
            if result.content and isinstance(result.content[0], mcp_types.TextContent):
                text = result.content[0].text
            raise ToolError(text)
        return ToolResult.from_mcp_result(result)


class ReadOnlyScopeMiddleware(Middleware):
    """Reject write tools and mailboxes or folders outside the local policy."""

    def __init__(self, policy: ReachabilityPolicy) -> None:
        self.policy = policy

    async def on_call_tool(
        self,
        context: MiddlewareContext[mcp_types.CallToolRequestParams],
        call_next: CallNext[mcp_types.CallToolRequestParams, object],
    ) -> object:
        name = context.message.name
        if name not in READ_TOOLS:
            raise ToolError(
                f"Tool {name} is disabled. The Microsoft 365 connector is read-only."
            )
        arguments = context.message.arguments or {}
        violation = self.policy.violation(name, arguments)
        if violation:
            raise ToolError(violation)
        return await call_next(context)


def _install_health(server: FastMCP) -> None:
    @server.custom_route("/health", methods=["GET"], include_in_schema=False)
    async def health(_request: Request) -> Response:
        return JSONResponse({"status": "ok", "service": "microsoft365mcp"})

    @server.custom_route("/ready", methods=["GET"], include_in_schema=False)
    async def ready(_request: Request) -> Response:
        return JSONResponse({"status": "ready", "service": "microsoft365mcp"})


def _client_for(upstream: FastMCP | str | None, settings: Settings) -> Client:
    if isinstance(upstream, FastMCP):
        return Client(upstream)
    return _oauth_client(_settings_for_url(upstream, settings))


def _settings_for_url(upstream: str | None, settings: Settings) -> Settings:
    if upstream is None:
        return settings
    return Settings(
        transport=settings.transport,
        host=settings.host,
        port=settings.port,
        mcp_url=upstream,
        api_token=settings.api_token,
        token_dir=settings.token_dir,
        token_encryption_key=settings.token_encryption_key,
        oauth_callback_host=settings.oauth_callback_host,
        oauth_callback_port=settings.oauth_callback_port,
        allowed_hosts=settings.allowed_hosts,
        log_level=settings.log_level,
        policy=settings.policy,
    )


def _oauth_client(settings: Settings) -> Client:
    token_storage = None
    if settings.token_encryption_key:
        token_storage = build_token_store(
            settings.token_dir, settings.token_encryption_key
        )
    auth = OAuth(
        mcp_url=settings.mcp_url,
        client_name="microsoft365mcp",
        token_storage=token_storage,
        callback_host=settings.oauth_callback_host,
        callback_port=settings.oauth_callback_port,
    )
    return Client(settings.mcp_url, auth=auth, name="microsoft365-upstream")


def _configure_logging(level: str) -> None:
    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
        stream=sys.stderr,
    )

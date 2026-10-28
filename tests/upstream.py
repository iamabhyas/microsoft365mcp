"""In-memory stand-in for the hosted Microsoft 365 connector."""

from __future__ import annotations

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError

from microsoft365mcp.catalog import READ_TOOLS, WRITE_TOOLS

SENSITIVE = "sensitive mailbox content"
OWN_MAILBOX = "user@contoso.com"


def build_upstream(
    calls: list[dict[str, str | None]],
    *,
    enforce_graph_acl: bool = False,
) -> FastMCP:
    server = FastMCP("m365-upstream")
    names = sorted(READ_TOOLS | WRITE_TOOLS | {"outlook_future_write"})

    def register(name: str) -> None:
        def _tool(
            query: str = "",
            mailbox: str | None = None,
            folder: str | None = None,
            uri: str | None = None,
        ) -> dict[str, object]:
            calls.append(
                {
                    "name": name,
                    "query": query,
                    "mailbox": mailbox,
                    "folder": folder,
                    "uri": uri,
                }
            )
            if enforce_graph_acl and mailbox and mailbox.casefold() != OWN_MAILBOX:
                raise ToolError("Access denied: mailbox is out of reach")
            if enforce_graph_acl and folder and "secret" in folder.casefold():
                raise ToolError("Access denied: folder is out of reach")
            if enforce_graph_acl and uri and "secret" in uri.casefold():
                raise ToolError("Access denied: folder is out of reach")
            return {
                "messages": [{"subject": "Hello", "body": SENSITIVE}],
                "folders": [{"name": folder or query}],
                "content": "sensitive file content",
            }

        server.tool(_tool, name=name)

    for name in names:
        register(name)
    return server

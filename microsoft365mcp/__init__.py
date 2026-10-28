"""Read-only Microsoft 365 connector proxy."""

from microsoft365mcp.catalog import CONNECTOR_URL
from microsoft365mcp.server import create_readonly_server

__all__ = ["CONNECTOR_URL", "create_readonly_server"]

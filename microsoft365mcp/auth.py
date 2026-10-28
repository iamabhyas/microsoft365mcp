"""Bearer token check for the HTTP deployment."""

from __future__ import annotations

import secrets

from fastmcp.server.auth import AccessToken, TokenVerifier


class DeploymentTokenVerifier(TokenVerifier):
    """Accept one deployment token using a constant-time comparison."""

    def __init__(self, token: str) -> None:
        super().__init__()
        self._token = token

    async def verify_token(self, token: str) -> AccessToken | None:
        if len(token) != len(self._token) or not secrets.compare_digest(
            token, self._token
        ):
            return None
        return AccessToken(
            token=token,
            client_id="microsoft365mcp",
            scopes=["read"],
        )

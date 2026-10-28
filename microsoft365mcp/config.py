"""Runtime settings for a local session or an HTTP deployment."""

from __future__ import annotations

import os
from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path

from cryptography.fernet import Fernet
from platformdirs import user_data_dir

from microsoft365mcp.catalog import CONNECTOR_URL
from microsoft365mcp.scope import ReachabilityPolicy, policy_from_env

_LOG_LEVELS = frozenset({"DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"})
_PLACEHOLDER_TOKENS = frozenset(
    {"changeme", "secret", "password", "token", "replace-me", "your-token-here"}
)
_MIN_API_TOKEN_LENGTH = 32


class ConfigError(Exception):
    """The process cannot start with the current environment."""

    def __init__(self, errors: list[str]) -> None:
        self.errors = errors
        super().__init__("\n".join(errors))


@dataclass(frozen=True)
class Settings:
    transport: str = "stdio"
    host: str = "127.0.0.1"
    port: int = 8000
    mcp_url: str = CONNECTOR_URL
    api_token: str | None = None
    token_dir: Path = field(
        default_factory=lambda: Path(user_data_dir("microsoft365mcp")) / "tokens"
    )
    token_encryption_key: str | None = None
    oauth_callback_host: str = "127.0.0.1"
    oauth_callback_port: int | None = None
    allowed_hosts: tuple[str, ...] = ()
    log_level: str = "INFO"
    policy: ReachabilityPolicy = field(default_factory=ReachabilityPolicy)

    def __repr__(self) -> str:
        return (
            f"Settings(transport={self.transport!r}, host={self.host!r}, "
            f"port={self.port}, mcp_url={self.mcp_url!r})"
        )

    @property
    def serves_http(self) -> bool:
        return self.transport == "http"

    def validate_for_serve(self) -> None:
        if not self.serves_http:
            return
        errors: list[str] = []
        if self.api_token is None:
            errors.append("Set M365_API_TOKEN before serving over HTTP.")
        elif (
            len(self.api_token) < _MIN_API_TOKEN_LENGTH
            or self.api_token.casefold() in _PLACEHOLDER_TOKENS
            or any(character.isspace() for character in self.api_token)
        ):
            errors.append(
                "M365_API_TOKEN must be at least "
                f"{_MIN_API_TOKEN_LENGTH} characters and contain no spaces."
            )
        if not self.token_encryption_key:
            errors.append("Set M365_TOKEN_ENCRYPTION_KEY before serving over HTTP.")
        elif not _valid_fernet_key(self.token_encryption_key):
            errors.append(
                "M365_TOKEN_ENCRYPTION_KEY must be a url-safe Fernet key."
            )
        if _binds_every_interface(self.host) and not self.allowed_hosts:
            errors.append(
                "Set M365_HTTP_ALLOWED_HOSTS to the hostnames clients send "
                "when the server listens on every interface."
            )
        if not 1 <= self.port <= 65535:
            errors.append("M365_PORT must be between 1 and 65535.")
        if errors:
            raise ConfigError(errors)

    @classmethod
    def from_env(cls, environ: Mapping[str, str] | None = None) -> Settings:
        env = os.environ if environ is None else environ
        transport = _optional(env, "M365_TRANSPORT", "stdio").lower()
        if transport not in {"stdio", "http"}:
            raise ConfigError(
                ["M365_TRANSPORT must be stdio or http."]
            )
        log_level = _optional(env, "M365_LOG_LEVEL", "INFO").upper()
        if log_level not in _LOG_LEVELS:
            raise ConfigError(
                ["M365_LOG_LEVEL must be DEBUG, INFO, WARNING, ERROR, or CRITICAL."]
            )
        token_dir = _optional(
            env,
            "M365_TOKEN_DIR",
            str(Path(user_data_dir("microsoft365mcp")) / "tokens"),
        )
        return cls(
            transport=transport,
            host=_optional(env, "M365_HOST", "127.0.0.1"),
            port=_port(env.get("M365_PORT")),
            mcp_url=_optional(env, "M365_MCP_URL", CONNECTOR_URL),
            api_token=_blank_as_none(env.get("M365_API_TOKEN")),
            token_dir=Path(token_dir).expanduser(),
            token_encryption_key=_blank_as_none(env.get("M365_TOKEN_ENCRYPTION_KEY")),
            oauth_callback_host=_optional(env, "M365_OAUTH_CALLBACK_HOST", "127.0.0.1"),
            oauth_callback_port=_optional_port(env.get("M365_OAUTH_CALLBACK_PORT")),
            allowed_hosts=_csv(env.get("M365_HTTP_ALLOWED_HOSTS")),
            log_level=log_level,
            policy=policy_from_env(env),
        )


def _valid_fernet_key(key: str) -> bool:
    try:
        Fernet(key.encode("ascii"))
    except (UnicodeEncodeError, ValueError):
        return False
    return True


def _binds_every_interface(host: str) -> bool:
    return host.strip("[]") in {"0.0.0.0", "::"}


def _optional(env: Mapping[str, str], name: str, default: str) -> str:
    value = env.get(name)
    if value is None or not value.strip():
        return default
    return value.strip()


def _blank_as_none(value: str | None) -> str | None:
    if value is None or not value.strip():
        return None
    return value.strip()


def _port(value: str | None) -> int:
    if value is None or not value.strip():
        return 8000
    try:
        port = int(value)
    except ValueError as exc:
        raise ConfigError(["M365_PORT must be an integer."]) from exc
    return port


def _optional_port(value: str | None) -> int | None:
    if value is None or not value.strip():
        return None
    try:
        port = int(value)
    except ValueError as exc:
        raise ConfigError(["M365_OAUTH_CALLBACK_PORT must be an integer."]) from exc
    if not 1 <= port <= 65535:
        raise ConfigError(["M365_OAUTH_CALLBACK_PORT must be between 1 and 65535."])
    return port


def _csv(value: str | None) -> tuple[str, ...]:
    if value is None or not value.strip():
        return ()
    return tuple(part.strip() for part in value.split(",") if part.strip())

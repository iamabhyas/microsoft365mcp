import asyncio
from pathlib import Path

import pytest
from cryptography.fernet import Fernet

from microsoft365mcp.auth import DeploymentTokenVerifier
from microsoft365mcp.config import ConfigError, Settings
from microsoft365mcp.server import create_readonly_server
from microsoft365mcp.tokens import build_token_store
from tests.upstream import build_upstream

API_TOKEN = "a" * 32


def _http_env(**overrides: str) -> dict[str, str]:
    env = {
        "M365_TRANSPORT": "http",
        "M365_HOST": "0.0.0.0",
        "M365_PORT": "8000",
        "M365_API_TOKEN": API_TOKEN,
        "M365_TOKEN_ENCRYPTION_KEY": Fernet.generate_key().decode(),
        "M365_HTTP_ALLOWED_HOSTS": "mcp.example.com",
        "M365_TOKEN_DIR": "/tmp/microsoft365mcp-tokens",
    }
    env.update(overrides)
    return env


def test_stdio_starts_without_deployment_secrets() -> None:
    settings = Settings.from_env({})
    settings.validate_for_serve()
    assert settings.transport == "stdio"
    assert settings.api_token is None


def test_http_requires_token_key_and_allowed_hosts() -> None:
    with pytest.raises(ConfigError, match="M365_API_TOKEN"):
        Settings.from_env(
            _http_env(M365_API_TOKEN=""),
        ).validate_for_serve()
    with pytest.raises(ConfigError, match="M365_API_TOKEN"):
        Settings.from_env(_http_env(M365_API_TOKEN="short")).validate_for_serve()
    with pytest.raises(ConfigError, match="M365_TOKEN_ENCRYPTION_KEY"):
        Settings.from_env(
            _http_env(M365_TOKEN_ENCRYPTION_KEY="not-a-fernet-key"),
        ).validate_for_serve()
    with pytest.raises(ConfigError, match="M365_HTTP_ALLOWED_HOSTS"):
        Settings.from_env(
            _http_env(M365_HTTP_ALLOWED_HOSTS=""),
        ).validate_for_serve()


def test_http_settings_hide_the_token() -> None:
    settings = Settings.from_env(_http_env())
    settings.validate_for_serve()
    assert API_TOKEN not in repr(settings)


def test_deployment_token_verifier_accepts_only_the_configured_token() -> None:
    verifier = DeploymentTokenVerifier(API_TOKEN)

    accepted = asyncio.run(verifier.verify_token(API_TOKEN))
    rejected = asyncio.run(verifier.verify_token("b" * 32))
    mismatched = asyncio.run(verifier.verify_token("a"))

    assert accepted is not None
    assert accepted.scopes == ["read"]
    assert rejected is None
    assert mismatched is None


def test_token_store_round_trip_is_encrypted_on_disk(tmp_path: Path) -> None:
    key = Fernet.generate_key().decode()
    store = build_token_store(tmp_path, key)
    asyncio.run(store.put("session", {"refresh_token": "secret-refresh"}))

    loaded = asyncio.run(store.get("session"))
    raw = "\n".join(path.read_text() for path in tmp_path.rglob("*") if path.is_file())

    assert loaded == {"refresh_token": "secret-refresh"}
    assert "secret-refresh" not in raw
    assert oct(tmp_path.stat().st_mode & 0o777) == "0o700"


def test_http_server_requires_a_bearer_token_and_exposes_health() -> None:
    server = create_readonly_server(
        build_upstream([]),
        settings=Settings.from_env(_http_env()),
    )
    paths = {route.path for route in server._additional_http_routes}

    assert server.auth is not None
    assert server._mask_error_details is True
    assert {"/health", "/ready"} <= paths


def test_main_rejects_an_unknown_command(capsys: pytest.CaptureFixture[str]) -> None:
    from microsoft365mcp.server import main

    with pytest.raises(SystemExit) as caught:
        main(["deploy"])
    assert caught.value.code == 2
    assert "Unknown command" in capsys.readouterr().err

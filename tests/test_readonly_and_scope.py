import asyncio

import pytest
from fastmcp import Client
from fastmcp.exceptions import ToolError

from microsoft365mcp.catalog import CONNECTOR_URL, READ_TOOLS, WRITE_TOOLS
from microsoft365mcp.scope import ReachabilityPolicy, policy_from_env
from microsoft365mcp.server import create_readonly_server
from tests.upstream import OWN_MAILBOX, SENSITIVE, build_upstream

TEAM_FOLDER = "/sites/team/Shared Documents"
SECRET_FOLDER = "/sites/secret/Private"
SECRET_MAILBOX = "exec@contoso.com"


def test_connector_url_is_the_hosted_microsoft_365_server() -> None:
    assert CONNECTOR_URL == "https://microsoft365.mcp.claude.com/mcp"


def test_policy_from_env_parses_mailboxes_and_folders() -> None:
    policy = policy_from_env(
        {
            "M365_REACHABLE_MAILBOXES": " User@Contoso.com , shared@contoso.com ",
            "M365_REACHABLE_FOLDERS": f" {TEAM_FOLDER} ",
        }
    )
    assert policy.mailboxes == frozenset({"user@contoso.com", "shared@contoso.com"})
    assert policy.violation(
        "outlook_email_search", {"mailbox": SECRET_MAILBOX}
    ) == f"Mailbox {SECRET_MAILBOX} is out of reach."
    assert (
        policy.violation("sharepoint_folder_search", {"query": SECRET_FOLDER})
        == f"Folder {SECRET_FOLDER} is out of reach."
    )


def test_open_policy_does_not_name_a_local_violation() -> None:
    policy = policy_from_env({})
    assert policy.violation("outlook_email_search", {"mailbox": SECRET_MAILBOX}) is None
    assert policy.violation("read_resource", {"uri": SECRET_FOLDER}) is None


def test_read_tools_are_exposed_and_write_tools_are_disabled() -> None:
    calls: list[dict[str, str | None]] = []
    server = create_readonly_server(build_upstream(calls))

    names = asyncio.run(_tool_names(server))

    assert READ_TOOLS <= names
    assert names.isdisjoint(WRITE_TOOLS)
    assert "outlook_future_write" not in names


def test_write_tool_call_returns_no_content() -> None:
    calls: list[dict[str, str | None]] = []
    server = create_readonly_server(build_upstream(calls))

    message = asyncio.run(_error_text(server, "outlook_send_mail", {"query": "hello"}))

    assert "sensitive" not in message.casefold()
    assert calls == []


def test_mailbox_out_of_reach_is_rejected_before_the_connector() -> None:
    calls: list[dict[str, str | None]] = []
    server = create_readonly_server(
        build_upstream(calls, enforce_graph_acl=True),
        policy=ReachabilityPolicy(
            mailboxes=frozenset({OWN_MAILBOX}),
            folders=frozenset({TEAM_FOLDER}),
        ),
    )

    message = asyncio.run(
        _error_text(
            server,
            "outlook_email_search",
            {"query": f"mailbox:{SECRET_MAILBOX}", "mailbox": SECRET_MAILBOX},
        )
    )

    assert message == f"Mailbox {SECRET_MAILBOX} is out of reach."
    assert SENSITIVE not in message
    assert calls == []


def test_folder_out_of_reach_is_rejected_before_the_connector() -> None:
    calls: list[dict[str, str | None]] = []
    server = create_readonly_server(
        build_upstream(calls, enforce_graph_acl=True),
        policy=ReachabilityPolicy(
            mailboxes=frozenset({OWN_MAILBOX}),
            folders=frozenset({TEAM_FOLDER}),
        ),
    )

    folder_message = asyncio.run(
        _error_text(server, "sharepoint_folder_search", {"query": SECRET_FOLDER})
    )
    file_message = asyncio.run(
        _error_text(
            server,
            "read_resource",
            {
                "uri": "https://contoso.sharepoint.com/sites/secret/Private/budget.xlsx"
            },
        )
    )

    assert folder_message == f"Folder {SECRET_FOLDER} is out of reach."
    assert "budget.xlsx" in file_message
    assert "out of reach" in file_message
    assert "sensitive" not in folder_message.casefold()
    assert "sensitive" not in file_message.casefold()
    assert calls == []


def test_reachable_mailbox_and_folder_return_content() -> None:
    calls: list[dict[str, str | None]] = []
    server = create_readonly_server(
        build_upstream(calls),
        policy=ReachabilityPolicy(
            mailboxes=frozenset({OWN_MAILBOX}),
            folders=frozenset({TEAM_FOLDER}),
        ),
    )

    mail = asyncio.run(
        _call(server, "outlook_email_search", {"mailbox": OWN_MAILBOX, "query": "status"})
    )
    folder = asyncio.run(
        _call(
            server,
            "sharepoint_folder_search",
            {"query": f"{TEAM_FOLDER}/Finance"},
        )
    )

    assert SENSITIVE in str(mail)
    assert "Finance" in str(folder)
    assert [call["name"] for call in calls] == [
        "outlook_email_search",
        "sharepoint_folder_search",
    ]


def test_upstream_denial_returns_no_mailbox_or_folder_content() -> None:
    calls: list[dict[str, str | None]] = []
    server = create_readonly_server(
        build_upstream(calls, enforce_graph_acl=True),
        policy=ReachabilityPolicy(),
    )

    mailbox_message = asyncio.run(
        _error_text(server, "outlook_email_search", {"mailbox": SECRET_MAILBOX})
    )
    folder_message = asyncio.run(
        _error_text(
            server,
            "read_resource",
            {"folder": SECRET_FOLDER, "uri": f"folder://{SECRET_FOLDER}/notes.txt"},
        )
    )

    assert "Access denied" in mailbox_message
    assert "Access denied" in folder_message
    assert SENSITIVE not in mailbox_message
    assert "sensitive file content" not in folder_message
    assert [call["name"] for call in calls] == [
        "outlook_email_search",
        "read_resource",
    ]


async def _tool_names(server) -> set[str]:
    async with Client(server) as client:
        tools = await client.list_tools()
    return {tool.name for tool in tools}


async def _call(server, name: str, arguments: dict) -> object:
    async with Client(server) as client:
        return await client.call_tool(name, arguments)


async def _error_text(server, name: str, arguments: dict) -> str:
    with pytest.raises(ToolError) as caught:
        await _call(server, name, arguments)
    return str(caught.value)

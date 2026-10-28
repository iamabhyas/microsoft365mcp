"""Local reachability checks for mailboxes and folders.

Microsoft Graph still applies the signed-in user's own permissions. These
checks run first when an operator lists the mailboxes or folders this process
may read. A named mailbox or folder outside that list is rejected before the
connector is called.
"""

from __future__ import annotations

import os
import re
from collections.abc import Mapping
from dataclasses import dataclass
from typing import Any
from urllib.parse import unquote, urlparse

_MAILBOX_KEYS = ("mailbox", "mailbox_id", "shared_mailbox")
_FOLDER_KEYS = ("folder", "folder_id", "folder_path", "path")
_FOLDER_QUERY_TOOLS = frozenset({"sharepoint_folder_search"})
_MAILBOX_URI = re.compile(r"mailbox://([^/\s]+)", re.IGNORECASE)
_MAILBOX_QUERY = re.compile(r"(?:^|\s)mailbox:(\S+)", re.IGNORECASE)


@dataclass(frozen=True)
class ReachabilityPolicy:
    """Mailboxes and folders this process may read.

    ``None`` means that dimension is not restricted here. An empty set means
    every named mailbox or folder is out of reach.
    """

    mailboxes: frozenset[str] | None = None
    folders: frozenset[str] | None = None

    def __post_init__(self) -> None:
        if self.mailboxes is not None:
            object.__setattr__(
                self,
                "mailboxes",
                frozenset(mailbox.casefold() for mailbox in self.mailboxes),
            )
        if self.folders is not None:
            object.__setattr__(
                self,
                "folders",
                frozenset(folder.strip() for folder in self.folders if folder.strip()),
            )

    def violation(self, tool_name: str, arguments: Mapping[str, Any]) -> str | None:
        for mailbox in mailbox_targets(arguments):
            if self.mailboxes is not None and mailbox.casefold() not in self.mailboxes:
                return f"Mailbox {mailbox} is out of reach."
        for folder in folder_targets(tool_name, arguments):
            if self.folders is not None and not folder_in_reach(folder, self.folders):
                return f"Folder {folder} is out of reach."
        return None


def policy_from_env(environ: Mapping[str, str] | None = None) -> ReachabilityPolicy:
    source = os.environ if environ is None else environ
    return ReachabilityPolicy(
        mailboxes=_parse_list(source.get("M365_REACHABLE_MAILBOXES")),
        folders=_parse_list(source.get("M365_REACHABLE_FOLDERS")),
    )


def mailbox_targets(arguments: Mapping[str, Any]) -> list[str]:
    found: list[str] = []
    for key in _MAILBOX_KEYS:
        found.extend(_strings(arguments.get(key)))
    for key in ("query", "q"):
        for value in _strings(arguments.get(key)):
            found.extend(match.group(1) for match in _MAILBOX_QUERY.finditer(value))
    for value in _strings(arguments.get("uri")):
        match = _MAILBOX_URI.search(unquote(value))
        if match:
            found.append(unquote(match.group(1)))
    return found


def folder_targets(tool_name: str, arguments: Mapping[str, Any]) -> list[str]:
    found: list[str] = []
    for key in _FOLDER_KEYS:
        found.extend(_strings(arguments.get(key)))
    if tool_name in _FOLDER_QUERY_TOOLS:
        found.extend(_strings(arguments.get("query")))
        found.extend(_strings(arguments.get("q")))
    for value in _strings(arguments.get("uri")):
        text = unquote(value).strip()
        if text.lower().startswith(("http://", "https://", "folder://")):
            found.append(text)
    return found


def folder_in_reach(target: str, allowed: frozenset[str]) -> bool:
    for candidate in _folder_forms(target):
        folded = candidate.casefold()
        for item in allowed:
            base = item.strip("/").casefold()
            if not base:
                continue
            if folded == base or folded.startswith(base + "/"):
                return True
            if folded == base.rsplit("/", 1)[-1]:
                return True
    return False


def _folder_forms(target: str) -> list[str]:
    text = unquote(target).strip()
    forms = [text.strip("/")]
    lowered = text.lower()
    if lowered.startswith(("http://", "https://")):
        forms.append(urlparse(text).path.strip("/"))
    elif lowered.startswith("folder://"):
        forms.append(text[len("folder://") :].strip("/"))
    return [form for form in forms if form]


def _parse_list(raw: str | None) -> frozenset[str] | None:
    if raw is None or not raw.strip():
        return None
    return frozenset(part.strip() for part in raw.split(",") if part.strip())


def _strings(value: Any) -> list[str]:
    if isinstance(value, str):
        text = value.strip()
        return [text] if text else []
    if isinstance(value, list):
        return [item.strip() for item in value if isinstance(item, str) and item.strip()]
    return []

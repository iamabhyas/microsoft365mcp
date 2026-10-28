"""Encrypted on-disk storage for the upstream Microsoft 365 sign-in."""

from __future__ import annotations

import logging
from pathlib import Path

from cryptography.fernet import Fernet
from key_value.aio.protocols import AsyncKeyValue
from key_value.aio.stores.filetree import FileTreeStore
from key_value.aio.wrappers.encryption import FernetEncryptionWrapper

logger = logging.getLogger(__name__)


def build_token_store(directory: Path, encryption_key: str) -> AsyncKeyValue:
    """Store connector tokens under ``directory``, encrypted with ``encryption_key``."""
    directory.mkdir(parents=True, exist_ok=True)
    try:
        directory.chmod(0o700)
    except OSError:
        logger.warning("Could not restrict permissions on %s", directory)
    return FernetEncryptionWrapper(
        FileTreeStore(data_directory=directory),
        fernet=Fernet(encryption_key.encode("ascii")),
    )

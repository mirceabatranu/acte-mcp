"""Parolele de aplicatie Gmail, in Windows Credential Manager (niciodata in fisiere)."""

from __future__ import annotations

import os

import keyring

SERVICE = os.environ.get("ACTE_MCP_KEYRING_SERVICE", "acte-mcp")


def get(email: str) -> str:
    value = keyring.get_password(SERVICE, email.lower())
    if not value:
        raise RuntimeError(
            f"Nu exista parola de aplicatie pentru {email}. Salveaz-o cu: acte-mcp parola {email}"
        )
    return value


def save(email: str, value: str) -> None:
    # Google afiseaza parola in grupuri de cate 4 litere; spatiile nu fac parte din ea.
    keyring.set_password(SERVICE, email.lower(), value.replace(" ", "").strip())


def has(email: str) -> bool:
    return bool(keyring.get_password(SERVICE, email.lower()))

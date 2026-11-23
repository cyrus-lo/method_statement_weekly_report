"""Normalize ECM outgoing references for revision-free lookup."""

from __future__ import annotations


def outgoing_reference_without_revision(outgoing_reference: str, document_revision: str) -> str:
    """Remove a matching terminal slash or parenthesized revision suffix."""
    reference = outgoing_reference.strip()
    revision = document_revision.strip()
    if not revision:
        return reference

    for suffix in (f"({revision})", f"({revision}", f"/{revision}"):
        if reference.endswith(suffix):
            return reference[:-len(suffix)]
    return reference

"""Shared error types + stable CLI exit codes for the office-comments skill."""
from __future__ import annotations


class CommentError(Exception):
    """Base class for all expected, user-facing failures."""


class UnsupportedFile(CommentError):
    """The file isn't a .docx/.xlsx/.pptx we can handle."""


class AnchorNotFound(CommentError):
    """The requested anchor (text / cell / slide) doesn't exist in the file."""


class AmbiguousAnchor(CommentError):
    """The anchor text matches in several places; the caller must say which one."""


class CommentNotFound(CommentError):
    """No comment with the given id exists in the file."""


# Stable exit codes used by all CLIs (mirrors the oml-docs convention).
EXIT_OK = 0
EXIT_GENERIC = 1
EXIT_BAD_FILE = 3      # not an Office file we support / corrupted
EXIT_ANCHOR = 4        # anchor not found
EXIT_AMBIGUOUS = 5     # anchor matched in multiple places
EXIT_NO_COMMENT = 6    # referenced comment id is missing

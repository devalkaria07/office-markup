"""Small shared helpers for the comment modules: IDs, timestamps, initials.

These run in normal CPython, so uuid/random/datetime are all available (the sandbox
limits that forbid them elsewhere do not apply to this shipped skill code)."""
from __future__ import annotations

import random
import uuid
from datetime import datetime, timezone


def hex8(exclude=()) -> str:
    """A random 8-char UPPER hex token — Word's paraId / durableId style — not in `exclude`.

    Only 31 bits: Word treats w14:paraId / w15:paraIdParent as SIGNED 32-bit integers, and a
    value with the high bit set (first hex digit 8-F) silently breaks comment threading — the
    reply link won't match. Real Word only ever emits paraIds below 0x80000000."""
    bad = set(exclude)
    while True:
        v = f"{random.getrandbits(31):08X}"
        if v not in bad:
            return v


def guid(brace: bool = True, exclude=()) -> str:
    """A random uppercase GUID — Excel / PowerPoint id style — braced by default."""
    bad = set(exclude)
    while True:
        g = str(uuid.uuid4()).upper()
        v = "{" + g + "}" if brace else g
        if v not in bad:
            return v


def utc_now() -> datetime:
    return datetime.now(timezone.utc)


def iso_z(dt: datetime | None = None) -> str:
    """UTC ISO-8601 with a Z suffix, no microseconds, e.g. 2026-06-27T01:38:00Z."""
    dt = dt or utc_now()
    if dt.tzinfo is not None:
        dt = dt.astimezone(timezone.utc)
    return dt.strftime("%Y-%m-%dT%H:%M:%SZ")


def local_z(dt: datetime | None = None) -> str:
    """Local wall-clock time + a Z suffix — mirrors Word's (technically loose) w:date format,
    so the comment shows the author's local time rather than UTC."""
    return (dt or datetime.now()).strftime("%Y-%m-%dT%H:%M:%SZ")


def initials(name: str, maxlen: int = 3) -> str:
    """Derive initials from a display name: 'Alex Morgan' -> 'AM'."""
    parts = [p for p in name.replace(".", " ").replace("-", " ").split() if p]
    return "".join(p[0].upper() for p in parts[:maxlen]) or "?"

from __future__ import annotations

import hashlib
import re
import secrets


WORKBOOK_ID_RE = re.compile(r"^wb_g(?P<generation>\d+)_[0-9a-f]{32}$")
WORKBOOK_ID_FULL_RE = re.compile(r"^wb_g(?P<generation>\d+)_(?P<fingerprint>[0-9a-f]{8})_[0-9a-f]{24}$")


def generate_id(prefix: str) -> str:
    return f"{prefix}_{secrets.token_hex(16)}"


def generate_workbook_id(generation: int, runtime_fingerprint: str | None = None) -> str:
    if runtime_fingerprint:
        suffix = secrets.token_hex(12)
        return f"wb_g{generation}_{runtime_fingerprint[:8]}_{suffix}"
    return f"wb_g{generation}_{secrets.token_hex(16)}"


def parse_workbook_generation(workbook_id: str) -> int | None:
    full_match = WORKBOOK_ID_FULL_RE.match(workbook_id)
    if full_match:
        return int(full_match.group("generation"))
    match = WORKBOOK_ID_RE.match(workbook_id)
    if not match:
        return None
    return int(match.group("generation"))


def parse_workbook_fingerprint(workbook_id: str) -> str | None:
    full_match = WORKBOOK_ID_FULL_RE.match(workbook_id)
    if full_match:
        return full_match.group("fingerprint")
    return None


def compute_runtime_fingerprint(pipe_name: str, data_dir: str) -> str:
    raw = f"{pipe_name}:{data_dir}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:8]


def is_same_runtime_fingerprint(fp1: str | None, fp2: str | None) -> bool:
    if fp1 is None or fp2 is None:
        return True
    return fp1 == fp2

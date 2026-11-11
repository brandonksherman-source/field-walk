"""File and folder names that are safe on Windows, Drive and OneDrive."""
from __future__ import annotations
import re
from datetime import date
from pathlib import Path

QUEUE = ("1 - Ready for Review", "2 - Sent", "3 - Passed")


def safe_name(s: str) -> str:
    s = re.sub(r"[/\\]", "-", s)
    s = re.sub(r'[:"*?<>|\x00-\x1f]', "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s.rstrip(". ") or "Untitled"


def job_folder(root: Path, d: date, client: str, project: str) -> Path:
    f = Path(root) / QUEUE[0] / f"{d.isoformat()} {safe_name(client)} - {safe_name(project)}"
    (f / "photos").mkdir(parents=True, exist_ok=True)
    return f

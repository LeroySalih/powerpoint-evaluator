"""The JSON requirements spec and the report it produces."""
from __future__ import annotations

import base64
from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, Field

Status = Literal["pass", "fail", "review"]


class Asset(BaseModel):
    """A reference image. Give `file` (relative to the spec) or `data_base64` (embedded in the JSON)."""
    file: str | None = None
    data_base64: str | None = None

    def load(self, base: Path | None) -> bytes | None:
        if self.data_base64:
            return base64.b64decode(self.data_base64)
        if self.file and base and (base / self.file).is_file():
            return (base / self.file).read_bytes()
        return None


class CheckSpec(BaseModel):
    id: str
    description: str
    marks: float = 1
    type: str
    params: dict[str, Any] = Field(default_factory=dict)
    tags: list[str] = Field(default_factory=list)  # e.g. ["LO1"]


class Section(BaseModel):
    id: str
    title: str
    checks: list[CheckSpec]


class Spec(BaseModel):
    spec_version: str = "0.1"
    id: str
    title: str
    assets: dict[str, Asset] = Field(default_factory=dict)
    sections: list[Section]


class CheckResult(BaseModel):
    id: str
    description: str
    marks_available: float
    marks_awarded: float
    status: Status
    detail: str = ""
    tags: list[str] = Field(default_factory=list)


class Report(BaseModel):
    spec_id: str
    title: str
    results: list[CheckResult]
    marks_available: float
    marks_awarded: float  # confirmed passes only
    marks_pending_review: float  # worth awarding if the review items pass
    by_tag: dict[str, dict[str, float]]

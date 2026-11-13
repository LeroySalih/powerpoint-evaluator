from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from ..images import ImageId, identify
from ..inspect import Deck
from ..spec import Asset, Status

EMU_PER_CM = 360_000


@dataclass
class Outcome:
    status: Status
    detail: str = ""


def ok(detail="") -> Outcome: return Outcome("pass", detail)
def bad(detail) -> Outcome: return Outcome("fail", detail)
def review(detail) -> Outcome: return Outcome("review", detail)


@dataclass
class Context:
    deck: Deck
    assets: dict[str, tuple[ImageId, bytes | None]]  # name -> identity (bytes kept for contrast maths)
    extras: dict[str, bytes] = field(default_factory=dict)  # name -> uploaded PDF/screenshot/answers
    candidate: dict = field(default_factory=dict)  # optional {"name":..., "class":...}; never logged

    @staticmethod
    def load_assets(assets: dict[str, Asset], base: Path | None):
        out = {}
        for name, a in assets.items():
            data = a.load(base)
            if data:
                out[name] = (identify(data), data)
        return out


Check = Callable[[Context, dict], Outcome]
REGISTRY: dict[str, Check] = {}


def check(name: str):
    def deco(fn: Check) -> Check:
        REGISTRY[name] = fn
        return fn
    return deco


def combine(outcomes: list[Outcome]) -> Outcome:
    """Any fail -> fail; else any review -> review; else pass."""
    fails = [o.detail for o in outcomes if o.status == "fail"]
    if fails:
        return bad("; ".join(fails))
    revs = [o.detail for o in outcomes if o.status == "review"]
    if revs:
        return review("; ".join(revs))
    return ok("; ".join(o.detail for o in outcomes if o.detail))

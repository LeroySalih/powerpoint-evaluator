"""Run a spec against a submission and build the report."""
from __future__ import annotations

from pathlib import Path

from .checks import REGISTRY, Context
from .checks.base import Outcome, combine
from .inspect import BadPackage, inspect_pptx
from .spec import CheckResult, Report, Spec


def _run(ctx: Context, type_: str, params: dict) -> Outcome:
    fn = REGISTRY.get(type_)
    return fn(ctx, params) if fn else Outcome("review", f"unknown check type {type_!r}")


def evaluate(spec: Spec, pptx: bytes, extras: dict[str, bytes] | None = None, candidate: dict | None = None,
             spec_dir: Path | None = None) -> Report:
    ctx = Context(inspect_pptx(pptx), Context.load_assets(spec.assets, spec_dir), extras or {}, candidate or {})
    results: list[CheckResult] = []
    for sec in spec.sections:
        for c in sec.checks:
            try:  # a broken check must never sink the whole report
                if c.type == "all_of":  # every sub-check must pass; any fail fails; else any review -> review
                    out = combine([_run(ctx, s["type"], s.get("params", {})) for s in c.params["checks"]])
                else:
                    out = _run(ctx, c.type, c.params)
            except Exception as e:
                out = Outcome("review", f"check error: {type(e).__name__}")
            results.append(CheckResult(
                id=c.id, description=c.description, marks_available=c.marks,
                marks_awarded=c.marks if out.status == "pass" else 0, status=out.status,
                detail=out.detail, tags=c.tags))
    avail = sum(r.marks_available for r in results)
    got = sum(r.marks_awarded for r in results)
    pend = sum(r.marks_available for r in results if r.status == "review")
    by_tag: dict[str, dict[str, float]] = {}
    for r in results:
        for t in r.tags:
            d = by_tag.setdefault(t, {"available": 0, "awarded": 0, "pending_review": 0})
            d["available"] += r.marks_available
            d["awarded"] += r.marks_awarded
            d["pending_review"] += r.marks_available if r.status == "review" else 0
    return Report(spec_id=spec.id, title=spec.title, results=results, marks_available=avail,
                  marks_awarded=got, marks_pending_review=pend, by_tag=by_tag)


__all__ = ["evaluate", "BadPackage"]

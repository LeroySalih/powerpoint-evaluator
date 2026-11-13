"""Checks that need a judge: screenshots in EVIDENCE and written answers. Without a judge they return `review`."""
from __future__ import annotations

from .base import Context, Outcome, bad, check, review
from .pdf_checks import get_extra


@check("evidence.screenshot")
def screenshot(ctx: Context, p: dict) -> Outcome:
    """A readable screenshot under heading `p['heading']` showing `p['expect']`."""
    if get_extra(ctx, p.get("file", "EVIDENCE")) is None:
        return bad("EVIDENCE file not supplied (no screenshot, no mark)")
    return review(f"check screenshot {p.get('heading', '')}: {p.get('expect', '')}")


@check("answer.rubric")
def answer(ctx: Context, p: dict) -> Outcome:
    if get_extra(ctx, p.get("file", "EVIDENCE")) is None:
        return bad("EVIDENCE file not supplied")
    return review(f"mark {p.get('question', '')} against the rubric")

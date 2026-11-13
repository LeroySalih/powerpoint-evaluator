"""Checks on the pupil's printed PDFs (HANDOUTS, NOTES, FULL PAGE) and on who the work belongs to."""
from __future__ import annotations

import re
from io import BytesIO

from pypdf import PdfReader

from .base import Context, Outcome, bad, check, combine, ok, review
from .pptx_checks import norm, slide_at


def extra_key(s: str) -> str:
    s = re.sub(r"\.[A-Za-z0-9]{2,5}$", "", s.strip())
    return re.sub(r"[\s_\-]+", " ", s).casefold()


def get_extra(ctx: Context, name: str) -> bytes | None:
    want = extra_key(name)
    return next((v for k, v in ctx.extras.items() if extra_key(k) == want), None)


def read_pdf(data: bytes):
    try:
        r = PdfReader(BytesIO(data))
        text = " ".join((pg.extract_text() or "") for pg in r.pages)
        pages = [(float(pg.mediabox.width), float(pg.mediabox.height), pg.rotation or 0) for pg in r.pages]
        return pages, norm(text)
    except Exception as e:  # corrupt / encrypted
        raise ValueError(f"cannot read PDF ({type(e).__name__})") from e


def _orient(w: float, h: float, rot: int) -> str:
    if rot in (90, 270):
        w, h = h, w
    return "landscape" if w > h else "portrait"


@check("pdf.print")
def pdf_print(ctx: Context, p: dict) -> Outcome:
    data = get_extra(ctx, p["file"])
    if data is None:
        return bad(f"{p['file']} PDF not supplied")
    try:
        pages, text = read_pdf(data)
    except ValueError as e:
        return bad(str(e))
    problems = []
    if "pages" in p and len(pages) != p["pages"]:
        problems.append(f"{len(pages)} page(s), expected {p['pages']}")
    if "orientation" in p and pages and any(_orient(*pg) != p["orientation"] for pg in pages):
        problems.append(f"pages are not {p['orientation']}")
    expect: list[str] = []
    if p.get("slide_titles") == "all":
        expect += [s.title.text for s in ctx.deck.slides if s.title]
    elif p.get("slide_title"):
        s = slide_at(ctx.deck, p["slide_title"])
        expect += [s.title.text] if s and s.title else []
    if not text:
        return bad("; ".join(problems)) if problems else review("PDF has no extractable text (printed as images?); check by eye")
    for t in expect:
        if norm(t) not in text:
            problems.append(f"slide title {t!r} not found in the PDF")
    if p.get("notes_of_slide"):
        s = slide_at(ctx.deck, p["notes_of_slide"])
        words = set(norm(s.notes).split()) if s else set()
        if not words or len(words & set(text.split())) / len(words) < 0.7:
            problems.append("the slide's speaker notes are not on the page")
    return bad("; ".join(problems)) if problems else ok(f"{len(pages)} page(s)")


@check("print.name_class")
def name_class(ctx: Context, p: dict) -> Outcome:
    """Name and class in the Notes & Handouts footer. Needs candidate {name, class}; otherwise a human looks."""
    footers = [norm(t) for t in ctx.deck.notes_handout_text]
    cand = ctx.candidate
    if not cand.get("name") or not cand.get("class"):
        return review(f"no candidate name/class supplied; footer text in file: {'yes' if footers else 'none'}")
    hay = " ".join(footers)
    for f in p.get("files", ["HANDOUTS", "NOTES"]):
        d = get_extra(ctx, f)
        if d:
            try:
                hay += " " + read_pdf(d)[1]
            except ValueError:
                pass
    if norm(cand["name"]) in hay and norm(cand["class"]) in hay:
        return ok()
    return bad("name and class not found in the Notes and Handouts footer")

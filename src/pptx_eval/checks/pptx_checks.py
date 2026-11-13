"""Checks that read the .pptx itself. All values are effective values (slide > layout > master)."""
from __future__ import annotations

import colorsys
import difflib
import re

from ..images import background_luminance, contrast_ratio, relative_luminance, same_image
from ..inspect import Deck, Shape, Slide
from .base import EMU_PER_CM, Context, Outcome, bad, check, combine, ok, review


# ---------------------------------------------------------------- helpers
def norm(s: str) -> str:
    s = s.replace("–", "-").replace("—", "-").replace("’", "'")
    return re.sub(r"\s+", " ", s).strip().casefold()


def pick(deck: Deck, sel) -> list[Slide]:
    """`sel`: "all" | [1,2] | {"all_except": [1]} | {"only": [2,4]}."""
    if sel in (None, "all"):
        return list(deck.slides)
    if isinstance(sel, dict):
        if "all_except" in sel:
            return [s for s in deck.slides if s.index not in sel["all_except"]]
        sel = sel.get("only", [])
    return [s for s in deck.slides if s.index in sel]


def slide_at(deck: Deck, n: int) -> Slide | None:
    return deck.slides[n - 1] if 0 < n <= len(deck.slides) else None


def img_matches(ctx: Context, shape: Shape, name: str) -> bool | None:
    """Does this picture show the named reference image?

    With the reference image supplied, only the picture's content counts (alt text/file name can lie).
    Without it, fall back to the file name PowerPoint keeps in the alt text; None = cannot decide.
    """
    if shape.image is None:
        return False
    ref = ctx.assets.get(name)
    if ref:
        return same_image(shape.image, ref[0])
    return True if name.casefold() in norm(f"{shape.descr} {shape.name}") else None


def body_shapes(slide: Slide) -> list[Shape]:
    return [s for s in slide.shapes if s.kind == "sp" and s.has_text and s.ph_type not in ("title", "ctrTitle", "dt", "ftr", "sldNum")]


def links_for_text(slide: Slide, text: str):
    t = norm(text)
    for sh in slide.shapes:
        for l in sh.links:
            if (l.scope == "run" and t in norm(l.text)) or (l.scope == "shape" and t in norm(sh.text)):
                yield l


def hsv(hexcol: str) -> tuple[float, float, float]:
    r, g, b = (int(hexcol[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    return h * 360, s, v


def lum_of(hexcol: str) -> float:
    return relative_luminance(*(int(hexcol[i:i + 2], 16) for i in (0, 2, 4)))


# ---------------------------------------------------------------- master / house style
@check("deck.slide_ratio")
def slide_ratio(ctx: Context, p: dict) -> Outcome:
    want, tol = p.get("ratio", 1.6), p.get("tolerance", 0.01)
    got = ctx.deck.ratio
    if abs(got - want) <= tol:
        return ok(f"slide size ratio {got:.3f}")
    return bad(f"slide size ratio is {got:.3f} ({ctx.deck.slide_cx}x{ctx.deck.slide_cy} EMU), expected {want}")


@check("master.background_image")
def background_image(ctx: Context, p: dict) -> Outcome:
    lo, hi = p.get("transparency_pct", [55, 65])
    problems = []
    for s in pick(ctx.deck, p.get("slides")):
        bg = s.background
        if bg.image is None:
            problems.append(f"slide {s.index}: no background image")
            continue
        ref = ctx.assets.get(p["asset"])
        if ref is None:
            return review(f"reference image {p['asset']} not supplied")
        if not same_image(bg.image, ref[0]):
            problems.append(f"slide {s.index}: background is not {p['asset']}")
        tr = 100 - (bg.opacity_pct if bg.opacity_pct is not None else 100)
        if not lo <= tr <= hi:
            problems.append(f"slide {s.index}: transparency {tr:g}% (want {lo}-{hi}%)")
    return bad("; ".join(problems)) if problems else ok("background image at accepted transparency on every slide")


@check("master.title_style")
def title_style(ctx: Context, p: dict) -> Outcome:
    problems = []
    min_contrast = p.get("min_contrast", 4.5)
    for s in pick(ctx.deck, p.get("slides")):
        runs = [r for para in (s.title.paragraphs if s.title else []) for r in para.runs if r.text.strip()]
        if not runs:
            problems.append(f"slide {s.index}: no title")
            continue
        for key in ("bold", "italic"):
            if p.get(key) and not all(getattr(r.props, key[0]) for r in runs):
                problems.append(f"slide {s.index}: title not {key}")
        col = runs[0].props.color
        if p.get("dark_blue"):
            if not col:
                problems.append(f"slide {s.index}: title colour not set")
                continue
            h, sat, v = hsv(col)
            if not (180 <= h <= 260 and sat >= 0.25 and v <= 0.6):
                problems.append(f"slide {s.index}: title colour #{col} is not dark blue")
            bgimg = s.background.image_bytes
            if bgimg:
                opacity = (s.background.opacity_pct if s.background.opacity_pct is not None else 100) / 100
                bl = background_luminance(bgimg, opacity)
                if bl is not None and (cr := contrast_ratio(lum_of(col), bl)) < min_contrast:
                    problems.append(f"slide {s.index}: title contrast {cr:.1f}:1 is below {min_contrast}:1")
    return bad("; ".join(problems)) if problems else ok("all titles meet the style")


@check("master.logo")
def logo(ctx: Context, p: dict) -> Outcome:
    problems = []
    W, H = ctx.deck.slide_cx, ctx.deck.slide_cy
    for s in pick(ctx.deck, p.get("slides")):
        found = False
        for src, pic in ctx.deck.visible_pictures(s):
            if img_matches(ctx, pic, p["asset"]) and pic.box():
                x0, y0, x1, y1 = pic.box()
                if (x0 + x1) / 2 > 0.6 * W and (y0 + y1) / 2 < 0.3 * H:
                    found = True
        if not found:
            problems.append(f"slide {s.index}: {p['asset']} not visible in the top right")
    return bad("; ".join(problems)) if problems else ok("logo top right on every slide")


def _text_of(sh: Shape) -> str:
    return norm(sh.text)


@check("master.footer_text")
def footer_text(ctx: Context, p: dict) -> Outcome:
    """Footer text centred, and an auto-updating date where the date is expected."""
    want, W = norm(p["text"]), ctx.deck.slide_cx
    problems = []
    for s in pick(ctx.deck, p.get("slides", {"all_except": [1]})):
        f = next((x for x in s.shapes if x.ph_type == "ftr"), None)
        if not f or _text_of(f) != want:
            problems.append(f"slide {s.index}: footer text missing or wrong")
        elif f.box() and abs((f.box()[0] + f.box()[2]) / 2 - W / 2) > 0.1 * W:
            problems.append(f"slide {s.index}: footer not centred")
    for n in p.get("date_slides", []):
        s = slide_at(ctx.deck, n)
        d = next((x for x in s.shapes if x.ph_type == "dt"), None) if s else None
        if not d:
            problems.append(f"slide {n}: no date")
        elif not any(t.startswith("datetime") for t in d.fields):
            problems.append(f"slide {n}: date is typed text, not an automatically updating date")
    return bad("; ".join(problems)) if problems else ok("centred footer text and auto date present")


@check("master.slide_number")
def slide_number(ctx: Context, p: dict) -> Outcome:
    W, problems = ctx.deck.slide_cx, []
    for s in pick(ctx.deck, p.get("slides", {"all_except": [1]})):
        n = next((x for x in s.shapes if x.ph_type == "sldNum"), None)
        if not n or "slidenum" not in n.fields:
            problems.append(f"slide {s.index}: no slide number")
        elif n.box() and (n.box()[0] + n.box()[2]) / 2 < 0.6 * W:
            problems.append(f"slide {s.index}: slide number not on the right")
    for n in p.get("absent_on", [1]):
        s = slide_at(ctx.deck, n)
        extra = [x.ph_type for x in s.shapes if x.ph_type in ("ftr", "sldNum", "dt")] if s else []
        if extra:
            problems.append(f"slide {n}: should have no footer/number/date, found {extra}")
    return bad("; ".join(problems)) if problems else ok("slide numbers right; title slide clean")


@check("layout.placeholder_hidden")
def placeholder_hidden(ctx: Context, p: dict) -> Outcome:
    lay = ctx.deck.layout_named(p["layout"])
    if not lay:
        return bad(f"no layout called {p['layout']}")
    problems = []
    if any(s.ph_type == p["ph_type"] for s in lay.shapes):
        problems.append(f"{p['ph_type']} placeholder is still on the {p['layout']} layout")
    for s in ctx.deck.slides:
        if s.layout.name.lower() == p["layout"].lower() and any(x.ph_type == p["ph_type"] for x in s.shapes):
            problems.append(f"slide {s.index} ({p['layout']}) still shows it")
    for n in p.get("shown_on_slides", []):
        s = slide_at(ctx.deck, n)
        if not s or not any(x.ph_type == p["ph_type"] for x in s.shapes):
            problems.append(f"slide {n} should still show it")
    return bad("; ".join(problems)) if problems else ok()


@check("layout.placeholder_added")
def placeholder_added(ctx: Context, p: dict) -> Outcome:
    lay = ctx.deck.layout_named(p["layout"])
    if not lay or not any(s.ph_type == p["ph_type"] for s in lay.shapes):
        return bad(f"no {p['ph_type']} placeholder on the {p['layout']} layout")
    sl = slide_at(ctx.deck, p["slide"])
    if not sl or not any(x.ph_type == p["ph_type"] and x.image for x in sl.shapes):
        return bad(f"slide {p['slide']} image is not in the {p['ph_type']} placeholder")
    return ok()


# ---------------------------------------------------------------- slides
@check("slides.layouts")
def slides_layouts(ctx: Context, p: dict) -> Outcome:
    want = p["layouts"]
    if len(ctx.deck.slides) != len(want):
        return bad(f"{len(ctx.deck.slides)} slides, expected {len(want)}")
    problems = []
    for s, name in zip(ctx.deck.slides, want):
        if s.layout.name.lower() != name.lower():
            problems.append(f"slide {s.index}: layout {s.layout.name!r}, expected {name!r}")
        if not s.title or not s.title.is_ph:
            problems.append(f"slide {s.index}: title is not a title placeholder")
        if norm(s.title.text if s.title else "") != norm(p.get("titles", {}).get(str(s.index), s.title.text if s.title else "")):
            problems.append(f"slide {s.index}: wrong title")
    return bad("; ".join(problems)) if problems else ok()


@check("slides.short_bullets")
def short_bullets(ctx: Context, p: dict) -> Outcome:
    mx, problems = p.get("max_words", 12), []
    for n in p["slides"]:
        s = slide_at(ctx.deck, n)
        paras = [para.text for sh in (body_shapes(s) if s else []) for para in sh.paragraphs if para.text.strip()]
        if not paras:
            problems.append(f"slide {n}: no body text")
        for t in paras:
            if len(t.split()) > mx:
                problems.append(f"slide {n}: {len(t.split())}-word paragraph (max {mx}): {t[:40]}...")
            if any(f.casefold() in t.casefold() for f in p.get("forbidden", [])):
                problems.append(f"slide {n}: contains text that should not be used: {t[:40]}")
    return bad("; ".join(problems)) if problems else ok()


@check("slides.images")
def slides_images(ctx: Context, p: dict) -> Outcome:
    outs = []
    for r in p.get("rules", []):
        s = slide_at(ctx.deck, r["slide"])
        pics = [x for x in (s.shapes if s else []) if x.kind == "pic"]
        verdicts = [img_matches(ctx, x, r["asset"]) for x in pics]
        if any(v is True for v in verdicts):
            hit = next(x for x, v in zip(pics, verdicts) if v)
            outs.append(bad(f"slide {r['slide']}: {r['asset']} not in a placeholder") if r.get("require_placeholder") and not hit.is_ph else ok())
        elif any(v is None for v in verdicts):
            outs.append(review(f"slide {r['slide']}: cannot confirm {r['asset']} (reference image not supplied)"))
        else:
            outs.append(bad(f"slide {r['slide']}: {r['asset']} image not found"))
    rel = p.get("relevant_on")
    if rel:
        s = slide_at(ctx.deck, rel["slide"])
        pics = [x for x in (s.shapes if s else []) if x.kind == "pic" and (x.is_ph or not rel.get("require_placeholder"))]
        if not pics:
            outs.append(bad(f"slide {rel['slide']}: no image in the picture placeholder"))
        else:
            res = [img_matches(ctx, x, a) for x in pics for a in rel["assets"]]
            outs.append(ok() if any(res) else review(f"slide {rel['slide']}: cannot confirm the image is one of {rel['assets']}") if None in res else bad(f"slide {rel['slide']}: image is not one of the aquarium images"))
    for name in p.get("forbidden_assets", []):
        every = [x for sl in ctx.deck.slides for x in sl.shapes if x.kind == "pic"] + [x for l in ctx.deck.layouts for x in l.shapes if x.kind == "pic"]
        if any(img_matches(ctx, x, name) for x in every):
            outs.append(bad(f"{name} image is used"))
    return combine(outs)


@check("slides.image_replaces_text")
def image_replaces_text(ctx: Context, p: dict) -> Outcome:
    s = slide_at(ctx.deck, p["slide"])
    if not s:
        return bad("slide missing")
    pics = [x for x in s.shapes if x.kind == "pic"]
    v = [img_matches(ctx, x, p["asset"]) for x in pics]
    words = sum(len(para.text.split()) for sh in body_shapes(s) for para in sh.paragraphs)
    if words > p.get("max_words", 25):
        return bad(f"{words} words of text still on the slide")
    if any(x is True for x in v):
        return ok()
    return review(f"cannot confirm {p['asset']} (reference image not supplied)") if None in v else bad(f"{p['asset']} not on slide {p['slide']}")


@check("slides.image_size_cm")
def image_size_cm(ctx: Context, p: dict) -> Outcome:
    tol, problems = p.get("tolerance_cm", 0.1), []
    for it in p["items"]:
        s = slide_at(ctx.deck, it["slide"])
        pic = next((x for x in (s.shapes if s else []) if x.kind == "pic" and img_matches(ctx, x, it["asset"])), None)
        if pic is None:
            problems.append(f"slide {it['slide']}: {it['asset']} not found")
        elif pic.cx is None or abs(pic.cx / EMU_PER_CM - it["width_cm"]) > tol or abs(pic.cy / EMU_PER_CM - it["height_cm"]) > tol:
            problems.append(f"slide {it['slide']}: {it['asset']} is {pic.cx / EMU_PER_CM:.2f} x {pic.cy / EMU_PER_CM:.2f} cm")
    return bad("; ".join(problems)) if problems else ok()


def _overlap_ratio(a, b) -> float:
    w, h = min(a[2], b[2]) - max(a[0], b[0]), min(a[3], b[3]) - max(a[1], b[1])
    if w <= 0 or h <= 0:
        return 0.0
    small = min((a[2] - a[0]) * (a[3] - a[1]), (b[2] - b[0]) * (b[3] - b[1])) or 1
    return w * h / small


@check("slides.fit_for_purpose")
def fit_for_purpose(ctx: Context, p: dict) -> Outcome:
    """Geometry/size checks only. The visual 'looks good' judgement stays with a human or vision judge."""
    min_pt, max_dist = p.get("min_font_pt", 18), p.get("max_distortion", 0.03)
    W, H = ctx.deck.slide_cx, ctx.deck.slide_cy
    hard, soft = [], []
    for s in ctx.deck.slides:
        content = [x for x in s.shapes if x.ph_type not in ("dt", "ftr", "sldNum") and x.box()]
        deco = [("logo", x) for src, x in ctx.deck.visible_pictures(s) if src != "slide" and x.box()]
        for sh in content:
            x0, y0, x1, y1 = sh.box()
            if x0 < 0 or y0 < 0 or x1 > W or y1 > H:
                hard.append(f"slide {s.index}: {sh.name} runs off the slide")
            for para in sh.paragraphs:
                for r in para.runs:
                    if r.text.strip() and r.props.sz and r.props.sz < min_pt:
                        hard.append(f"slide {s.index}: {r.props.sz:g}pt text in {sh.name}")
                        break
            if sh.kind == "pic" and sh.image and sh.image.width and sh.cy:
                l, t, rr, b = sh.crop
                native = (sh.image.width * (1 - l - rr)) / (sh.image.height * (1 - t - b))
                if abs(sh.cx / sh.cy / native - 1) > max_dist:
                    hard.append(f"slide {s.index}: {sh.name} is stretched ({abs(sh.cx / sh.cy / native - 1):.0%})")
        pairs = [(a, b) for i, a in enumerate(content) for b in content[i + 1:]] + [(a, b) for a in content for _, b in deco]
        for a, b in pairs:
            o = _overlap_ratio(a.box(), b.box())
            if o > 0.25:
                hard.append(f"slide {s.index}: {a.name} overlaps {b.name} ({o:.0%})")
            elif o > 0.02:
                soft.append(f"slide {s.index}: {a.name} slightly overlaps {b.name} ({o:.0%})")
    if hard:
        return bad("; ".join(hard))
    return review("measurements pass; " + ("; ".join(soft) + "; " if soft else "") + "text overflow and overall look need a visual check")


# ---------------------------------------------------------------- animations / transitions
def _targets(s: Slide):
    out = []
    for e in s.effects:
        sh = s.shape_by_id(e.shape_id)
        text = (sh.paragraphs[e.para].text if sh and e.para is not None and e.para < len(sh.paragraphs) else sh.text if sh else "")
        out.append((e, sh, text))
    return out


@check("anim.same_entrance")
def same_entrance(ctx: Context, p: dict) -> Outcome:
    s = slide_at(ctx.deck, p["slide"])
    if not s or not s.effects:
        return bad("no animations on the slide")
    t, problems = _targets(s), []
    texts = [norm(x) for x in p["texts"]]
    for e, sh, text in t:
        if sh is None or sh.kind == "pic" or sh.ph_type in ("title", "ctrTitle") or norm(text) not in texts:
            problems.append(f"animates {sh.name if sh else 'unknown'} ({text[:20]!r}), which should not be animated")
        if e.preset_class != p.get("preset_class", "entr"):
            problems.append("an effect is not an entrance effect")
    if len({(e.preset_id, e.preset_class) for e, _, _ in t}) > 1:
        problems.append("different effects used")
    if {norm(x) for _, _, x in t if norm(x) in texts} != set(texts):
        problems.append("not every item is animated")
    return bad("; ".join(sorted(set(problems)))) if problems else ok()


@check("anim.after_previous_in_order")
def after_previous(ctx: Context, p: dict) -> Outcome:
    s = slide_at(ctx.deck, p["slide"])
    t = _targets(s) if s else []
    if not t:
        return bad("no animations")
    got = [norm(x) for _, _, x in t]
    problems = []
    if got != [norm(x) for x in p["texts"]]:
        problems.append(f"order is {got}")
    bad_nodes = [e.node_type for e, _, _ in t if e.node_type != "afterEffect"]
    if bad_nodes:
        problems.append(f"start is {sorted(set(bad_nodes))}, not After Previous")
    return bad("; ".join(problems)) if problems else ok()


@check("anim.duration")
def anim_duration(ctx: Context, p: dict) -> Outcome:
    s = slide_at(ctx.deck, p["slide"])
    if not s or not s.effects:
        return bad("no animations")
    off = [e.duration_ms for e in s.effects if abs(e.duration_ms - p["duration_ms"]) > p.get("tolerance_ms", 50)]
    return bad(f"durations {off} ms, expected {p['duration_ms']}") if off else ok()


@check("transition.auto_advance")
def auto_advance(ctx: Context, p: dict) -> Outcome:
    s = slide_at(ctx.deck, p["slide"])
    t = s.transition if s else None
    if t is None:
        return bad("no transition")
    problems = []
    if t.advance_after_ms is None or abs(t.advance_after_ms - p["after_ms"]) > 50:
        problems.append(f"advance after {t.advance_after_ms} ms, expected {p['after_ms']}")
    if t.advance_on_click:
        problems.append("On Mouse Click is still ticked")
    return bad("; ".join(problems)) if problems else ok()


# ---------------------------------------------------------------- links
@check("links.slide_links")
def slide_links(ctx: Context, p: dict) -> Outcome:
    s = slide_at(ctx.deck, p["slide"])
    problems = []
    for it in p["items"]:
        ls = list(links_for_text(s, it["text"])) if s else []
        if not any(l.kind == "slide" and l.target_slide == it["target_slide"] for l in ls):
            got = [l.target_slide for l in ls if l.kind == "slide"]
            problems.append(f"{it['text']!r} does not link to slide {it['target_slide']} (links to {got or 'nothing'})")
    return bad("; ".join(problems)) if problems else ok()


@check("links.home_button")
def home_button(ctx: Context, p: dict) -> Outcome:
    prst, problems, boxes = p.get("shape_prst", "actionButtonHome"), [], []
    for n in p["slides"]:
        s = slide_at(ctx.deck, n)
        hit = next((x for x in (s.shapes if s else []) if x.prst == prst and any(l.kind == "slide" and l.target_slide == p["target_slide"] for l in x.links)), None)
        if hit is None:
            problems.append(f"slide {n}: no Home action button linked to slide {p['target_slide']}")
        else:
            boxes.append((n, (hit.x, hit.y, hit.cx, hit.cy)))
    if p.get("same_position", True) and len({b for _, b in boxes}) > 1:
        problems.append("Home button is not in the same position on every slide")
    return bad("; ".join(problems)) if problems else ok()


@check("links.external")
def external(ctx: Context, p: dict) -> Outcome:
    s = slide_at(ctx.deck, p["slide"])
    got = [l.target.rstrip("/").casefold() for sh in (s.shapes if s else []) for l in sh.links if l.target]
    missing = [u for u in p["targets"] if u.rstrip("/").casefold() not in got]
    return bad(f"missing or misspelt link(s): {missing}") if missing else ok()


@check("links.file")
def file_link(ctx: Context, p: dict) -> Outcome:
    s = slide_at(ctx.deck, p["slide"])
    ls = list(links_for_text(s, p["text"])) if s else []
    if any(l.kind == "file" and norm(l.target or "").rsplit("/", 1)[-1].rsplit("\\", 1)[-1].startswith(norm(p["file_stem"])) for l in ls):
        return ok()
    return bad(f"{p['text']!r} is not linked to the file {p['file_stem']} (links: {[(l.kind, l.target) for l in ls]})")


# ---------------------------------------------------------------- notes
@check("notes.present")
def notes_present(ctx: Context, p: dict) -> Outcome:
    s = slide_at(ctx.deck, p["slide"])
    n = s.notes if s else ""
    if len(n.split()) < p.get("min_words", 5):
        return bad("no speaker notes" if not n else "speaker notes too short")
    ref = p.get("reference_text")
    if ref:
        r = difflib.SequenceMatcher(None, norm(n), norm(ref)).ratio()
        if r < p.get("min_similarity", 0.8):
            return review(f"notes differ from the supplied text (similarity {r:.0%})")
    return ok()

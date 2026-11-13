"""Read a .pptx into plain facts. No PowerPoint needed; only zip + lxml. Nothing in the file is ever executed.

Every value is resolved the way PowerPoint resolves it (slide > layout > master), so checks can ask
"what does the pupil's slide actually show?" and not "where did they click?".
"""
from __future__ import annotations

import posixpath
import re
import zipfile
from dataclasses import dataclass, field
from io import BytesIO

from lxml import etree

from .images import ImageId, identify

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
EMU_PER_CM = 360_000
MAX_UNCOMPRESSED = 500 * 1024 * 1024
MAX_MEMBERS = 5000
_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False)


class BadPackage(Exception):
    """The upload is not a usable .pptx."""


def q(tag: str) -> str:
    pfx, name = tag.split(":")
    return f"{{{NS[pfx]}}}{name}"


# ------------------------------------------------------------------ package
@dataclass
class Rel:
    id: str
    type: str
    target: str  # absolute part path inside the zip, or the raw external target
    external: bool


class Package:
    def __init__(self, data: bytes):
        try:
            self.zf = zipfile.ZipFile(BytesIO(data))
        except zipfile.BadZipFile as e:
            raise BadPackage("not a zip/pptx file") from e
        infos = self.zf.infolist()
        if len(infos) > MAX_MEMBERS or sum(i.file_size for i in infos) > MAX_UNCOMPRESSED:
            raise BadPackage("package too large")
        self.names = {i.filename for i in infos}
        if "ppt/presentation.xml" not in self.names:
            raise BadPackage("no ppt/presentation.xml: not a PowerPoint presentation (.ppt/.key/.odp are not supported)")

    def read(self, path: str) -> bytes:
        return self.zf.read(path)

    def xml(self, path: str) -> etree._Element:
        return etree.fromstring(self.read(path), _PARSER)

    def rels(self, part: str) -> dict[str, Rel]:
        d, f = posixpath.split(part)
        rp = posixpath.join(d, "_rels", f + ".rels")
        out: dict[str, Rel] = {}
        if rp not in self.names:
            return out
        for r in self.xml(rp):
            ext = r.get("TargetMode") == "External"
            tgt = r.get("Target", "")
            if not ext:
                tgt = posixpath.normpath(posixpath.join(d, tgt)) if not tgt.startswith("/") else tgt.lstrip("/")
            out[r.get("Id")] = Rel(r.get("Id"), r.get("Type", ""), tgt, ext)
        return out


# ------------------------------------------------------------------ model
@dataclass
class Props:
    sz: float | None = None  # points
    b: bool | None = None
    i: bool | None = None
    color: str | None = None  # RRGGBB


@dataclass
class Run:
    text: str
    props: Props  # effective (resolved) properties


@dataclass
class Paragraph:
    text: str
    level: int
    runs: list[Run]


@dataclass
class Link:
    scope: str  # "shape" | "run"
    text: str  # shape text or run text the link sits on
    kind: str  # slide | url | mailto | file | show_jump | other
    action: str | None
    target_slide: int | None = None
    target: str | None = None  # url / mailto / file path / jump name


@dataclass
class Shape:
    id: int
    name: str
    descr: str
    kind: str  # sp | pic | other
    ph_type: str | None  # None when not a placeholder
    ph_idx: int | None
    prst: str | None
    x: int | None
    y: int | None
    cx: int | None
    cy: int | None
    paragraphs: list[Paragraph] = field(default_factory=list)
    links: list[Link] = field(default_factory=list)
    fields: list[str] = field(default_factory=list)  # a:fld types inside the shape
    image: ImageId | None = None
    crop: tuple[float, float, float, float] = (0, 0, 0, 0)  # l,t,r,b as fractions
    in_group: bool = False

    @property
    def is_ph(self) -> bool:
        return self.ph_type is not None

    @property
    def text(self) -> str:
        return "\n".join(p.text for p in self.paragraphs)

    @property
    def has_text(self) -> bool:
        return any(p.text.strip() for p in self.paragraphs)

    def box(self) -> tuple[int, int, int, int] | None:
        if None in (self.x, self.y, self.cx, self.cy):
            return None
        return (self.x, self.y, self.x + self.cx, self.y + self.cy)


@dataclass
class Background:
    image: ImageId | None = None
    image_bytes: bytes | None = None
    opacity_pct: float | None = None  # alphaModFix; 100 = opaque. transparency = 100 - opacity
    color: str | None = None
    source: str = ""  # slide | layout | master


@dataclass
class Effect:
    order: int
    shape_id: int | None
    para: int | None
    preset_id: int | None
    preset_class: str | None  # entr | exit | emph | path
    node_type: str | None  # clickEffect | withEffect | afterEffect
    duration_ms: int
    delay_ms: int


@dataclass
class Transition:
    advance_on_click: bool
    advance_after_ms: int | None
    effect: str | None


@dataclass
class Layout:
    index: int
    name: str
    shapes: list[Shape]
    show_master_shapes: bool
    background: Background | None


@dataclass
class Slide:
    index: int  # 1-based
    layout: Layout
    shapes: list[Shape]
    show_master_shapes: bool
    background: Background  # effective
    effects: list[Effect]
    transition: Transition | None
    notes: str
    title: Shape | None

    def shape_by_id(self, sid: int | None) -> Shape | None:
        return next((s for s in self.shapes if s.id == sid), None)


@dataclass
class Deck:
    slide_cx: int
    slide_cy: int
    slides: list[Slide]
    layouts: list[Layout]
    master_shapes: list[Shape]
    notes_handout_text: list[str]  # header/footer text on notes master, handout master, notes slides

    @property
    def ratio(self) -> float:
        return self.slide_cx / self.slide_cy

    def layout_named(self, name: str) -> Layout | None:
        return next((l for l in self.layouts if l.name.lower() == name.lower()), None)

    def visible_pictures(self, slide: Slide) -> list[tuple[str, Shape]]:
        """Pictures drawn on `slide`, with where they come from (slide/layout/master)."""
        out = [("slide", s) for s in slide.shapes if s.kind == "pic"]
        if slide.show_master_shapes:
            out += [("layout", s) for s in slide.layout.shapes if s.kind == "pic" and not s.is_ph]
            if slide.layout.show_master_shapes:
                out += [("master", s) for s in self.master_shapes if s.kind == "pic" and not s.is_ph]
        return out


# ------------------------------------------------------------------ helpers
def _int(v: str | None) -> int | None:
    try:
        return int(v) if v is not None else None
    except ValueError:
        return None


def _theme_colors(pkg: Package, master_path: str) -> dict[str, str]:
    cmap = {}
    master = pkg.xml(master_path)
    cm = master.find("p:clrMap", NS)
    if cm is not None:
        cmap = dict(cm.attrib)
    theme_rel = next((r for r in pkg.rels(master_path).values() if r.type.endswith("/theme")), None)
    scheme: dict[str, str] = {}
    if theme_rel:
        th = pkg.xml(theme_rel.target)
        cs = th.find(".//a:clrScheme", NS)
        if cs is not None:
            for el in cs:
                name = etree.QName(el).localname
                c = el[0] if len(el) else None
                if c is None:
                    continue
                scheme[name] = (c.get("val") if etree.QName(c).localname == "srgbClr" else c.get("lastClr")) or ""
    out = dict(scheme)
    for k, v in cmap.items():  # bg1->lt1, tx1->dk1 ...
        if v in scheme:
            out[k] = scheme[v]
    return out


def _color(el: etree._Element | None, theme: dict[str, str]) -> str | None:
    if el is None:
        return None
    fill = el.find("a:solidFill", NS)
    if fill is None or not len(fill):
        return None
    c = fill[0]
    kind = etree.QName(c).localname
    if kind == "srgbClr":
        return (c.get("val") or "").upper() or None
    if kind == "schemeClr":
        return (theme.get(c.get("val", ""), "") or "").upper() or None
    if kind == "sysClr":
        return (c.get("lastClr") or "").upper() or None
    return None


def _props(el: etree._Element | None, theme: dict[str, str]) -> Props:
    if el is None:
        return Props()
    sz = _int(el.get("sz"))
    return Props(
        sz=sz / 100 if sz else None,
        b=None if el.get("b") is None else el.get("b") in ("1", "true"),
        i=None if el.get("i") is None else el.get("i") in ("1", "true"),
        color=_color(el, theme),
    )


def _merge(*layers: Props) -> Props:
    """First non-None wins (layers ordered most- to least-specific)."""
    out = Props()
    for f in ("sz", "b", "i", "color"):
        for l in layers:
            if getattr(l, f) is not None:
                setattr(out, f, getattr(l, f))
                break
    return out


def _lvl_props(lst: etree._Element | None, lvl: int, theme) -> Props:
    if lst is None:
        return Props()
    p = lst.find(f"a:lvl{lvl + 1}pPr", NS)
    return _props(p.find("a:defRPr", NS) if p is not None else None, theme)


def _ph_of(sp: etree._Element) -> tuple[str | None, int | None]:
    ph = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
    if ph is None:
        ph = sp.find("p:nvPicPr/p:nvPr/p:ph", NS)
    if ph is None:
        return None, None
    return ph.get("type", "obj"), _int(ph.get("idx"))


def _find_ph(shapes: list[Shape], ph_type: str | None, ph_idx: int | None) -> Shape | None:
    if ph_type is None:
        return None
    fixed = {"title", "ctrTitle", "dt", "ftr", "sldNum", "subTitle"}
    group = {"title": "title", "ctrTitle": "title"}
    if ph_type not in fixed and ph_idx is not None:
        for s in shapes:
            if s.ph_idx == ph_idx and s.is_ph:
                return s
    for s in shapes:
        if s.is_ph and group.get(s.ph_type, s.ph_type) == group.get(ph_type, ph_type):
            return s
    return None


# ------------------------------------------------------------------ shape parsing
class _Ctx:
    def __init__(self, pkg: Package, part: str, theme, slide_index: dict[str, int]):
        self.pkg, self.part, self.theme, self.slide_index = pkg, part, theme, slide_index
        self.rels = pkg.rels(part)
        self.image_cache: dict[str, tuple[ImageId, bytes]] = {}

    def image(self, rid: str | None) -> tuple[ImageId, bytes] | None:
        rel = self.rels.get(rid or "")
        if not rel or rel.external or rel.target not in self.pkg.names:
            return None
        if rel.target not in self.image_cache:
            data = self.pkg.read(rel.target)
            self.image_cache[rel.target] = (identify(data), data)
        return self.image_cache[rel.target]

    def link(self, h: etree._Element, scope: str, text: str) -> Link:
        action = h.get("action")
        rel = self.rels.get(h.get(q("r:id")) or "")
        if rel and rel.type.endswith("/slide") and not rel.external:
            return Link(scope, text, "slide", action, target_slide=self.slide_index.get(rel.target))
        if rel and rel.external:
            t = rel.target
            kind = "mailto" if t.lower().startswith("mailto:") else "url" if re.match(r"^[a-z][a-z0-9+.-]*://", t, re.I) else "file"
            return Link(scope, text, kind, action, target=t)
        if action and "jump=" in action:
            return Link(scope, text, "show_jump", action, target=action.split("jump=")[1])
        return Link(scope, text, "other", action)


def _parse_paragraphs(tx: etree._Element | None, shape_lst, layer_lsts, base: Props, theme) -> list[Paragraph]:
    """Resolve each run's properties: run > paragraph defaults > shape lstStyle > layout ph > master ph > style group."""
    out: list[Paragraph] = []
    if tx is None:
        return out
    for p in tx.findall("a:p", NS):
        ppr = p.find("a:pPr", NS)
        lvl = _int(ppr.get("lvl")) if ppr is not None else 0
        lvl = lvl or 0
        inherited = _merge(_lvl_props(shape_lst, lvl, theme), *[_lvl_props(l, lvl, theme) for l in layer_lsts], base)
        runs, parts = [], []
        for el in p:
            tag = etree.QName(el).localname
            if tag in ("r", "fld"):
                t = "".join(el.xpath("a:t/text()", namespaces=NS))
                runs.append(Run(t, _merge(_props(el.find("a:rPr", NS), theme), inherited)))
                parts.append(t)
            elif tag == "br":
                parts.append(" ")
        out.append(Paragraph("".join(parts), lvl, runs))
    return out


def _parse_shape(el, ctx: _Ctx, layer_ph: list[Shape], master_style: dict[str, Props], in_group=False) -> Shape | None:
    tag = etree.QName(el).localname
    if tag not in ("sp", "pic", "cxnSp", "graphicFrame"):
        return None
    c = el.find("*/p:cNvPr", NS)
    if c is None:
        return None
    ph_type, ph_idx = _ph_of(el)
    xf = el.find("p:spPr/a:xfrm", NS)
    if xf is None:
        xf = el.find("p:xfrm", NS)
    x = y = cx = cy = None
    if xf is not None and xf.find("a:off", NS) is not None and xf.find("a:ext", NS) is not None:
        x, y = _int(xf.find("a:off", NS).get("x")), _int(xf.find("a:off", NS).get("y"))
        cx, cy = _int(xf.find("a:ext", NS).get("cx")), _int(xf.find("a:ext", NS).get("cy"))
    parent = _find_ph(layer_ph, ph_type, ph_idx)
    if x is None and parent is not None:
        x, y, cx, cy = parent.x, parent.y, parent.cx, parent.cy
    geom = el.find("p:spPr/a:prstGeom", NS)
    s = Shape(
        id=_int(c.get("id")) or 0, name=c.get("name", ""), descr=c.get("descr", ""),
        kind="pic" if tag == "pic" else "sp" if tag == "sp" else "other",
        ph_type=ph_type, ph_idx=ph_idx, prst=geom.get("prst") if geom is not None else None,
        x=x, y=y, cx=cx, cy=cy, in_group=in_group,
    )
    tx = el.find("p:txBody", NS)
    if tx is not None:
        key = {"title": "title", "ctrTitle": "title", "subTitle": "body", "body": "body", "obj": "body"}.get(ph_type or "", "other" if ph_type else "default")
        base = master_style.get(key, Props())
        shape_lst = tx.find("a:lstStyle", NS)
        layer_lsts = [l._lst for l in [parent] if l is not None and hasattr(l, "_lst")]
        s.paragraphs = _parse_paragraphs(tx, shape_lst, layer_lsts, base, ctx.theme)
        s._lst = shape_lst  # type: ignore[attr-defined]  (kept so children layers can inherit)
        s.fields = [f.get("type", "") for f in tx.iter(q("a:fld"))]
        fs = tx.find("a:bodyPr/a:normAutofit", NS)
        if fs is not None and fs.get("fontScale"):
            scale = (_int(fs.get("fontScale")) or 100000) / 100000
            for p in s.paragraphs:
                for r in p.runs:
                    if r.props.sz:
                        r.props.sz = round(r.props.sz * scale, 1)
        for h in tx.iter(q("a:hlinkClick")):
            run = h.getparent().getparent()
            s.links.append(ctx.link(h, "run", "".join(run.xpath("a:t/text()", namespaces=NS))))
    h = c.find("a:hlinkClick", NS)
    if h is not None:
        s.links.append(ctx.link(h, "shape", s.text))
    if tag == "pic":
        blip = el.find("p:blipFill/a:blip", NS)
        got = ctx.image(blip.get(q("r:embed")) if blip is not None else None)
        if got:
            s.image = got[0]
            s._bytes = got[1]  # type: ignore[attr-defined]
        sr = el.find("p:blipFill/a:srcRect", NS)
        if sr is not None:
            g = lambda k: (_int(sr.get(k)) or 0) / 100000
            s.crop = (g("l"), g("t"), g("r"), g("b"))
    return s


def _walk(tree, ctx, layer_ph, master_style, in_group=False) -> list[Shape]:
    out: list[Shape] = []
    for el in tree:
        if etree.QName(el).localname == "grpSp":
            out += _walk(el, ctx, layer_ph, master_style, True)
        else:
            s = _parse_shape(el, ctx, layer_ph, master_style, in_group)
            if s:
                out.append(s)
    return out


def _background(root, ctx: _Ctx, source: str) -> Background | None:
    bg = root.find("p:cSld/p:bg/p:bgPr", NS)
    if bg is None:
        return None
    out = Background(source=source)
    blip = bg.find("a:blipFill/a:blip", NS)
    if blip is not None:
        got = ctx.image(blip.get(q("r:embed")))
        if got:
            out.image, out.image_bytes = got
        am = blip.find("a:alphaModFix", NS)
        out.opacity_pct = (_int(am.get("amt")) or 100000) / 1000 if am is not None else 100.0
    out.color = _color(bg, ctx.theme)
    return out


def _effects(root, order_base=0) -> list[Effect]:
    out: list[Effect] = []
    for ctn in root.iter(q("p:cTn")):
        if ctn.get("presetClass") is None:
            continue
        tgt = ctn.find(".//p:spTgt", NS)
        pr = tgt.find("p:txEl/p:pRg", NS) if tgt is not None else None
        durs = [int(d) for d in ctn.xpath(".//p:cTn/@dur", namespaces=NS) if d.isdigit() and int(d) > 1]
        cond = ctn.find("p:stCondLst/p:cond", NS)
        delay = int(cond.get("delay")) if cond is not None and (cond.get("delay") or "").isdigit() else 0
        out.append(Effect(
            len(out), _int(tgt.get("spid")) if tgt is not None else None,
            _int(pr.get("st")) if pr is not None else None,
            _int(ctn.get("presetID")), ctn.get("presetClass"), ctn.get("nodeType"),
            max(durs) if durs else 0, delay,
        ))
    return out


def _transition(root) -> Transition | None:
    t = root.find(".//p:transition", NS)  # also inside mc:AlternateContent
    if t is None:
        return None
    after = t.get("advTm")
    eff = next((etree.QName(c).localname for c in t if etree.QName(c).namespace == NS["p"]), None)
    return Transition(t.get("advClick", "1") in ("1", "true"), int(after) if after and after.isdigit() else None, eff)


# ------------------------------------------------------------------ entry point
def inspect_pptx(data: bytes) -> Deck:
    pkg = Package(data)
    pres = pkg.xml("ppt/presentation.xml")
    prels = pkg.rels("ppt/presentation.xml")
    sz = pres.find("p:sldSz", NS)
    if sz is None:
        raise BadPackage("presentation has no slide size")
    slide_parts = [prels[s.get(q("r:id"))].target for s in pres.findall("p:sldIdLst/p:sldId", NS)]
    slide_index = {p: i + 1 for i, p in enumerate(slide_parts)}

    master_part = next(r.target for r in prels.values() if r.type.endswith("/slideMaster"))
    theme = _theme_colors(pkg, master_part)
    mroot = pkg.xml(master_part)
    mctx = _Ctx(pkg, master_part, theme, slide_index)

    # text style groups on the master: title / body / other / default
    def style(tag: str) -> Props:
        st = mroot.find(f"p:txStyles/p:{tag}", NS)
        return _lvl_props(st, 0, theme) if st is not None else Props()

    dts = pres.find("p:defaultTextStyle", NS)
    master_style = {"title": style("titleStyle"), "body": style("bodyStyle"), "other": style("otherStyle"),
                    "default": _merge(_lvl_props(dts, 0, theme), Props(sz=18.0))}
    master_shapes = _walk(mroot.find("p:cSld/p:spTree", NS), mctx, [], master_style)
    master_bg = _background(mroot, mctx, "master")

    layouts: list[Layout] = []
    layout_by_part: dict[str, Layout] = {}
    layout_roots = {}
    for i, rel in enumerate(r for r in pkg.rels(master_part).values() if r.type.endswith("/slideLayout")):
        lroot = pkg.xml(rel.target)
        lctx = _Ctx(pkg, rel.target, theme, slide_index)
        shapes = _walk(lroot.find("p:cSld/p:spTree", NS), lctx, master_shapes, master_style)
        lay = Layout(i + 1, lroot.find("p:cSld", NS).get("name", ""), shapes,
                     lroot.get("showMasterSp", "1") in ("1", "true"), _background(lroot, lctx, "layout"))
        layouts.append(lay)
        layout_by_part[rel.target] = lay
        layout_roots[rel.target] = lroot

    slides: list[Slide] = []
    for idx, part in enumerate(slide_parts, 1):
        root = pkg.xml(part)
        ctx = _Ctx(pkg, part, theme, slide_index)
        lrel = next(r for r in ctx.rels.values() if r.type.endswith("/slideLayout"))
        lay = layout_by_part[lrel.target]
        shapes = _walk(root.find("p:cSld/p:spTree", NS), ctx, lay.shapes, master_style)
        bg = _background(root, ctx, "slide") or lay.background or master_bg or Background(source="none")
        notes = ""
        nrel = next((r for r in ctx.rels.values() if r.type.endswith("/notesSlide")), None)
        if nrel:
            nroot = pkg.xml(nrel.target)
            for sp in nroot.iterfind(".//p:sp", NS):
                if _ph_of(sp)[0] == "body":
                    notes += "\n".join("".join(p.xpath(".//a:t/text()", namespaces=NS)) for p in sp.findall(".//a:p", NS))
        title = next((s for s in shapes if s.ph_type in ("title", "ctrTitle")), None)
        slides.append(Slide(idx, lay, shapes, root.get("showMasterSp", "1") in ("1", "true"), bg,
                            _effects(root), _transition(root), notes.strip(), title))

    nh: list[str] = []
    for name in sorted(pkg.names):
        if re.match(r"ppt/(notesMasters/notesMaster|handoutMasters/handoutMaster|notesSlides/notesSlide)\d+\.xml$", name):
            r = pkg.xml(name)
            for sp in r.iterfind(".//p:sp", NS):
                if _ph_of(sp)[0] in ("ftr", "hdr"):
                    t = "".join(sp.xpath(".//a:t/text()", namespaces=NS)).strip()
                    if t:
                        nh.append(t)
    return Deck(_int(sz.get("cx")) or 0, _int(sz.get("cy")) or 0, slides, layouts, master_shapes, nh)

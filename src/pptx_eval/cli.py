from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .engine import BadPackage, evaluate
from .spec import Spec

MARK = {"pass": "PASS  ", "fail": "FAIL  ", "review": "REVIEW"}


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="pptx-eval")
    sub = ap.add_subparsers(dest="cmd", required=True)
    ev = sub.add_parser("evaluate", help="check a .pptx against a spec")
    ev.add_argument("--spec", required=True, type=Path)
    ev.add_argument("--pptx", required=True, type=Path)
    ev.add_argument("--extra", action="append", default=[], metavar="NAME=PATH", help="HANDOUTS=x.pdf (repeatable)")
    ev.add_argument("--name"); ev.add_argument("--class", dest="cls")
    ev.add_argument("--json", action="store_true", help="print the full JSON report")
    a = ap.parse_args(argv)
    spec = Spec.model_validate_json(a.spec.read_text(encoding="utf8"))
    extras = {k: Path(v).read_bytes() for k, v in (e.split("=", 1) for e in a.extra)}
    try:
        rep = evaluate(spec, a.pptx.read_bytes(), extras, {"name": a.name, "class": a.cls}, a.spec.parent)
    except BadPackage as e:
        print(f"error: {e}", file=sys.stderr)
        return 2
    if a.json:
        print(rep.model_dump_json(indent=2))
        return 0
    for r in rep.results:
        print(f"{MARK[r.status]} {r.id:<4} {r.marks_awarded:g}/{r.marks_available:g}  {r.description}" + (f"\n         -> {r.detail}" if r.detail and r.status != "pass" else ""))
    print(f"\nconfirmed {rep.marks_awarded:g}/{rep.marks_available:g}; a further {rep.marks_pending_review:g} marks await review")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

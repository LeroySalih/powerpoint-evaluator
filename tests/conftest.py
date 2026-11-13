import io
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
MODEL = ROOT / "private" / "PEARL_KIOSK_MODEL_ANSWER.pptx"
SPEC = ROOT / "specs" / "pearl-kiosk.json"
ASSETS = ROOT / "specs" / "pearl-kiosk" / "assets"

needs_private = pytest.mark.skipif(not (MODEL.is_file() and ASSETS.is_dir()), reason="private model answer / reference images not present")


def mutate(src: bytes, changes: dict) -> bytes:
    """Rewrite a pptx, applying fn(bytes)->bytes to named members (or str replacements given as (old, new) tuples)."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(src)) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            fn = changes.get(info.filename)
            if fn is not None:
                data = fn(data) if callable(fn) else data.decode("utf8").replace(*fn).encode("utf8")
            zout.writestr(info, data)
    return out.getvalue()


def many(*pairs):
    """Several (old, new) replacements on one member."""
    def fn(b: bytes) -> bytes:
        s = b.decode("utf8")
        for old, new in pairs:
            assert old in s, f"mutation target not found: {old[:50]}"
            s = s.replace(old, new)
        return s.encode("utf8")
    return fn

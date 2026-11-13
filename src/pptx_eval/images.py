"""Image identity: match a picture in a deck to a reference image even after PowerPoint recompresses or resizes it."""
from __future__ import annotations

import hashlib
import io
from dataclasses import dataclass

from PIL import Image

Image.MAX_IMAGE_PIXELS = 50_000_000  # decompression-bomb guard


@dataclass(frozen=True)
class ImageId:
    sha256: str
    dhash: int | None  # 64-bit difference hash, None if the format cannot be decoded
    width: int | None
    height: int | None


def dhash_of(img: Image.Image) -> int:
    g = img.convert("L").resize((9, 8), Image.LANCZOS)
    px = list(g.tobytes())
    bits = 0
    for row in range(8):
        for col in range(8):
            bits = (bits << 1) | (px[row * 9 + col] > px[row * 9 + col + 1])
    return bits


def identify(data: bytes) -> ImageId:
    sha = hashlib.sha256(data).hexdigest()
    try:
        with Image.open(io.BytesIO(data)) as im:
            im.load()
            return ImageId(sha, dhash_of(im), im.width, im.height)
    except Exception:  # EMF/WMF/corrupt: exact-hash matching only
        return ImageId(sha, None, None, None)


def same_image(a: ImageId, b: ImageId, max_distance: int = 8) -> bool:
    if a.sha256 == b.sha256:
        return True
    if a.dhash is None or b.dhash is None:
        return False
    return bin(a.dhash ^ b.dhash).count("1") <= max_distance


def background_luminance(data: bytes, opacity: float) -> float | None:
    """Relative luminance of an image's mean colour after blending over white at `opacity` (0..1)."""
    try:
        with Image.open(io.BytesIO(data)) as im:
            r, g, b = im.convert("RGB").resize((1, 1), Image.BOX).getpixel((0, 0))
    except Exception:
        return None
    mix = [255 * (1 - opacity) + c * opacity for c in (r, g, b)]
    return relative_luminance(*mix)


def relative_luminance(r: float, g: float, b: float) -> float:
    def lin(c: float) -> float:
        c /= 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b)


def contrast_ratio(l1: float, l2: float) -> float:
    hi, lo = max(l1, l2), min(l1, l2)
    return (hi + 0.05) / (lo + 0.05)

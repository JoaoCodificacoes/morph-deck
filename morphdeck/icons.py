"""Icons and technology logos, rendered to cached PNGs.

* ``icon()`` — any of Google's Material Symbols (Rounded), in any colour, filled or outlined. Names are the
  ones on https://fonts.google.com/icons, e.g. ``"cloud_off"``, ``"schema"``, ``"lock"``.
* ``logo()`` — a technology logo from Simple Icons in its brand colour, inside a small white disc, for badges on
  diagram boxes. Logos are trademarks of their owners; use them to name the technology, nothing more.

Fonts and SVGs are downloaded once into ``~/.cache/morphdeck`` (override with ``MORPHDECK_CACHE``).
"""
import json
import os
import unicodedata
import urllib.request
from pathlib import Path

from .draw import box, picture, shadow
from .theme import current

CACHE = Path(os.environ.get("MORPHDECK_CACHE", Path.home() / ".cache" / "morphdeck"))
_MS = "https://github.com/google/material-design-icons/raw/master/variablefont/MaterialSymbolsRounded%5BFILL,GRAD,opsz,wght%5D"
SIMPLE_ICONS = "13.21.0"
_SI = f"https://cdn.jsdelivr.net/npm/simple-icons@{SIMPLE_ICONS}"


def _fetch(url, dest):
    dest.parent.mkdir(parents=True, exist_ok=True)
    if not dest.exists():
        tmp = dest.with_suffix(dest.suffix + ".part")
        urllib.request.urlretrieve(url, tmp)
        tmp.replace(dest)
    return dest


def icon_png(name, color, px=256, fill=1):
    """Path to a PNG of Material Symbol ``name`` in ``color`` (hex); ``fill=0`` for the outlined style."""
    from PIL import Image, ImageDraw, ImageFont

    out = CACHE / "icons" / f"{name}_{color}_{fill}.png"
    if out.exists():
        return out
    ttf = _fetch(_MS + ".ttf", CACHE / "MaterialSymbolsRounded.ttf")
    cps_file = _fetch(_MS + ".codepoints", CACHE / "MaterialSymbolsRounded.codepoints")
    cps = dict(l.split() for l in cps_file.read_text().splitlines() if l.strip())
    if name not in cps:
        raise KeyError(f"no Material Symbol named '{name}' (see https://fonts.google.com/icons)")
    font = ImageFont.truetype(str(ttf), int(px * 0.9))
    try:
        font.set_variation_by_axes([fill, 0, 48, 500])  # FILL, GRAD, opsz, wght
    except Exception:
        pass
    img = Image.new("RGBA", (px, px), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    ch = chr(int(cps[name], 16))
    l, t, r, b = d.textbbox((0, 0), ch, font=font)
    d.text(((px - (r - l)) / 2 - l, (px - (b - t)) / 2 - t), ch, font=font,
           fill=tuple(int(color[i:i + 2], 16) for i in (0, 2, 4)) + (255,))
    out.parent.mkdir(parents=True, exist_ok=True)
    img.save(out)
    return out


def icon(slide, name, x, y, size, color=None, fill=1, name_=None):
    """Material Symbol ``name`` as a ``size``-inch picture. ``name_`` names the shape (e.g. for Morph)."""
    return picture(slide, icon_png(name, color or current().palette.ink, fill=fill), x, y, w=size, h=size,
                   name=name_)


def _slug(title):
    s = title.lower().replace("+", "plus").replace(".", "dot").replace("&", "and")
    s = unicodedata.normalize("NFD", s)
    return "".join(c for c in s if c.isascii() and c.isalnum())


def brand_color(slug):
    data = json.loads(_fetch(f"{_SI}/_data/simple-icons.json", CACHE / f"simple-icons-{SIMPLE_ICONS}.json").read_text())
    for item in data.get("icons", data) if isinstance(data, dict) else data:
        if item.get("slug", _slug(item["title"])) == slug:
            return item["hex"]
    raise KeyError(f"no Simple Icons logo '{slug}' (see https://simpleicons.org)")


def logo_png(slug, color=None, px=256):
    """Path to a PNG of the Simple Icons logo ``slug`` (e.g. ``"postgresql"``), in its brand colour."""
    import cairosvg

    color = color or brand_color(slug)
    out = CACHE / "logos" / f"{slug}_{color}.png"
    if out.exists():
        return out
    svg = _fetch(f"{_SI}/icons/{slug}.svg", CACHE / "logos" / f"{slug}.svg").read_text()
    svg = svg.replace("<path ", f'<path fill="#{color}" ', 1)
    cairosvg.svg2png(bytestring=svg.encode(), write_to=str(out), output_width=px, output_height=px)
    return out


def logo(slide, slug, x, y, d=0.42, color=None):
    """Logo badge: a white disc with the logo. Returns ``[disc, picture]``."""
    pal = current().palette
    disc = box(slide, x, y, d, d, fill=pal.white, line=pal.line, radius=0.5, lw=0.75)
    shadow(disc, blur=4, dist=1, alpha=18)
    pad = d * 0.18
    pic = picture(slide, logo_png(slug, color), x + pad, y + pad, w=d - 2 * pad, h=d - 2 * pad)
    return [disc, pic]

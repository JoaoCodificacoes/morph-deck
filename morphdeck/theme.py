"""The active theme: palette and fonts every drawing helper falls back to."""
from dataclasses import dataclass, field

from .palette import Palette


@dataclass
class Theme:
    palette: Palette = field(default_factory=Palette)
    font: str = "Aptos"
    font_head: str = "Aptos Display"


_current = Theme()


def current() -> Theme:
    return _current


def use(theme: Theme) -> Theme:
    global _current
    _current = theme
    return theme

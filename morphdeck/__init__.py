"""morph-deck: PowerPoint decks built by code (or by an AI agent), with colour that means something and Morph
transitions that carry shapes from slide to slide."""
from .draw import (box, cylinder, dashed, label, line, phone, picture, pill, polyline, rgb, round_corners,
                   shadow, style_runs, text)
from .icons import icon, logo
from .morph import fade, link, morph, to_front, twin
from .motion import animate, click_build, path, stage
from .palette import Palette, Tone
from .template import Area, Deck
from .theme import Theme, current, use

__version__ = "0.1.0"

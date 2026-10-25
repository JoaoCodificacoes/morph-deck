"""Colour that means something.

Two kinds of colour, and nothing else:

* **status** — a judgement: ``bad`` (the problem, the before), ``good`` (the result, the after, what passes) and
  ``fail`` (rejected, broken). Use them only where the slide is judging something.
* **identity** — one colour per thread of the talk (a contribution, a component, a service), reused on every
  slide where that thread appears, so the audience can follow it.

Everything else is neutral: ink for text, muted for secondary text, line for rules and connectors, card for
backgrounds. A colour that is neither a status nor an identity is a bug.
"""
from dataclasses import dataclass, field


@dataclass(frozen=True)
class Tone:
    """One colour in three strengths: ``base`` for marks, ``tint`` for backgrounds, ``dark`` for text on tint."""
    base: str
    tint: str
    dark: str

    def __iter__(self):
        return iter((self.base, self.tint, self.dark))


@dataclass
class Palette:
    ink: str = "1F2933"
    muted: str = "6B7785"
    line: str = "D5DBE1"
    edge: str = "AEB6BF"
    card: str = "F2F4F7"
    white: str = "FFFFFF"
    navy: str = "1B2A38"
    slate: str = "45535E"
    bad: Tone = Tone("F28C28", "FDEBD8", "B35F12")
    good: Tone = Tone("2E9E5B", "DFF3E7", "23794A")
    fail: Tone = Tone("D64545", "FBE3E3", "A12F2F")
    identities: dict = field(default_factory=dict)

    def identity(self, name: str) -> Tone:
        """The colour of one thread of the talk; raises if the thread was never declared."""
        try:
            return self.identities[name]
        except KeyError:
            raise KeyError(f"identity '{name}' is not declared; declare it once in the palette") from None

    def declare(self, name: str, base: str, tint: str, dark: str) -> Tone:
        self.identities[name] = Tone(base, tint, dark)
        return self.identities[name]


# Identity colours that stay apart from the status colours and from each other.
SUGGESTED_IDENTITIES = [
    ("C2477F", "F8E4EE", "96305F"),  # berry
    ("7E57C2", "EEE7F7", "5E3A9E"),  # purple
    ("009DE0", "E3F4FC", "0077B6"),  # blue
    ("0E9AA7", "DCF2F4", "0A7680"),  # teal
    ("5C6BC0", "E8EAF6", "3949AB"),  # indigo
]

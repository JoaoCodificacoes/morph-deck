"""Generate skill/SKILL.md (Claude Code) from AGENTS.md, so the two never drift. Run after editing AGENTS.md:

    python scripts/build_skill.py
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FRONT = """---
name: morph-deck
description: Build PowerPoint decks with morph-deck — semantic colour, Morph-linked transitions, click builds, and
  lint/render/video checks. Use when the person asks for slides, a presentation or a .pptx, wants to restyle a deck
  on a template, or asks for Morph transitions or slide animations.
---

<!-- generated from AGENTS.md by scripts/build_skill.py; edit AGENTS.md instead -->

"""


def build():
    return FRONT + (ROOT / "AGENTS.md").read_text()


if __name__ == "__main__":
    out = ROOT / "skill" / "SKILL.md"
    out.parent.mkdir(exist_ok=True)
    out.write_text(build())
    print(f"wrote {out}")

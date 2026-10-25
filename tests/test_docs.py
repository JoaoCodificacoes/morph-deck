"""The Claude skill is generated from AGENTS.md and must not drift from it."""
import importlib.util
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_skill_matches_agents_md():
    spec = importlib.util.spec_from_file_location("build_skill", ROOT / "scripts" / "build_skill.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    assert (ROOT / "skill" / "SKILL.md").read_text() == mod.build(), "run: python scripts/build_skill.py"

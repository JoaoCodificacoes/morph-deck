"""Checks on a built .pptx for the mistakes that break Morph decks.

| rule | level | what |
|---|---|---|
| morph-target-animated | error | a shape that arrives by Morph has an entrance animation, so the Morph breaks |
| entrance-not-first | error | a shape's first effect is not its entrance, so it shows before its click |
| motion-not-at-rest | error | a shape Morph carries on ends its motion path away from its resting position |
| duplicate-link | error | two shapes on one slide share a ``!!`` name |
| double-effect-list | error | a shape has two effect lists; PowerPoint refuses the file |
| missing-target | error | an animation points at a shape that does not exist |
| inherited-font | warning | text without an explicit font; letters break on machines without the template font |
| no-link | warning | a Morph transition that carries nothing but the title |
| one-sided-link | info | a ``!!`` name on only one side of a Morph transition |
"""
import re
import sys
from dataclasses import dataclass
from pathlib import Path


NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
LEVELS = {"error": 0, "warning": 1, "info": 2}


@dataclass
class Finding:
    level: str
    slide: int
    rule: str
    message: str

    def __str__(self):
        return f"{self.level:7} slide {self.slide:>2}  {self.rule:22} {self.message}"


class _Slide:
    def __init__(self, n, el):
        self.n = n
        self.el = el
        self.names = {}  # shape id -> name
        for nv in el.iter(f"{{{NS['p']}}}cNvPr"):
            self.names[nv.get("id")] = nv.get("name", "")
        self.links = {}
        self.duplicates = set()
        for sid, name in self.names.items():
            if name.startswith("!!"):
                if name in self.links:
                    self.duplicates.add(name)
                self.links[name] = sid
        self.morph = any(isinstance(e.tag, str) and e.tag.endswith("}morph") for e in el.iter())
        self.effects = []  # (shape id, class, element) in timeline order
        timing = el.find("p:timing", NS)
        if timing is not None:
            for ctn in timing.iter(f"{{{NS['p']}}}cTn"):
                cls = ctn.get("presetClass")
                if not cls:
                    continue
                tgt = ctn.find(".//p:spTgt", NS)
                if tgt is not None:
                    self.effects.append((tgt.get("spid"), cls, ctn))

    def classes(self, sid):
        return [c for s, c, _ in self.effects if s == sid]


def _path_end(path):
    nums = re.findall(r"-?\d*\.?\d+(?:[eE]-?\d+)?", path or "")
    return (float(nums[-2]), float(nums[-1])) if len(nums) >= 2 else (0.0, 0.0)


def _load(path):
    from pptx import Presentation
    prs = Presentation(str(path))
    return [_Slide(i + 1, s._element) for i, s in enumerate(prs.slides)]


def lint(path):
    slides = _load(path)
    out = []
    add = lambda level, n, rule, msg: out.append(Finding(level, n, rule, msg))
    for i, sl in enumerate(slides):
        # structure of each slide
        for name in sorted(sl.duplicates):
            add("error", sl.n, "duplicate-link", f"'{name}' is used by more than one shape")
        for sp in sl.el.iter(f"{{{NS['p']}}}spPr"):
            if len(sp.findall("a:effectLst", NS)) > 1:
                owner = sp.getparent().find(".//p:cNvPr", NS)
                add("error", sl.n, "double-effect-list", f"'{owner.get('name') if owner is not None else '?'}'")
        for sid in dict.fromkeys(s for s, _, _ in sl.effects):
            if sid not in sl.names:
                add("error", sl.n, "missing-target", f"an animation targets shape id {sid}, which is not on the slide")
                continue
            cls = sl.classes(sid)
            if "entr" in cls and cls[0] != "entr":
                add("error", sl.n, "entrance-not-first",
                    f"'{sl.names[sid]}' starts with a {cls[0]} effect, so it is visible before its click; "
                    "put its entrance stage first")
        bare = []
        for r in sl.el.iter(f"{{{NS['a']}}}r"):
            t = r.find("a:t", NS)
            if t is None or not (t.text or "").strip():
                continue
            rpr = r.find("a:rPr", NS)
            if rpr is None or rpr.find("a:latin", NS) is None:
                bare.append(t.text.strip())
        if bare:
            sample = ", ".join(f"'{b[:24]}'" for b in bare[:3])
            add("warning", sl.n, "inherited-font", f"{len(bare)} text run(s) without an explicit font: {sample}")

        # the transition into this slide
        if i == 0 or not sl.morph:
            continue
        prev = slides[i - 1]
        shared = set(sl.links) & set(prev.links)
        for name in sorted(shared):
            if "entr" in sl.classes(sl.links[name]):
                add("error", sl.n, "morph-target-animated",
                    f"'{name}' arrives by Morph but has an entrance animation; make it static")
        for name in sorted(shared):
            sid = prev.links[name]
            moves = [(c, e) for s, c, e in prev.effects if s == sid]
            motion = [e for c, e in moves if c == "path"]
            if motion and "exit" not in [c for c, _ in moves]:
                am = motion[-1].find(".//p:animMotion", NS)
                x, y = _path_end(am.get("path") if am is not None else "")
                if abs(x) > 1e-3 or abs(y) > 1e-3:
                    add("error", prev.n, "motion-not-at-rest",
                        f"'{name}' ends its motion path away from where it rests, so it jumps back when Morph "
                        "starts; place the shape at the end of its path and make the path end at 0 0")
        if not (shared - {"!!title"}):
            add("warning", sl.n, "no-link", f"the Morph from slide {prev.n} carries nothing but the title")
        for name in sorted((set(sl.links) ^ set(prev.links)) - {"!!title"}):
            side = "this slide" if name in sl.links else f"slide {prev.n}"
            add("info", sl.n, "one-sided-link", f"'{name}' exists only on {side}")
    return sorted(out, key=lambda f: (f.slide, LEVELS[f.level]))


def report(found, path="", verbose=False):
    shown = [f for f in found if verbose or f.level != "info"]
    errors = sum(f.level == "error" for f in found)
    warnings = sum(f.level == "warning" for f in found)
    lines = [str(f) for f in shown]
    lines.append(f"{path}: {errors} error(s), {warnings} warning(s)" +
                 ("" if verbose else f", {len(found) - len(shown)} info hidden (-v)"))
    return "\n".join(lines)


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(prog="morphdeck lint")
    ap.add_argument("deck")
    ap.add_argument("-v", "--verbose", action="store_true")
    a = ap.parse_args(argv)
    found = lint(Path(a.deck))
    print(report(found, a.deck, a.verbose))
    return 1 if any(f.level == "error" for f in found) else 0


if __name__ == "__main__":
    sys.exit(main())

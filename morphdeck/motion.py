"""Click animations written as raw OOXML timing.

A slide's animation is a list of **clicks**; a click is a list of **stages** that play one after another; a stage
is a set of shapes that start together with one effect::

    animate(slide, [
        [stage([title_card])],                                   # click 1
        [stage([link], "wipe-right", 400),                       # click 2: draw the link...
         stage([target_box], delay=-100)],                       # ...then the box, overlapping by 0.1 s
    ])

``delay`` is measured from the end of the previous stage in the same click; negative values overlap.

Rules (``morphdeck lint`` enforces the first two):

* A shape's **first** effect must be its entrance. If a motion path comes first, PowerPoint shows the shape
  before its click. Put the entrance stage first and overlap the motion with a negative delay.
* A shape that arrives by Morph must not have an entrance (see ``morphdeck.morph``).
* Use a wipe only when the direction means something: a link drawn from its source, a curve drawn along time.
"""

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

# direction the reveal travels -> (presetSubtype, filter). Checked in PowerPoint: the filter names read backwards.
WIPE = {"down": (4, "wipe(up)"), "left": (8, "wipe(right)"), "up": (1, "wipe(down)"), "right": (2, "wipe(left)")}
SLIDE_W, SLIDE_H = 13.333, 7.5


def stage(shapes, effect="fade", dur=450, delay=0, path=None, offset=0):
    """One step of a click. ``effect``:

    * ``fade`` — entrance; ``wipe-down|up|left|right`` — entrance revealed in that direction
    * ``fade-out`` — exit
    * ``motion`` — move along ``path`` (see ``path()``)
    * ``pulse-loop`` — grow and shrink forever; ``offset`` staggers several loops
    """
    return {"shapes": list(shapes), "effect": effect, "dur": dur, "delay": delay, "path": path, "offset": offset}


def path(points, slide_w=SLIDE_W, slide_h=SLIDE_H, curve=False):
    """Motion path through ``points`` given in inches *relative to the shape's resting position*.

    Make the last point ``(0, 0)`` whenever the shape stays on the slide: Morph and the static view use the
    resting position, so a path that ends elsewhere snaps back when the slide changes. With ``curve=True`` and
    four points, the path is a cubic Bézier.
    """
    pts = [f"{x / slide_w:.4f} {y / slide_h:.4f}" for x, y in points]
    if curve and len(pts) == 4:
        return f"M {pts[0]} C {pts[1]} {pts[2]} {pts[3]} E"
    return "M " + " L ".join(pts) + " E"


def _effect_xml(ids, st, spid, node):
    eff, dur = st["effect"], st["dur"]
    tgt = f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
    if eff == "motion":
        return (f'<p:par><p:cTn id="{next(ids)}" presetID="0" presetClass="path" presetSubtype="0" accel="40000" '
                f'decel="40000" fill="hold" nodeType="{node}"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
                f'<p:childTnLst><p:animMotion origin="layout" path="{st["path"]}" pathEditMode="relative" '
                f'rAng="0" ptsTypes=""><p:cBhvr><p:cTn id="{next(ids)}" dur="{dur}" fill="hold"/>{tgt}'
                f'<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
                f'</p:cBhvr></p:animMotion></p:childTnLst></p:cTn></p:par>')
    if eff == "pulse-loop":
        return (f'<p:par><p:cTn id="{next(ids)}" presetID="6" presetClass="emph" presetSubtype="0" '
                f'repeatCount="indefinite" autoRev="1" fill="hold" nodeType="{node}"><p:stCondLst>'
                f'<p:cond delay="{st.get("offset", 0)}"/></p:stCondLst><p:childTnLst><p:animScale><p:cBhvr>'
                f'<p:cTn id="{next(ids)}" dur="{dur}" fill="hold"/>{tgt}</p:cBhvr>'
                f'<p:by x="150000" y="150000"/></p:animScale></p:childTnLst></p:cTn></p:par>')
    if eff == "fade-out":
        return (f'<p:par><p:cTn id="{next(ids)}" presetID="10" presetClass="exit" presetSubtype="0" fill="hold" '
                f'grpId="1" nodeType="{node}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                f'<p:animEffect transition="out" filter="fade"><p:cBhvr><p:cTn id="{next(ids)}" dur="{dur}"/>'
                f'{tgt}</p:cBhvr></p:animEffect><p:set><p:cBhvr><p:cTn id="{next(ids)}" dur="1" fill="hold">'
                f'<p:stCondLst><p:cond delay="{dur - 1}"/></p:stCondLst></p:cTn>{tgt}<p:attrNameLst>'
                f'<p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to>'
                f'<p:strVal val="hidden"/></p:to></p:set></p:childTnLst></p:cTn></p:par>')
    if eff == "fade":
        preset, sub, filt = 10, 0, "fade"
    elif eff.startswith("wipe-") and eff[5:] in WIPE:
        preset, (sub, filt) = 22, WIPE[eff[5:]]
    else:
        raise ValueError(f"unknown effect '{eff}'")
    return (f'<p:par><p:cTn id="{next(ids)}" presetID="{preset}" presetClass="entr" presetSubtype="{sub}" '
            f'fill="hold" grpId="0" nodeType="{node}"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f'<p:childTnLst><p:set><p:cBhvr><p:cTn id="{next(ids)}" dur="1" fill="hold">'
            f'<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>{tgt}'
            f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
            f'</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
            f'<p:animEffect transition="in" filter="{filt}"><p:cBhvr><p:cTn id="{next(ids)}" dur="{dur}"/>'
            f'{tgt}</p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>')


def animate(slide, clicks, auto_first=False):
    """Replace the slide's animation with ``clicks``. A click may also be a plain list of shapes (one fade).
    ``auto_first``: the first click plays by itself when the slide opens."""
    from lxml import etree
    from pptx.oxml.ns import qn

    ids = iter(range(3, 100_000))
    out, built = [], []
    for ci, click in enumerate(clicks):
        if click and not isinstance(click[0], dict):
            click = [stage(click)]
        inner, t = [], 0
        for si, st in enumerate(click):
            effects = []
            for k, sh in enumerate(st["shapes"]):
                if si == 0 and k == 0:
                    node = "afterEffect" if (auto_first and ci == 0) else "clickEffect"
                else:
                    node = "afterEffect" if k == 0 else "withEffect"
                effects.append(_effect_xml(ids, st, sh.shape_id, node))
                if st["effect"] in ("fade", "fade-out") or st["effect"].startswith("wipe"):
                    built.append((sh.shape_id, 1 if st["effect"] == "fade-out" else 0))
            t += st["delay"]
            inner.append(f'<p:par><p:cTn id="{next(ids)}" fill="hold"><p:stCondLst><p:cond delay="{max(t, 0)}"/>'
                         f'</p:stCondLst><p:childTnLst>{"".join(effects)}</p:childTnLst></p:cTn></p:par>')
            t += st["dur"]
        start = '<p:cond delay="indefinite"/>'
        if auto_first and ci == 0:
            start += '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>'
        out.append(f'<p:par><p:cTn id="{next(ids)}" fill="hold"><p:stCondLst>{start}</p:stCondLst>'
                   f'<p:childTnLst>{"".join(inner)}</p:childTnLst></p:cTn></p:par>')
    bld = "".join(f'<p:bldP spid="{i}" grpId="{g}" animBg="1"/>' for i, g in dict.fromkeys(built))
    xml = (
        f'<p:timing xmlns:p="{P_NS}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" '
        f'nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" '
        f'dur="indefinite" nodeType="mainSeq"><p:childTnLst>{"".join(out)}</p:childTnLst></p:cTn>'
        f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
        f'</p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl>'
        f'</p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
        f'<p:bldLst>{bld}</p:bldLst></p:timing>')
    sld = slide._element
    for old in sld.findall(qn("p:timing")):
        sld.remove(old)
    el = etree.fromstring(xml)
    ext = sld.find(qn("p:extLst"))
    if ext is not None:
        ext.addprevious(el)
    else:
        sld.append(el)


def click_build(slide, steps, dur=450):
    """Each list of shapes fades in on its own click."""
    animate(slide, [[stage(g, dur=dur)] for g in steps])

---
name: morph-deck
description: Build PowerPoint decks with morph-deck — semantic colour, Morph-linked transitions, click builds, and
  lint/render/video checks. Use when the person asks for slides, a presentation or a .pptx, wants to restyle a deck
  on a template, or asks for Morph transitions or slide animations.
---

<!-- generated from AGENTS.md by scripts/build_skill.py; edit AGENTS.md instead -->

# Building a deck with morph-deck

Instructions for an AI agent (or a person) building a PowerPoint deck with morph-deck. The deck is a Python
script; running it writes the `.pptx`. Read all of this before writing the first slide.

## Setup

```bash
pip install "morphdeck[logos] @ git+https://github.com/JoaoCodificacoes/morph-deck"
morphdeck inspect their-template.pptx -o deck.toml     # layouts, placeholders, fonts → starter config
```

Check the layout guesses `inspect` prints (cover, content, end, title and slide-number placeholders) and fix the
config by hand. Declare the deck's identity colours in `[identities]` (see Colour).

```python
from morphdeck import Deck, box, text, label, line, icon, animate, stage, link
deck = Deck("deck.toml")
s = deck.slide("content", "Architecture")      # title placeholder is named !!title on every slide
...
deck.save("out/deck.pptx")                       # saves and runs the linter
```

Units are inches, colours are hex strings without `#`, `deck.area` is the content area of a content slide.

## Workflow

Do these in order and agree each step with the person before moving on. Don't jump to animation.

1. **Skeleton.** List the slides: a title and one idea for each. About one slide per minute of talk.
2. **Content.** Build the slides statically. Leave animation out.
3. **Colour pass.** Remove every colour that doesn't mean something (see Colour).
4. **Link pass.** For every transition, write down what carries over: *the shape on slide N that becomes the shape
   on slide N+1, and why it's the same thing*. Then name both shapes the same with `!!`. If nothing honestly
   carries over, use a fade (`deck.slide(..., morph=False)` and `fade(s)`). Don't invent a link.
5. **Animation pass.** Builds only where order carries meaning (see Animation).
6. **Verify.** `morphdeck lint`, `morphdeck render`, `morphdeck video` + `morphdeck sheet`, and look at the
   images. A transition isn't done until you have seen its frames.

## Writing

- **Titles are plain statements of what's on the slide**, engineering style: "Recording without network",
  "What is encrypted, and what is not". No catchy slogans, no questions as titles, no colons used for drama.
- **No small summary phrases** under titles or at the bottom of a slide. Say them out loud instead.
- One idea per slide. If a slide needs a paragraph, it's two slides or it belongs in the speaker notes.
- Concrete numbers with units over adjectives: "142 → 310 ms", not "much faster".
- Numbers the person can't defend don't go on slides. If an estimate rests on an assumption, keep the
  assumption ready to say out loud.

## Colour

Colour is information. Two kinds exist, and nothing else:

| Kind | Use | In code |
|---|---|---|
| **Status** | judgements only: the problem or the before (`bad`), the result or what passes (`good`), what's refused or broken (`fail`) | `P.bad`, `P.good`, `P.fail` |
| **Identity** | one colour per thread of the talk (a contribution, a component), reused on every slide where that thread appears | `P.identity("data")` |

Everything else is neutral: `P.ink` for text, `P.muted` for secondary text, `P.line`/`P.edge` for connectors and
outlines, `P.card` for backgrounds. Each colour comes as a `Tone`: `.base` for marks, `.tint` for backgrounds,
`.dark` for text on the tint.

- Don't use the same colour for a status and an identity, e.g. orange for "requirements" and orange for "bad".
- Diagrams (ER models, architectures) are neutral unless a colour marks a thread the talk follows.
- Highlight only what the speaker is talking about.

## Morph

Shapes with the same `!!name` on consecutive slides are carried by Morph: moved, resized, recoloured, rotated.
This is what makes a deck feel like one continuous explanation.

**Patterns that work**

| Pattern | How |
|---|---|
| Zoom into a component | the box on slide N becomes a larger box on N+1; its label is a separate text box with its own `!!` name, so it glides along |
| A thing becomes what it led to | e.g. paper forms straighten into the cards of the contributions they led to |
| One source, two targets | `twin(slide, shape, "other")` draws a copy under the original with another name |
| A reference line keeps its meaning | a dashed "limit" line on a bar chart rotates into the limit line of the next chart |
| A table header from the previous diagram | entities fly up and become the criteria the next slide judges |

**Rules**

1. **The target must be static.** A shape that arrives by Morph must not have an entrance animation on slide N+1,
   or the link breaks. Animate other things.
2. **Morph uses resting positions.** If a shape moves by a motion path, place it where the path *ends* and make
   paths relative to that point (`path([(dx, dy), (0, 0)])`). Otherwise it jumps back when the slide changes.
3. **Names are unique per slide.**
4. **Keep travelling shapes out of groups.** A page whose content is grouped: the page is a separate shape, the
   contents are the group.
5. **z-order travels with the shape.** Use `to_front()` on both slides for a shape that must stay on top.
6. **Every Morph transition carries something besides the title**, or it should be a fade.

## Animation

```python
animate(s, [
    [stage([a1], "wipe-right", 350), stage([box_b], delay=-100)],    # click 1: draw the link, then the box
    [stage(group_c)],                                                 # click 2
])
```

A slide is a list of clicks; a click is a list of stages that play one after another; `delay` counts from the
end of the previous stage, and a negative delay overlaps.

- **Build in the order you speak.** One component per click on an architecture; one row per click on a table
  only if you discuss each row.
- **Wipes only when direction means something:** a connector drawn from its source, a curve drawn along time,
  a path of steps left to right. Map the direction: `wipe-right` reveals left to right.
- **A shape's first effect must be its entrance.** If a motion path comes first, the shape is visible before its
  click. Put the entrance stage first and overlap the motion with a negative delay.
- A travelling dot along a connector (fade in, move, fade out) reads as "this is passed on".
- Loops (`pulse-loop`) only for something that's genuinely ongoing, e.g. an agent "thinking".
- Don't animate what the audience should read at once, like a short list of three items.

## Fonts and templates

- Every text run gets an explicit font (`text()` and `style_runs(..., font=)` do this). Text that inherits a
  template font breaks letter spacing on machines that don't have it. Use `deck.fill()` for placeholders.
- Don't edit the template file. Draw on top of its layouts.
- Logos and brand assets of the template stay in the template; don't copy them into the deck code.

## Verify

```bash
morphdeck lint out/deck.pptx            # errors must be zero; read every warning
morphdeck render out/deck.pptx out/png  # slide PNGs (PowerPoint on Windows or WSL)
morphdeck video out/deck.pptx --slides 6,7 -o out/t.mp4
morphdeck sheet out/t.mp4 -o out/t.png  # one image of frames: look at it
```

Look at the renders and the frames yourself before saying the deck is done. Check text overflow, overlaps, a
shape visible too early, a Morph that fades instead of moving, and a wipe going the wrong way. Video export
can't run while the same file is open for editing in PowerPoint. Close it first.

## Lint rules

| Rule | Meaning | Fix |
|---|---|---|
| `morph-target-animated` | a Morph target has an entrance | make it static; build other shapes |
| `entrance-not-first` | shape visible before its click | put the entrance stage first |
| `motion-not-at-rest` | a carried shape snaps back | rest it at the end of its path |
| `duplicate-link` | two shapes share a `!!` name | rename one |
| `double-effect-list` | the file won't open | add effects to the existing effect list (`shadow()` does) |
| `missing-target` | an animation points at a deleted shape | rebuild the slide's animation |
| `inherited-font` (warning) | text without an explicit font | pass `font=` |
| `no-link` (warning) | Morph carries only the title | link something real, or use a fade |

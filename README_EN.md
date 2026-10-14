<div align="center">

<img src="docs/logo.png" width="120" alt="Report Master PPT logo"/>

# Report Master PPT · huibao-ppt

**Turn a one-line brief into a native, fully editable, presentation-grade PPTX.**

[![CI](https://github.com/DF-Guan/huibao-ppt/actions/workflows/ci.yml/badge.svg)](https://github.com/DF-Guan/huibao-ppt/actions/workflows/ci.yml)
[![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)](CHANGELOG.md)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

</div>

---

## What it solves

AI-made decks fail three ways: **template-sameness**, **silent breakage**
(misaligned figures, overflow, black-boxed logos), and **nothing to say when
you present**. This skill ships an aesthetic methodology, a rendering engine and
three quality gates, so one sentence in gets you a deck that looks right,
reads right and speaks right.

## Three styles, one flag apart

| Style | Character | Best for |
|---|---|---|
| `brutal` | Thick ink borders, hard shadows, acid yellow | Internal reviews, young teams |
| `dark_premium` | Navy field, brass gold, restrained | Client proposals, overseas decks |
| `editorial` | Serif display, asymmetric columns, inverted focus cells | Brand proposals, annual reviews |

## Highlights

- **23 layouts**: covers, TOC, stats, comparisons, lists, closings, native charts,
  photo walls, logo walls, org charts, timelines, scene pages, numbered shot stacks,
  staggered cards, free-form, quote pages, section pages, magazine triptych,
  newspaper figures, native tables — one composition per page
- **Aesthetic methodology** (`references/thinking.md`): five-dimension diagnosis,
  sample-based style selection, symptom→cause→fix table, eight-question self-review
- **Typography**: CJK/Latin font pairing, artistic text (stroke / gradient / vertical /
  italic serif / tracking), lining-figure auto-protection for numerals
- **Three quality gates**: preflight (pre-render) → LINT + alignment monitor (render)
  → one-command PNG export review (post-render)
- **Interaction**: clickable TOC jumps, back navigation, page transitions,
  speaker notes on every page

## Quick start

```bash
pip install -r requirements.txt

# User can't articulate a style? Generate a multi-style sample booklet in one command
python scripts/sample.py --title "Your topic" --lead "One-line intro" --stats "38 closed;214 revoked"

# Write a spec (copy an example), preflight, render
python scripts/preflight.py examples/deck-spec.example.json
python scripts/build_deck.py examples/deck-spec.example.json out.pptx --style editorial

# Export PNGs for page-by-page review
python scripts/export_preview.py out.pptx preview/
```

## License

[MIT](LICENSE) · Copyright (c) 2026 DF · Report Master PPT

Repo: https://github.com/DF-Guan/huibao-ppt · See [README.md](README.md) for the Chinese edition.

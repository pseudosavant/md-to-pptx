from __future__ import annotations

from pathlib import Path

import pytest

from markdown_slides.errors import ParseError, UnsupportedContentError
from markdown_slides.parser import parse_deck


def test_parse_document_and_slide_front_matter() -> None:
    deck = parse_deck(
        """---
aspect_ratio: "4:3"
fonts:
  body: Aptos
  headings: Aptos Display
title_color: "var(--light-1)"
body_color: "rgb(68, 85, 102)"
color_scheme:
  preset: Office
background: "var(--accent-1)"
---

# Intro
---
layout: Title Slide
title_color: "var(--accent-2)"
notes: |
  Speaker notes.
---

Subtitle text
""",
        input_path=Path("deck.md"),
        source_name="deck.md",
    )

    assert deck.aspect_ratio == "4:3"
    assert deck.fonts.body == "Aptos"
    assert deck.text_colors is not None
    assert deck.text_colors.title == "var(--light-1)"
    assert deck.text_colors.body == "#445566"
    assert deck.background is not None
    assert deck.background.value == "var(--accent-1)"
    assert deck.fonts_override is True
    assert deck.color_scheme.name == "Office"
    assert deck.slides[0].layout == "Title Slide"
    assert deck.slides[0].text_colors is not None
    assert deck.slides[0].text_colors.title == "var(--accent-2)"
    assert deck.slides[0].text_colors.body is None
    assert deck.slides[0].notes == "Speaker notes."


def test_parse_defaults_do_not_force_theme_overrides() -> None:
    deck = parse_deck(
        "# Slide\n\nBody\n",
        input_path=Path("deck.md"),
        source_name="deck.md",
    )

    assert deck.fonts.body == "Aptos"
    assert deck.fonts.headings == "Aptos Display"
    assert deck.fonts_override is False
    assert deck.color_scheme is None


def test_dashed_text_color_keys_are_rejected() -> None:
    with pytest.raises(ParseError) as excinfo:
        parse_deck(
            """---
title_color: "#112233"
title-color: "#445566"
---

# Slide

Body
""",
            input_path=Path("deck.md"),
            source_name="deck.md",
        )

    assert excinfo.value.context.code == "unknown_front_matter_keys"


def test_slide_front_matter_must_be_immediately_after_h1() -> None:
    with pytest.raises(ParseError) as excinfo:
        parse_deck(
            "# Slide\n\n---\nlayout: Title Only\n---\n",
            input_path=Path("deck.md"),
            source_name="deck.md",
        )

    assert excinfo.value.context.code == "setext_headings_unsupported"


def test_setext_headings_are_rejected() -> None:
    with pytest.raises(ParseError) as excinfo:
        parse_deck(
            "# Slide\n\nSubtitle\n--------\n",
            input_path=Path("deck.md"),
            source_name="deck.md",
        )

    assert excinfo.value.context.code == "setext_headings_unsupported"


def test_blank_slide_defaults_when_empty_title_and_body() -> None:
    deck = parse_deck(
        "# \n",
        input_path=Path("deck.md"),
        source_name="deck.md",
    )

    assert deck.slides[0].layout == "Blank"


def test_empty_title_with_body_defaults_to_title_and_content() -> None:
    deck = parse_deck(
        "# \n\nBody\n",
        input_path=Path("deck.md"),
        source_name="deck.md",
    )

    assert deck.slides[0].layout == "Title and Content"


def test_title_only_rejects_body() -> None:
    with pytest.raises(UnsupportedContentError):
        parse_deck(
            "# Slide\n---\nlayout: Title Only\n---\n\nBody\n",
            input_path=Path("deck.md"),
            source_name="deck.md",
        )


def test_title_and_content_rejects_mixed_text_and_image() -> None:
    with pytest.raises(UnsupportedContentError):
        parse_deck(
            "# Slide\n\nParagraph.\n\n![alt](./image.png)\n",
            input_path=Path("deck.md"),
            source_name="deck.md",
        )


def test_h1_inside_fence_does_not_start_slide() -> None:
    deck = parse_deck(
        "# Slide\n\n```python\n# not a slide\n```\n",
        input_path=Path("deck.md"),
        source_name="deck.md",
    )

    assert len(deck.slides) == 1
    assert deck.slides[0].body.paragraphs[0].kind == "code"

from __future__ import annotations

import io
import json
import shutil
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest

from markdown_slides import __version__
from markdown_slides.assets import default_template_path
from markdown_slides.cli import _apply_color_ignore_flags, build_parser, main
from markdown_slides.parser import parse_deck


def test_no_args_prints_help() -> None:
    stdout = io.StringIO()
    stderr = io.StringIO()

    exit_code = main([], stdout=stdout, stderr=stderr)

    assert exit_code == 0
    assert "markdown-pptx deck.md" in stdout.getvalue()
    assert stderr.getvalue() == ""


def test_version_flag_prints_version() -> None:
    stdout = io.StringIO()
    stderr = io.StringIO()

    with pytest.raises(SystemExit) as excinfo:
        main(["--version"], stdout=stdout, stderr=stderr)

    assert excinfo.value.code == 0
    assert stdout.getvalue() == f"markdown-pptx {__version__}\n"
    assert stderr.getvalue() == ""


def test_help_mentions_agent_friendly_modes() -> None:
    help_text = build_parser().format_help()
    assert "--list-layouts" in help_text
    assert "--list-color-schemes" in help_text
    assert "--syntax" in help_text
    assert "--ignore-document-colors" in help_text
    assert "--ignore-slide-colors" in help_text
    assert "markdown-pptx deck.md" in help_text


def test_list_color_schemes_plain_output() -> None:
    stdout = io.StringIO()
    stderr = io.StringIO()

    exit_code = main(["--list-color-schemes"], stdout=stdout, stderr=stderr)

    assert exit_code == 0
    assert "Office" in stdout.getvalue().splitlines()
    assert "Blue Warm" in stdout.getvalue().splitlines()
    assert stderr.getvalue() == ""


def test_list_color_schemes_json_output() -> None:
    stdout = io.StringIO()
    stderr = io.StringIO()

    exit_code = main(["--list-color-schemes", "--json"], stdout=stdout, stderr=stderr)

    payload = json.loads(stdout.getvalue())
    assert exit_code == 0
    assert payload["ok"] is True
    assert payload["mode"] == "list_color_schemes"
    assert "Office" in payload["color_schemes"]
    assert stderr.getvalue() == ""


def test_syntax_json_output() -> None:
    stdout = io.StringIO()
    stderr = io.StringIO()

    exit_code = main(["--syntax", "--json"], stdout=stdout, stderr=stderr)

    payload = json.loads(stdout.getvalue())
    assert exit_code == 0
    assert payload["ok"] is True
    assert "document_front_matter_keys" in payload
    assert "slide_front_matter_keys" in payload
    assert payload["aspect_ratio_values"] == ["16:9", "4:3"]
    assert payload["layout_values"] == ["Title Slide", "Title and Content", "Section Header", "Title Only", "Blank"]
    assert payload["theme_color_syntax"] == (
        "Use var(--slot-name) in text colors and backgrounds, for example var(--accent-1) or var(--dark-1)."
    )
    assert payload["color_scheme_syntax"]["preset_example"] == {"preset": "Office"}
    assert payload["color_scheme_syntax"]["custom_keys"] == [
        "dark_1",
        "light_1",
        "dark_2",
        "light_2",
        "accent_1",
        "accent_2",
        "accent_3",
        "accent_4",
        "accent_5",
        "accent_6",
        "hyperlink",
        "followed_hyperlink",
    ]
    assert payload["theme_color_variables"] == [
        "var(--dark-1)",
        "var(--light-1)",
        "var(--dark-2)",
        "var(--light-2)",
        "var(--accent-1)",
        "var(--accent-2)",
        "var(--accent-3)",
        "var(--accent-4)",
        "var(--accent-5)",
        "var(--accent-6)",
        "var(--hyperlink)",
        "var(--followed-hyperlink)",
    ]
    assert stderr.getvalue() == ""


def test_syntax_plain_output_lists_all_theme_color_variables() -> None:
    stdout = io.StringIO()
    stderr = io.StringIO()

    exit_code = main(["--syntax"], stdout=stdout, stderr=stderr)

    output = stdout.getvalue()
    assert exit_code == 0
    assert "aspect_ratio values: 16:9, 4:3" in output
    assert "layout values: Title Slide, Title and Content, Section Header, Title Only, Blank" in output
    assert "Theme color syntax:" in output
    assert "Theme color variables:" in output
    assert 'color_scheme preset example: {"preset": "Office"}' in output
    assert "color_scheme custom keys: dark_1, light_1, dark_2, light_2, accent_1, accent_2, accent_3, accent_4, accent_5, accent_6, hyperlink, followed_hyperlink" in output
    assert "var(--dark-1)" in output
    assert "var(--light-2)" in output
    assert "var(--accent-6)" in output
    assert "var(--hyperlink)" in output
    assert "var(--followed-hyperlink)" in output
    assert stderr.getvalue() == ""


def test_list_layouts_uses_default_template() -> None:
    stdout = io.StringIO()
    stderr = io.StringIO()

    exit_code = main(["--list-layouts"], stdout=stdout, stderr=stderr)

    assert exit_code == 0
    layouts = stdout.getvalue().splitlines()
    assert "Title Slide" in layouts
    assert "Title and Content" in layouts
    assert "Blank" in layouts
    assert stderr.getvalue() == ""


def test_stdin_requires_output() -> None:
    stdout = io.StringIO()
    stderr = io.StringIO()

    exit_code = main(["--input", "-"], stdout=stdout, stderr=stderr)

    assert exit_code == 2
    assert "--input - requires --output" in stderr.getvalue()


def test_positional_and_flag_input_are_mutually_exclusive(tmp_path: Path) -> None:
    deck = tmp_path / "deck.md"
    deck.write_text("# Title\n\nBody\n", encoding="utf-8")
    stdout = io.StringIO()
    stderr = io.StringIO()

    exit_code = main([str(deck), "--input", str(deck)], stdout=stdout, stderr=stderr)

    assert exit_code == 2
    assert "Use either positional input or --input" in stderr.getvalue()


def test_render_prints_default_output_path(tmp_path: Path) -> None:
    deck = tmp_path / "myFavoriteSlides.md"
    deck.write_text("# Slide\n\nBody text.\n", encoding="utf-8")
    stdout = io.StringIO()
    stderr = io.StringIO()

    exit_code = main([str(deck)], stdout=stdout, stderr=stderr)

    expected = str(deck.with_suffix(".pptx").resolve())
    assert exit_code == 0
    assert stdout.getvalue().strip() == expected
    assert Path(expected).exists()
    assert stderr.getvalue() == ""


def test_render_json_output(tmp_path: Path) -> None:
    deck = tmp_path / "deck.md"
    deck.write_text("# Slide\n\nBody text.\n", encoding="utf-8")
    output = tmp_path / "out.pptx"
    stdout = io.StringIO()
    stderr = io.StringIO()

    exit_code = main([str(deck), str(output), "--json"], stdout=stdout, stderr=stderr)

    payload = json.loads(stdout.getvalue())
    assert exit_code == 0
    assert payload["ok"] is True
    assert payload["mode"] == "render"
    assert payload["slides"] == 1
    assert payload["ignore_document_colors"] is False
    assert payload["ignore_slide_colors"] is False
    assert Path(payload["output"]).exists()
    assert stderr.getvalue() == ""


def test_apply_color_ignore_flags_keeps_image_backgrounds() -> None:
    deck = parse_deck(
        """---
background: "linear-gradient(90deg, #112233 0%, #445566 100%)"
title_color: "#010203"
body_color: "#040506"
color_scheme:
  preset: Office
---

# Slide A
---
background: "url('./bg.png')"
title_color: "#111111"
body_color: "#222222"
---

Body

# Slide B
---
background: "#778899"
title_color: "#333333"
body_color: "#444444"
---

Body
""",
        input_path=Path("deck.md"),
        source_name="deck.md",
    )

    adjusted = _apply_color_ignore_flags(deck, ignore_document_colors=True, ignore_slide_colors=True)

    assert adjusted.color_scheme is None
    assert adjusted.text_colors is None
    assert adjusted.background is None
    assert adjusted.slides[0].background is not None
    assert adjusted.slides[0].background.kind == "image"
    assert adjusted.slides[0].text_colors is None
    assert adjusted.slides[1].background is None
    assert adjusted.slides[1].text_colors is None


def test_ignore_document_colors_preserves_template_theme_and_slide_overrides(tmp_path: Path) -> None:
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    template = tmp_path / "template.pptx"
    shutil.copyfile(default_template_path(), template)
    customized_template = tmp_path / "template-customized.pptx"
    with zipfile.ZipFile(template, "r") as source, zipfile.ZipFile(
        customized_template, "w", compression=zipfile.ZIP_DEFLATED
    ) as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "ppt/theme/theme1.xml":
                root = ET.fromstring(data)
                clr = root.find(".//a:clrScheme", ns)
                assert clr is not None
                clr.set("name", "Custom Template Theme")
                accent1 = clr.find("a:accent1", ns)
                assert accent1 is not None
                accent1[0].set("val", "ABCDEF")
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            target.writestr(info, data)

    deck = tmp_path / "deck.md"
    deck.write_text(
        """---
color_scheme:
  preset: Blue Warm
title_color: "#112233"
body_color: "#445566"
---

# Slide
---
layout: Title and Content
title_color: "#778899"
body_color: "#AABBCC"
---

Body
""",
        encoding="utf-8",
    )
    output = tmp_path / "out.pptx"
    stdout = io.StringIO()
    stderr = io.StringIO()

    exit_code = main(
        [str(deck), str(output), "--template", str(customized_template), "--ignore-document-colors"],
        stdout=stdout,
        stderr=stderr,
    )

    assert exit_code == 0
    with zipfile.ZipFile(output) as zf:
        theme = ET.fromstring(zf.read("ppt/theme/theme1.xml"))
        slide_xml = zf.read("ppt/slides/slide1.xml").decode("utf-8")
    clr = theme.find(".//a:clrScheme", ns)
    assert clr is not None
    assert clr.attrib["name"] == "Custom Template Theme"
    accent1 = clr.find("a:accent1", ns)
    assert accent1 is not None
    assert accent1[0].attrib["val"] == "ABCDEF"
    assert 'val="778899"' in slide_xml
    assert 'val="AABBCC"' in slide_xml
    assert 'val="112233"' not in slide_xml
    assert 'val="445566"' not in slide_xml
    assert stderr.getvalue() == ""


def test_ignore_slide_colors_keeps_document_colors(tmp_path: Path) -> None:
    deck = tmp_path / "deck.md"
    deck.write_text(
        """---
title_color: "#112233"
body_color: "#445566"
---

# Slide
---
layout: Title and Content
title_color: "#778899"
body_color: "#AABBCC"
---

Body
""",
        encoding="utf-8",
    )
    output = tmp_path / "out.pptx"
    stdout = io.StringIO()
    stderr = io.StringIO()

    exit_code = main([str(deck), str(output), "--ignore-slide-colors"], stdout=stdout, stderr=stderr)

    assert exit_code == 0
    with zipfile.ZipFile(output) as zf:
        slide_xml = zf.read("ppt/slides/slide1.xml").decode("utf-8")
    assert 'val="112233"' in slide_xml
    assert 'val="445566"' in slide_xml
    assert 'val="778899"' not in slide_xml
    assert 'val="AABBCC"' not in slide_xml
    assert stderr.getvalue() == ""

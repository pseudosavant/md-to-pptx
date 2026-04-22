---
aspect_ratio: "16:9"
fonts:
  body: Aptos
  headings: Aptos Display
color_scheme:
  preset:
  dark_1: "#10263F"
  light_1: "#F9F9F9"
  dark_2: "#355A78"
  light_2: "#EEF8FF"
  accent_1: "#1D6FA8"
  accent_2: "#5AA9E6"
  accent_3: "#7FC8F8"
  accent_4: "#FFB347"
  accent_5: "#FF6392"
  accent_6: "#144F79"
  hyperlink: "#1D6FA8"
  followed_hyperlink: "#355A78"
title_color: "var(--accent-6)"
body_color: "var(--dark-1)"
background: "linear-gradient(90deg, #F3FAFF 0%, #E6F4FF 54%, #FFE7BF 100%)"
---

# markdown-pptx
---
layout: Title Slide
---

Markdown in. Editable PowerPoint out.

# What markdown-pptx is

`markdown-pptx` converts constrained Markdown slide decks into editable PowerPoint `.pptx` presentations.

It uses real PowerPoint layouts and placeholders, so the output stays easy to edit in PowerPoint instead of becoming a pile of free-positioned text boxes.

# How the format works

An optional document front matter block sets deck-wide defaults such as aspect ratio, fonts, colors, and background behavior.

Each `# H1` starts a new slide, optional slide front matter appears immediately after the H1, and everything until the next H1 becomes that slide's content.

# Basic CLI usage

Run `markdown-pptx deck.md` to write `deck.pptx` next to the source markdown file.

Use `markdown-pptx deck.md out.pptx` to choose the output path, `--template theme.pptx` to render against an existing PowerPoint template, and `--force` to overwrite an existing file.

# Inspection-friendly CLI modes

Use `--syntax` to print the supported deck format, `--list-layouts` to inspect the layouts available in a template, and `--list-color-schemes` to see the built-in theme presets.

These modes are especially useful when a person or coding agent needs to discover the valid inputs before generating a deck.

# Color and template control

Markdown can set document-level and slide-level colors with `color_scheme`, `title_color`, `body_color`, and background values such as solid colors, gradients, and images.

When a template already has the colors you want, `--ignore-document-colors` and `--ignore-slide-colors` let the template remain the source of truth for those color settings.

# Feature examples
---
layout: Section Header
---

Each slide after this one isolates a single `markdown-pptx` feature.

# Title Slide layout
---
layout: Title Slide
---

This slide uses the Title Slide layout.

# Section Header layout
---
layout: Section Header
---

This slide uses the Section Header layout.

# Title and Content layout

This slide uses the default Title and Content layout.

# Title Only layout
---
layout: Title Only
---

# Body text with hyperlinks

`markdown-pptx` can include links in normal body text, such as the [project repository](https://github.com/pseudosavant/markdown-pptx) and the [python-pptx documentation](https://python-pptx.readthedocs.io/en/latest/).

# H2 through H4 headings

## H2 heading
H2 stays inside the current slide body.

### H3 heading
H3 also stays inside the same slide body.

#### H4 heading
H4 is rendered smaller while still using heading styling.

# Bulleted list

- Write a readable Markdown deck
- Choose a template if needed
- Render an editable PowerPoint file

# Numbered list

1. Write the markdown source
2. Run the CLI
3. Open the generated `.pptx`

# Blockquote

> markdown-pptx keeps the source format simple enough to read directly while still producing a real presentation.

# Code block

```powershell
markdown-pptx sample/showcase.md sample/showcase.pptx --force
```

# Local image

![Local markdown-pptx sample image](./showcase-local.png)

# Remote image

![Remote Python logo](https://raw.githubusercontent.com/github/explore/main/topics/python/python.png)

# Background color
---
background: "#EAF3FF"
---

This slide uses a solid slide background color.

# Background image
---
background: "url('./showcase-local.png')"
---

This slide uses a slide background image.

# Linear gradient background
---
background: "linear-gradient(90deg, var(--accent-2) 0%, var(--accent-4) 100%)"
---

This slide uses a linear gradient background.

# Radial gradient background
---
background: "radial-gradient(circle, var(--accent-2) 0%, var(--accent-4) 58%, var(--light-1) 100%)"
---

This slide uses a radial gradient background.

# Slide colors override document colors
---
title_color: "var(--light-1)"
body_color: "var(--light-1)"
background: "linear-gradient(90deg, var(--accent-1) 0%, var(--accent-4) 100%)"
---

This slide overrides the document-level title and body colors.

# 
---
layout: Blank
notes: |
  This slide intentionally demonstrates the Blank layout.
---

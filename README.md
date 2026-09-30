# ppt-skills

Two agent skills for presentation automation — reverse-engineer any .pptx into a reusable skeleton pack, then generate new decks from plain text in minutes.

[![License: CC BY-NC-SA 4.0](https://img.shields.io/badge/License-CC%20BY--NC--SA%204.0-lightgrey.svg)](https://creativecommons.org/licenses/by-nc-sa/4.0/)
![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)
![Made for Claude Code](https://img.shields.io/badge/made%20for-Claude%20Code-blueviolet.svg)

[中文文档](README_zh.md)

## Skills

| Skill | What it does |
|---|---|
| [`/pptx`](pptx/SKILL.md) | Mechanically dissects a sample .pptx — XML measurement + pixel clustering, zero LLM guessing — into a six-piece skeleton pack: `skeleton.json`, `extracted_scheme.json`, `build_deck.py`, `examples`, `README`, asset slot specs. |
| [`/ppt-direct`](ppt-direct/SKILL.md) | Dual-branch pipeline: replicate a template's style from screenshots/.pptx, or design from scratch. The model makes design decisions and hand-writes python-pptx vector rendering. Minutes per deck. |

## Quickstart

```bash
git clone https://github.com/xu-jin-cs/ppt-skills.git
cd ppt-skills
./install.sh
```

Or one-liner (no clone needed):

```bash
curl -fsSL https://raw.githubusercontent.com/xu-jin-cs/ppt-skills/main/install.sh | bash
```

The installer symlinks both skills into `~/.claude/skills/`. Existing skills with the same name are skipped, never overwritten.

### Dependencies

- Python 3.10+
- Required: `python-pptx`, `numpy`, `Pillow`
- For replication / rasterization (pptx → page PNGs) and render-QC: LibreOffice (`soffice`) and PyMuPDF (`fitz`)

```bash
pip install python-pptx numpy Pillow PyMuPDF
# LibreOffice: https://www.libreoffice.org/download/ (macOS: brew install --cask libreoffice)
```

`soffice` is resolved from `SOFFICE_PATH` env var → `PATH` → the default macOS LibreOffice location.

## Usage

Once installed, the skills are available as slash commands in your agent:

```
# 1. Forge a reusable skeleton pack from a sample deck
/pptx /path/to/sample.pptx mypack

# 2. Generate a new deck from plain text
/ppt-direct 帮我做一份季度营收汇报PPT，8页

# or replicate an existing template's style first, then generate
/ppt-direct /path/to/template_screenshots 帮我做一份产品发布PPT，复刻该模板风格
```

## Compatibility

| Agent | Skills directory |
|---|---|
| Claude Code | `~/.claude/skills` (installed here by default) |
| Codex | `~/.codex/skills` (symlink the two skill dirs there manually) |
| Kimi Code | `~/.agents/skills` (symlink the two skill dirs there manually) |

## Examples

Three slides of a demo deck generated in one pass by [`examples/demo_deck.py`](examples/demo_deck.py) using only the public `generator_skeleton.py` helpers (fictional data):

| Divider | Data page | Closing |
|---|---|---|
| ![divider](examples/demo_deck_p01.png) | ![data page](examples/demo_deck_p02.png) | ![closing](examples/demo_deck_p03.png) |

The editable source is checked in at [`examples/demo_deck.pptx`](examples/demo_deck.pptx) — open it in PowerPoint/Keynote/LibreOffice; every element is a native vector shape.

## License

Personal, non-commercial use only. [CC BY-NC-SA 4.0](LICENSE), attribution **xu-jin-cs**.

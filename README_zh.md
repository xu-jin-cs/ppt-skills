# ppt-skills

两个演示文稿自动化 Agent 技能 —— 把任意 .pptx 逆向解剖成可复用龙骨包，再从纯文案分钟级直出全新 PPT。

[![License: CC BY-NC-SA 4.0](https://img.shields.io/badge/License-CC%20BY--NC--SA%204.0-lightgrey.svg)](https://creativecommons.org/licenses/by-nc-sa/4.0/)
![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)
![Made for Claude Code](https://img.shields.io/badge/made%20for-Claude%20Code-blueviolet.svg)

[English README](README.md)

## 技能一览

| 技能 | 能力 |
|---|---|
| [`/pptx`](pptx/SKILL.md) | 对样本 .pptx 做机械解剖 —— XML 实测 + 像素聚类，零 LLM 猜测 —— 产出龙骨包六件套：`skeleton.json`、`extracted_scheme.json`、`build_deck.py`、`examples`、`README`、素材槽位规格。 |
| [`/ppt-direct`](ppt-direct/SKILL.md) | 双分支流水线：从截图/.pptx 复刻模板风格，或从零直出设计；由模型做设计决策并手写 python-pptx 矢量渲染，分钟级出稿。 |

## 快速开始

```bash
git clone https://github.com/xu-jin-cs/ppt-skills.git
cd ppt-skills
./install.sh
```

或免克隆一键安装：

```bash
curl -fsSL https://raw.githubusercontent.com/xu-jin-cs/ppt-skills/main/install.sh | bash
```

安装器会把两个技能软链进 `~/.claude/skills/`。同名技能已存在时提示跳过，绝不覆盖。

### 依赖

- Python 3.10+
- 必需：`python-pptx`、`numpy`、`Pillow`
- 复刻/转图（pptx → 分页 PNG）与渲染目检需要：LibreOffice（`soffice`）与 PyMuPDF（`fitz`）

```bash
pip install python-pptx numpy Pillow PyMuPDF
# LibreOffice：https://www.libreoffice.org/download/（macOS：brew install --cask libreoffice）
```

`soffice` 解析顺序：`SOFFICE_PATH` 环境变量 → `PATH` → macOS 默认 LibreOffice 路径。

## 使用示例

安装完成后，在 Agent 中直接以斜杠命令调用：

```
# 1. 把样本 PPT 锻造成可复用龙骨包
/pptx /path/to/sample.pptx mypack

# 2. 从纯文案直出新 PPT
/ppt-direct 帮我做一份季度营收汇报PPT，8页

# 或先复刻已有模板风格再生成
/ppt-direct /path/to/模板截图目录 帮我做一份产品发布PPT，复刻该模板风格
```

## 兼容性

| Agent | 技能目录 |
|---|---|
| Claude Code | `~/.claude/skills`（默认安装位置） |
| Codex | `~/.codex/skills`（手动把两个技能目录软链过去） |
| Kimi Code | `~/.agents/skills`（手动把两个技能目录软链过去） |

## 示例

由 [`examples/demo_deck.py`](examples/demo_deck.py) 仅调用公开的 `generator_skeleton.py` helper 一次性生成的演示 PPT 三页（虚构数据）：

| 章节页 | 数据页 | 收尾页 |
|---|---|---|
| ![章节页](examples/demo_deck_p01.png) | ![数据页](examples/demo_deck_p02.png) | ![收尾页](examples/demo_deck_p03.png) |

可编辑源文件在 [`examples/demo_deck.pptx`](examples/demo_deck.pptx)——可用 PowerPoint/Keynote/LibreOffice 打开，全部元素为原生矢量图形。

## 许可证

仅限个人非商业使用。[CC BY-NC-SA 4.0](LICENSE)，署名 **xu-jin-cs**。

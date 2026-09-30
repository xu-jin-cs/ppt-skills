---
name: pptx
description: PPT 复刻母版 V2.0（纯母版功能，2026-08-23 用户二次裁定）。输入样本 .pptx → skeletons/forge_skeleton.py 机械解剖（XML 实测+像素聚类，零 LLM）→ 输出龙骨包六件套（skeleton.json 骨架+色彩结构 / extracted_scheme.json v2.1 / asset_slots 素材槽位 / build_deck.py 复刻脚本 / examples / README）。本技能只复刻生成复刻脚本，不生成完整 PPT；出片走 /ppt。触发：/pptx <样本pptx路径> [包名] [--style 风格词]。
aliases: ["/pptx"]
---

# pptx — PPT 复刻母版 V2.0（纯母版功能）

> **/pptx = 纯母版功能**——输入样本 .pptx，输出龙骨包（复刻脚本包）；**本技能自己不生成完整 PPT，出片侧走 /ppt-direct**。
> 母版三件套：`skeletons/forge_skeleton.py`（母版分析器）+ `skeletons/_engine/deck_engine.py`（通用引擎，由包内脚本调用）+ 龙骨包目录协议。视觉解析由 forge_skeleton 机械解剖承担（XML 实测 + 像素聚类，零 LLM 估算误差）。

## 触发语法（唯一入口）

```
/pptx <样本pptx路径> [包名] [--style 风格词]
```

匹配成功即执行（等价直跑）：

```bash
python3 ~/.claude/skills/pptx/skeletons/forge_skeleton.py \
  --pptx <样本pptx路径> --name <包名> \
  --out ~/.claude/skills/pptx/skeletons/<包名> [--style 风格词]
```

- 包名缺省：取样本文件名去扩展名（清洗为合法标识）。
- 样本不是 .pptx / 路径不存在 → 不启动，提示用户提供样本源文件；只有模板截图目录（page_*.png）时指引去 /ppt-direct 复刻分支。

### 启动回执（必须先输出）

```
✅ /pptx 母版已启动（纯复刻，不出片）
动作：forge_skeleton 机械解剖样本 → 龙骨包六件套落盘 skeletons/<包名>/
完成后的出片路径：/ppt-direct <文案或文案JSON> <包名>
```

### 龙骨包六件套（forge 产物契约）

| 产物 | 内容 |
|---|---|
| `skeleton.json` | 骨架（每页配方/元素坐标/字号字体/蒙版位置）+ 色彩结构 color_structure（hue-free 血统参数）+ 每页主色调占比 + asset_slots 素材槽位规格 |
| `extracted_scheme.json` | schema v2.1 兼容层（六 binding 字段机械推导，供 /ppt-direct 链消费） |
| `build_deck.py` | 复刻脚本（薄封装 → 通用引擎；由 /ppt-direct 出片侧调用，本技能不直接执行出片） |
| `examples/content.example.json` | 文案模板（原稿可提取文本已填入，烘焙巨字为占位符） |
| `README.md` | 包用法说明 |
| `asset_slots`（skeleton.json 内） | 素材槽位规格：id / kind / box / suggested_query——只记"要什么样的图"，不含图 |

产出完成后回执必须指向：**「出片请走 /ppt-direct 文案 + 包名」**。

## 复刻哲学（铁律，写入即契约）

**复刻 = 骨架、血统、年龄；不能有一处一模一样。**

- **骨架**：每页排版结构、元素角色与坐标、目录设计、蒙版位置、字号层级——实测自模板 XML，写死于 skeleton.json。
- **血统**：色彩**结构**（各角色饱和度/明度/色相偏移关系、每页主色调占比、明暗基调）；叙事节奏（页面序列/章节闭环）；工艺配方（幽灵回声巨字/柔边扣窗/竖条暗化等参数化）。
- **抹去铁律**：龙骨包内**零原稿素材文件、零模板色值默认值**。任何一张原图、任何一个模板 HEX 都不得成为包内资产或出片默认配色。
- **独立配色（出片侧执行，母版只备结构）**：主色调由新文案派生（auto 关键词映射：科技→蓝/党政→红/自然→绿/金融→金/医疗→青/教育→茶褐，无命中给中性色）或显式指定；全套 7 角色色板由 color_structure 结构关系现算——色相=新主色，饱和度/明度关系=模板血统。

## 边界

- 本技能**只 forge 不 build**；收到"用某包出 PPT"的诉求 → 引导用户走 `/ppt-direct <文案> <包名>`，禁止在 /pptx 内拼完整 PPT。
- 样本分析失败（非 pptx/加密/损坏）→ 照实报错终止，禁止降级为"看图猜测"冒充机械解剖。

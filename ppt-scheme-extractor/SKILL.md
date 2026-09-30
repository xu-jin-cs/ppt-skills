---
name: ppt-scheme-extractor
description: PPT 模板截图视觉逆向解析 → schema v2.1 设计规则 JSON（纯提示词技能，零脚本零网络）。适用边界：只有模板截图、无 .pptx 源文件时使用；有源文件一律走 /pptx skeletons/forge_skeleton.py 机械解剖（XML+像素级实测，精度更高、零 LLM）。大模型视觉（Read 直接读图）逐页解析 page_*.png，输出 7 大复刻维度：色板场景绑定 color_usage_map / 图层分层 layer_stack / 字体规范 font_spec / 版式区间 page_layout_rules / 装饰配额 decoration_spec / 素材黑白名单 asset_policy / 页面序列过滤 sequence_filter，附 11 节解析报告与四层复刻验收清单。触发：/ppt-extract <模板截图目录>、逆向解析PPT模板、提取PPT设计规范。
aliases: ["/ppt-extract"]
---

# ppt-scheme-extractor — PPT 截图视觉逆向解析（schema v2.1）

> **定位**：独立 Agent。仅在**只有模板截图、无 pptx 源文件**时使用——有 pptx 源文件请改用 /pptx 的 forge_skeleton.py 机械解剖（精度更高、零 LLM 估算误差）。
> **图像识别方式**：**大模型视觉**（Read 工具直接读图，多模态 LLM 在上下文内完成全部识别），禁止调用外部 OCR / ImageAnalysis 二进制，全程无联网。
> **schema v2.1 核心**：① 全部比例参数为 `[min, max]` 固定区间；② 7 大复刻维度（图层分层 / 字体规范 / 素材黑白名单 / 页面序列过滤 / 画布网格 / 章节意象 / 复刻验收清单）；③ 色值绑定使用场景；④ 装饰配额统一定义为比例区间。

## 触发语法（硬性）

```
/ppt-extract <templates_root>
```

- `templates_root`：多模板根目录。其下每个子文件夹 = 一套模板（内含 `page_01.png ~ page_N.png`，序号连续无断档）。
- 单模板兼容：若 `<templates_root>` 本身直接含 `page_*.png`，按一套模板处理。
- 参数缺失或目录不存在时，禁止启动解析，仅输出调用范例：

```
请使用标准格式触发截图逆向解析：
/ppt-extract <多模板根目录>
示例：/ppt-extract /path/to/template_screenshots
```

匹配成功必须先输出启动回执：

```
✅ PPT截图逆向解析已启动（schema v2.1）
识别方式：大模型视觉（Read 直接读图）
输入：{templates_root}（待解析模板数={N}）
输出：每模板文件夹内 extracted_scheme.json（v2.1）+ image_parse_full_report.md
```

---

## 1. 职责定位

- **核心职责**：批量遍历模板截图文件夹，通过大模型视觉逐套逆向提取设计视觉规则，输出带**风格/主题/颜色三级分类字段**与 **7 大复刻维度**的 `extracted_scheme.json`（schema v2.1）。
- **岗位价值**：schema v2.1 输出可直接驱动下游 PPT 生成流水线在量化区间内取值，解决复刻时字体走形、装饰杂乱、页面结构混乱、画面无层次四类高发跑偏。

## 2. 核心技能（共 18 项，全部内嵌本文件）

1. **有序图片批量加载（大模型视觉）**：Glob 列出 `page_*.png` 按数字排序锁序 → Read 逐页读图 → 过滤损坏/空白页并记录缺失页码 → 读首张像素尺寸推导画布宽高比写入 `canvas_grid`。
2. **页面类型逐页分类 + 枚举白名单**：page_type 必须命中 `cover/toc/chapter/content/summary`；`ending`（封底）归入 `cover`；枚举外类型映射最近邻并写入 `_extract_meta.sequence_adaptation`。
3. **页面序列过滤**：区分**标准业务序列**（cover→toc→chapter→content×N→summary）与**模板预览冗余页**（模板商展示用备选封面/版式互斥页）；输出 `sequence_filter.business_sequence`（下游唯一依据）+ `preview_pages`（逐页剔除理由）；连续 ≥3 个同 type=cover 页仅留首枚，其余入 `alternative_covers`。
4. **全局 7 维色彩提取 + 场景绑定**：primary/secondary(≥2)/accent/text_dark/text_light/background 统一大写 6 位 HEX；`color_usage_map` 为每个色角色标注允许使用的元素场景白名单（如 accent 仅允许印章/分割线/小面积点缀，禁止大面积铺底），下游禁止越界用色。
5. **三级分类打标**：`topic_tag`（[travel, business, tech, nature, culture, general] 六选一）、`style_label`（2~6 字风格词，同批同视觉语言必须复用同标签保证聚类收敛）、`color_name`（primary HEX 映射中文颜色主基调）。
6. **分页版式结构化归纳**：`page_layout_rules` 每类页面输出 desc / `margin_ratio`[min,max] / `fill_max_ratio`[min,max] / `composition`（构图坐标对象：每个视觉区块在归一化坐标系中的位置区间，如 `main_title: "垂直0.38~0.5居中"`）/ `card_spec`（count/shape/arrangement + padding_ratio/divider/index_mark 子组件，无卡片置 null）/ `decoration_quota`[min,max]。
7. **全局量化参数（layout_metrics）**：`page_margin_ratio`/`card_fill_max_ratio`/`title_body_scale`/`line_height`/`paragraph_spacing` 五字段全部 [min,max] 区间；同步估算 `image_constraint.radius_pt`（8/10/12 三档）与 `overlay_alpha` 区间。
8. **图层分层提取**：按叠加关系拆 2~5 个有序图层（底→顶），每层输出 layer/role（base_texture | background_art | foreground_decoration | content）/desc/alpha 区间/fill_range；极简模板允许 2 层，禁止虚构中层。**三层素材边界（硬性写入 layer_stack._boundary_note）**：① 底层仅纯色/肌理/渐变，禁止插画写入底层；② 中层插画仅 0.15~0.3 透明度局部/半幅，禁止全出血铺满当底色；③ 前景装饰仅边角小面积；④ 浅色底模板必须标注「浅色底·深色文字」，禁止下游反转为深色底+白字。
9. **字体四类规范**：title/subtitle_or_number/body/caption 四角色各自输出 style_guess（视觉推断：书法/行楷/宋体/黑体/无衬线等）/scale_ratio 区间/weight/letter_spacing/line_height/effect/apply_pages；**title 必须追加** `screen_ratio`（标题块占屏宽高双区间）与 `shadow_layer`（影子字双层结构：applicable/shadow_color_role/shadow_alpha/offset/layout）。真实字体不可从截图获取，`style_guess` 仅为视觉推断，渲染时映射兜底字体最近邻——该限制写入 `_extract_meta.origin_font`。
10. **素材白名单/黑名单**：`asset_policy.whitelist`（模板实际使用的装饰/意象元素，截图实证）/ `blacklist`（与模板视觉语言冲突、AI 复刻最易跑偏混入的元素）/ `page_exclusive`（页面类型专属素材，无则 null）。**自洽预检**：blacklist 含圆角卡片类词条时，全部 page_layout_rules 的 card_spec 禁现圆角形态（见技能 18）。
11. **装饰分项配额**：`decoration_spec.elements` 逐个登记 name/zone（corner|edge|background|foreground）/size_ratio 区间/max_per_page；`zone_quota` 按分区给装饰面积配额区间；每个元素追加 `asset_path`（本地素材相对路径，无则 null 并登记「装饰素材缺口」）与 `render_priority="illustration_first"`（插画优先，几何 shape 仅无图兜底）。
12. **图片约束 + 透明度冲突隔离**：radius_pt / allow_panorama / ban_micro_flower / overlay_alpha 区间 / `layer_alpha`（far_scenery/mid_illustration/texture/speckle_density 四分层）。**同层 alpha 区间交叉即判冲突**，按层序切分不交叉子区间（底层取高值段、上层取低值段），冲突明细必须登记报告「5.1 图层透明度冲突日志」。
13. **文本排版规则**：use_card_group / ban_raw_bullet / keyword_highlight / max_item_per_card。
14. **画布网格基准**：pixel_size（读图精准值）/ aspect_ratio（由像素推导）/ grid_columns / safe_area_ratio / reference_lines（后三项为视觉推断，写入 `_extract_meta` 标注）。
15. **章节意象映射（叙事类模板专用）**：判定是否叙事/章节驱动型；是则逐章节提取「章节→意象」对应（意象词必须来自截图实证，禁止编造），每条附 `suggested_query` 检索词建议；非叙事类 `applicable=false, mapping=null`。
16. **复刻验收清单生成**：每套模板生成四层验收清单（色调一致性 / 字体规范 / 装饰合规 / 版式比例），写入报告末节供复刻完成后逐项核对。
17. **标准 Scheme 对齐组装**：顶层字段/层级/key 名以本文件「标准样例」节为对齐基准；`_extract_meta` 逐字段组标注【精准像素提取】/【视觉估算-区间】/【视觉推断】/【图片无法获取，本地兜底】。
18. **黑白名单一致性阻断校验（最高优先级源头门禁）**：组装完成后、Write 输出前，扫描全部页面类型 card_spec——blacklist 含「圆角卡片/rounded_rect」类词条且 card_spec.shape 为圆角类即判**自冲突**，**中断输出**，回原图复核区分「无边文字块」（card_spec 置 null 并改述真实形态）或「真卡片容器」（修正黑名单误提取）；校验结论留痕报告第 7 节「黑白名单一致性校验」子节。

### 权责边界

1. 仅做视觉规则提取，**不生成任何业务 PPT 文案**；
2. 不修改、裁剪、重绘原始截图；
3. 无法还原英寸网格、真实字体文件、精确控件坐标——该类字段统一标记本地模板兜底；
4. **无联网、无下载、无脚本执行**，仅本地图片读取 + 文本写入；
5. **禁止外部 OCR/ImageAnalysis 二进制**，识别全部由大模型视觉完成；
6. 白名单/黑名单/图层/意象按每套模板截图实证提取——国风元素（祥云/印章/卷轴）仅为示例，禁止把任何单一风格元素库硬编码为全局默认。

---

## 3. 输入契约

1. `<templates_root>/<模板文件夹>/page_*.png`，每套模板起始 `page_01`，序号连续无断档；
2. 校验：每套截图 ≥5 张、清晰度满足色块/文字/版式/图层关系识别（LLM 视觉可读）。

## 4. 输出契约（逐模板）

1. `<模板文件夹>/extracted_scheme.json`（schema v2.1 全量设计规则 + 三级分类 + 7 大复刻维度）；
2. `<模板文件夹>/image_parse_full_report.md`（固定 11 节 + 附录A）。

### extracted_scheme.json v2.1 标准样例（对齐基准）

```json
{
  "schema_version": "2.1",
  "scheme_name": "水墨中国风山水书香主题教育（截图逆向提取）",
  "style_label": "中国风",
  "topic_tag": "culture",
  "color_name": "橙",
  "full_page_sequence": ["cover","toc","chapter","content","chapter","content","summary","cover"],
  "sequence_filter": {
    "business_sequence": ["cover","toc","chapter","content","chapter","content","summary"],
    "preview_pages": [{"page": "page_02", "reason": "模板商备选封面预览，与正式封面版式互斥"}],
    "alternative_covers": [{"page": "page_26", "desc": "备选封面版式B"}],
    "note": "下游渲染以 business_sequence 为唯一依据；末页 cover 为封底（首尾呼应）属业务序列；连续≥3个cover仅保留第一个"
  },
  "canvas_grid": {
    "pixel_size": [1280, 720],
    "aspect_ratio": "16:9",
    "grid_columns": 12,
    "safe_area_ratio": [0.05, 0.08],
    "reference_lines": "标题基线0.18 / 水平中线0.5 / 底部落款线0.92"
  },
  "fixed_theme_color": {
    "primary": "#6B3E1F",
    "secondary": ["#7FA8AD","#B8CCD0"],
    "accent": "#D65A36",
    "text_dark": "#3E2B1F",
    "text_light": "#FFFFFF",
    "background": "#E8EFF1"
  },
  "color_usage_map": {
    "primary": ["书法标题字", "章节序号", "卡片标题"],
    "secondary": ["卡片底色", "淡墨远山", "图标"],
    "accent": ["印章", "分割线", "红日小面积点缀——禁止大面积铺色"],
    "text_dark": ["正文", "说明文字"],
    "text_light": ["深色底反白标题"],
    "background": ["页面底色", "云雾留白"]
  },
  "layer_stack": [
    {"layer": 1, "role": "base_texture", "desc": "浅灰蓝白宣纸感底色", "alpha": [1.0, 1.0], "fill_range": "全画布"},
    {"layer": 2, "role": "background_art", "desc": "淡墨山水+云雾底纹", "alpha": [0.15, 0.3], "fill_range": "页面0~0.9高度"},
    {"layer": 3, "role": "foreground_decoration", "desc": "竹叶/红日/印章点缀", "alpha": [0.8, 1.0], "fill_range": "边角区域，单页≤2处"},
    {"layer": 4, "role": "content", "desc": "书法标题/卡片/正文", "alpha": [1.0, 1.0], "fill_range": "安全区内"}
  ],
  "_boundary_note": "三层素材边界（binding）：底层仅纯色/肌理，禁止插画写入底层；中层插画仅 0.15~0.3 透明度局部/半幅，禁止全出血铺满；前景装饰仅边角小面积；本模板为浅色底·深色文字，禁止反转",
  "page_layout_rules": {
    "cover": {
      "desc": "居中书法标题+水墨山水背景",
      "margin_ratio": [0.05, 0.07],
      "fill_max_ratio": [0.25, 0.35],
      "composition": {"background_art": "页面0~0.9高度满铺", "main_title": "垂直0.38~0.5居中，水平居中", "decoration_anchor": "右上角红日/左下角竹叶"},
      "card_spec": null,
      "decoration_quota": [0.6, 0.7]
    },
    "content": {
      "desc": "顶部标题栏+印章/图标锚定无边文字块组",
      "margin_ratio": [0.06, 0.08],
      "fill_max_ratio": [0.68, 0.78],
      "composition": {"title_bar": "页面0.06~0.18高度", "text_block_group": "页面0.25~0.85高度按网格分布"},
      "card_spec": null,
      "decoration_quota": [0.1, 0.2]
    }
  },
  "layout_metrics": {
    "page_margin_ratio": [0.06, 0.08],
    "card_fill_max_ratio": [0.68, 0.78],
    "title_body_scale": [2.2, 2.6],
    "line_height": [1.35, 1.5],
    "paragraph_spacing": [0.9, 1.2],
    "_note": "视觉估算区间，binding——下游排版必须在区间内取值"
  },
  "font_spec": {
    "title": {"style_guess": "书法/行楷", "scale_ratio": [2.2, 2.6], "weight": "bold", "letter_spacing": [0.04, 0.08], "line_height": [1.2, 1.3], "effect": "柔光", "apply_pages": ["cover","chapter"], "screen_ratio": {"width": [0.45, 0.6], "height": [0.2, 0.35]}, "shadow_layer": {"applicable": true, "shadow_color_role": "background 邻近淡墨色", "shadow_alpha": [0.15, 0.25], "offset": "主字下层居中大一号淡色衬底", "layout": "多行堆叠"}},
    "subtitle_or_number": {"style_guess": "行楷", "scale_ratio": [1.4, 1.7], "weight": "normal", "letter_spacing": [0.06, 0.12], "line_height": [1.2, 1.4], "effect": "none", "apply_pages": ["cover","chapter","toc"]},
    "body": {"style_guess": "宋体/黑体", "scale_ratio": [1.0, 1.0], "weight": "normal", "letter_spacing": [0.0, 0.02], "line_height": [1.35, 1.5], "effect": "none", "apply_pages": ["content","summary","toc"]},
    "caption": {"style_guess": "细楷/小号无衬线", "scale_ratio": [0.55, 0.7], "weight": "light", "letter_spacing": [0.02, 0.05], "line_height": [1.2, 1.3], "effect": "none", "apply_pages": ["cover","summary"]},
    "_note": "字体名为视觉推断，真实字体不可获取，渲染时映射兜底字体最近邻"
  },
  "decoration_spec": {
    "elements": [
      {"name": "竹叶", "zone": "corner", "size_ratio": [0.05, 0.09], "max_per_page": 2, "asset_path": null, "render_priority": "illustration_first"},
      {"name": "朱砂印章", "zone": "foreground", "size_ratio": [0.03, 0.05], "max_per_page": 1, "asset_path": null, "render_priority": "illustration_first"}
    ],
    "zone_quota": {"corner": [0.05, 0.12], "background": [0.15, 0.3], "foreground": [0.02, 0.06]},
    "_note": "asset_path 为本地插画素材相对路径，null 表示素材缺口；render_priority=illustration_first：有插画素材禁止用几何 shape 顶替"
  },
  "asset_policy": {
    "whitelist": ["淡墨山水", "云雾", "竹叶", "红日", "朱砂印章", "书法字"],
    "blacklist": ["纯色圆形色块", "细碎小花", "扁平现代圆角卡片", "百叶窗", "高饱和渐变", "霓虹色"],
    "page_exclusive": {"chapter": ["卷轴式序号底"], "content": ["序号徽章"]}
  },
  "imagery_map": {
    "applicable": true,
    "mapping": {"第一章": {"imagery": "远山/启程", "suggested_query": "水墨远山 国画"}, "第二章": {"imagery": "松柏/深耕", "suggested_query": "松柏 国画 插画"}},
    "_note": "仅叙事/章节驱动型模板提取；意象词须来自截图实证"
  },
  "image_constraint": {
    "radius_pt": 8,
    "allow_panorama": true,
    "ban_micro_flower": true,
    "overlay_alpha": [0.2, 0.3],
    "layer_alpha": {"far_scenery": [0.15, 0.25], "mid_illustration": [0.3, 0.5], "texture": [0.08, 0.15], "speckle_density": "low"}
  },
  "text_constraint": {"use_card_group": true, "ban_raw_bullet": true, "keyword_highlight": false, "max_item_per_card": 5},
  "_extract_meta": {
    "color_group": "精准像素聚合提取",
    "canvas_grid.pixel_size": "精准像素提取（读图获取）",
    "canvas_grid.grid_columns": "视觉推断",
    "style_label": "视觉归纳，同批复用同标签保证聚类收敛",
    "sequence_adaptation": "原截图含 ending 封底，已映射为 cover（首尾呼应）",
    "layout_metrics": "视觉估算区间，binding",
    "font_spec": "视觉推断，真实字体映射兜底最近邻",
    "origin_font": "图片无法获取，使用全局兜底字体"
  }
}
```

### image_parse_full_report.md 固定 11 节（硬性）

1. 逐页分类表（页码/类型/一句依据）；2. 色彩提取明细与置信度（附 color_usage_map 场景对照）；3. 单页组件拆解表；4. 相邻页面类型差异对比（禁止只写共性）；5. 图层分层规则 + **5.1 透明度冲突日志**（无冲突标注「本次提取无透明度冲突」）；6. 版式归纳依据（区间参数+构图坐标）；7. 素材白名单/黑名单与跑偏替代对照 + 末节「黑白名单一致性校验」子节；8. 章节意象映射表（非叙事类标注不适用）；9. 页面序列过滤说明；10. 复刻验收清单（四层核对，可勾选）；11. 兜底说明（不可获取项与置信度声明）。

---

## 5. 工作流程（Phase A，对每套模板循环）

1. Glob 校验 page_*.png 文件列表连续性，读首张像素尺寸写入 canvas_grid；
2. Read 逐页读图分类（枚举白名单校验），生成 full_page_sequence；
3. 页面序列过滤（preview_pages / alternative_covers / business_sequence）；
4. 全局色彩 7 维提取 + color_usage_map；
5. 三级分类打标（topic_tag / style_label / color_name）；
6. layer_stack 图层分层（含 _boundary_note）；
7. page_layout_rules 版式结构化（区间化 + composition 坐标）；
8. layout_metrics + image_constraint（透明度冲突隔离与留痕）；
9. font_spec 四角色（含 title.screen_ratio / shadow_layer）；
10. decoration_spec + asset_policy；
11. imagery_map（叙事类）/ text_constraint；
12. 组装 JSON v2.1 + _extract_meta；
13. **阻断自检**：按第 6 节逐项校验，任一不过直接终止，不输出产物；
14. Write 输出 extracted_scheme.json + image_parse_full_report.md 至该模板文件夹。

## 6. 阻断级自检清单（任一不通过禁止输出）

1. full_page_sequence 全部命中 cover/toc/chapter/content/summary 枚举，无枚举外值；
2. fixed_theme_color 7 key 齐全，secondary ≥2；三级分类字段齐全且 topic_tag 命中六选一枚举；
3. 所有区间字段为 [min, max] 双元素数组且 min ≤ max，禁止单点数值；
4. page_layout_rules 每类页面六 key 齐全，composition ≥2 个区块坐标描述；
5. sequence_filter.business_sequence 非空且全合法枚举；连续 ≥3 cover 仅留首枚，其余入 alternative_covers；
6. canvas_grid.pixel_size 为实际读图像素，aspect_ratio 与像素比一致；
7. color_usage_map 覆盖全部 7 色角色；layer_stack 2~5 层有序、末层 role=content、含 _boundary_note；
8. font_spec 四角色齐全，title 含 screen_ratio 双区间 + shadow_layer 全子字段；
9. asset_policy 三 key 齐全；decoration_spec 含 elements（含 asset_path/render_priority）+ zone_quota；
10. **黑白名单一致性**：blacklist 含圆角卡片类词条时全部 card_spec 无圆角形态，冲突已处置并留痕；
11. 报告为固定 11 节结构，含 5.1 冲突日志与第 10 节四层验收清单。

## 7. 运行约束

- **工具**：Read（含 PNG）、Write、Glob、Bash（仅 ls/文件校验）；
- **禁用**：外部 OCR、ImageAnalysis、ImageMagick、任何网络指令；
- **max_retry**：2（仅图片读取失败重试；字段缺失判定解析失效，重跑解析）；
- 错误时输出完整缺失清单，提示重新执行。

---
name: ppt-direct
description: PPT 直出工作流 V2.0（双分支单入口：首参数为 PPT 模板路径→复刻分支【00逆向解析→01设计决策(复刻模式)→02直出脚本→03渲染目检】；首参数为文案→直出分支【01→02→03，绕过逆向解析】。模型直接做设计决策+手写 python-pptx 矢量渲染，不走 12 节点重流程）
aliases: ["/ppt-direct", "/ppt-extract"]
---

# ppt-direct

## 工作流基础信息

名称：ppt-direct
描述：PPT 直出工作流 V2.0（双分支：复刻分支 4 节点 00逆向解析→01设计决策→02直出脚本→03渲染目检；直出分支 3 节点 01→02→03。模型直接做设计决策+手写 python-pptx 矢量渲染，不走 12 节点重流程，单份汇报型 PPT 分钟级出稿）

## 输入分流（启动第一步，强制）

拿到参数后先做一次位置判定，决定走哪条分支：

- **第一个位置参数是已存在的路径**（含 `page_*.png` 的模板截图目录，或 `.pptx/.potx/.ppt` 文件）→ **复刻分支**：
  - 模板为 `.pptx/.potx/.ppt` 文件时，先转换为截图：`soffice --headless --convert-to pdf` 转 PDF，再 PyMuPDF 按页出 `page_XX.png`，落入 `output_delivery/template_reverse_parser/shots/`；
  - 文案需求取第二个位置参数；**缺失时停手提示补充**（复刻 = 模板风格 + 用户内容，两者缺一不可）；
  - 执行链路：00template_reverse_parser → 01design_planner（复刻模式）→ 02direct_generator → 03render_delivery。
- **第一个位置参数是文案**（非路径的自然语言）→ **直出分支**：绕过逆向解析，执行链路 01design_planner（原创模式）→ 02direct_generator → 03render_delivery（原 V1.0 路径不变）。
- 判定歧义（如路径不存在）时向用户澄清，禁止瞎猜分支。

## 触发语法

pattern: ^/ppt-direct\s+(.+)$
example_format: /ppt-direct {模板截图目录或PPT文件路径 或 文案需求} [复刻分支必填：文案需求]
variables:
- name: requirement
  description: 首参数=路径→复刻分支（第二参数为文案需求）；首参数=文案→直出分支，保留原始格式含换行
  strip: false
success_receipt: |
  ✅ ppt-direct流水线已启动
  复刻分支：00template_reverse_parser → 01design_planner(复刻) → 02direct_generator → 03render_delivery
  直出分支：01design_planner(原创) → 02direct_generator → 03render_delivery
failure_example: |
  请使用标准格式触发ppt-direct流水线：
  /ppt-direct 你的需求描述                          # 直出分支
  /ppt-direct /path/to/模板截图目录 你的文案需求      # 复刻分支
  示例：
  /ppt-direct 帮我做一份季度营收汇报PPT，8页，数据驱动风格，含同比环比数据
  /ppt-direct /path/to/ppt_attachments/模板A 帮我做一份产品发布PPT，复刻该模板风格

## 执行模式

execution_mode: serial
allow_skip_node: true（仅限 00 template_reverse_parser 按输入分流条件跳过——直出分支强制跳过，复刻分支禁止跳过）
allow_insert_node: true
allow_reorder: false

## 可用工具白名单

available_tools: ["Read", "Write", "Bash"]

说明：节点「工具列表」仅可从上述白名单中选择真实存在的工具；禁止虚构不存在工具。节点无需工具时，整行省略「工具列表」。

## 工具配置区

mode: beginner
log_level: 2
input_schema_description: 首参数=PPT模板路径（截图目录或.pptx文件）→复刻分支（第二参数为文案需求）；首参数=文案→直出分支

锁值机器可读区（解析器强制读取，禁止删除）：
input_schema.input.description = 首参数=PPT模板路径→复刻分支（第二参数为文案需求）；首参数=文案→直出分支

## 终点约束源（必填）

constraint_source: python-pptx 直出成品 output.pptx 顶层规格 + 03_delivery.md 交付单字段定义
artifact_spec:
- field: output.pptx
  type: file(pptx)
  required: true
  spec: 画布 13.333×7.5in；slide 枚举 {cover,toc,chapter,content,summary} 至少含 cover+summary；全部图形为原生 shape/textbox（矢量），禁止位图图表；字体统一含 ea typeface 设置；元素 x+w≤13.343、y+h≤7.51、最小 0.5in；字号全局最小 14pt（2026-08-17 用户裁定删除 title/body 上限，Hero ≥88pt 由 design_trace 闸核验）；单页装饰 ≤3
- field: 03_delivery.md
  type: markdown
  required: true
  spec: 固定三章——成品（路径/页数/大小）/ 目检结论（逐页 PASS 或 问题+处置）/ 设计溯源（色板/版式/Hero 对应 01_design_spec 条目）
key_rules:
- 全部上游节点交付物字段必须服从终点约束源，禁止发明约束源之外的字段
- 终前节点必须删除约束源不识别的冗余字段
- 质检环节（render_delivery 目检）必须对照 artifact_spec 逐字段校验

## 质量门禁（强制保留完整配置，不可删减）

blocking_metrics:
- script_run_failure_count（生成脚本运行失败次数）
- preview_overflow_count（目检溢出/越界/遮挡页数）
- bitmap_chart_count（位图图表出现次数，矢量合规红线）
- desktop_copy_missing（成品未复制到 Desktop）
scorecard: 4维×25分满分100（设计溯源一致性 / 矢量合规 / 目检通过率 / 交付完整性）

> **机考分项禁止自打分**：矢量合规与交付完整性分项必须跑 `python3 ~/.claude/skills/ppt-direct/scripts/pptx_check.py --pptx <output.pptx>`（exit 0 全过 / exit 1 有违例，JSON 输出 checks 与 violations，照抄结论）；设计溯源一致性分项必须跑 `python3 ~/.claude/skills/ppt-direct/scripts/design_trace_check.py --pptx <output.pptx> --design-spec <01_design_spec.md> [--extracted-scheme <00_extracted_scheme.json>]`（exit 0 全过 / exit 2 违例，接线见 render_delivery 业务规则）；仅目检通过率与设计美观仍属软层目检。

rollback_threshold: '80'
converge_threshold: '95'
circuit_breaker:
- 单节点子循环3轮未通过→停手请用户裁决
- 同源补丁失败2次→第3次必须更换素材/方案

## 全局强制约束（原工作流全局规则完整复用后，续接固定结构性规则）

1. 全链路串行单向：直出分支 3 节点（01→02→03），复刻分支 4 节点（00→01→02→03）；仅 render_delivery 目检不通过可回流 direct_generator（覆盖重写 02_generator.py 并重跑）。
2. 设计哲学（本工作流灵魂，design_planner 必须体现、direct_generator 必须执行）：**直出分支（原创模式）**——设计决策由模型直接完成，禁止套用任何既有模板骨架；深浅底色交替节奏（深色封面/章节/总结 + 浅色数据页）；每份 PPT 至少 1 个大数字 Hero 焦点（核心结论数字 ≥88pt）；数据一律用 python-pptx 原生矢量图形绘制（圆角条/徽章/卡片），禁止贴位图图表；涨跌徽章制——提升绿 #2ED47A、下降优化橙 #FF7847，数据页右侧对齐；章节页大序号 ≥100pt 主色。**复刻分支（复刻模式）例外**——色板/版式/字体规范以 00_extracted_scheme.json 为唯一设计依据（binding），模型只做文案到模板的映射适配，禁止自创色板颠覆模板风格；Hero 焦点与矢量红线不变。
3. 全部产物写入 output_delivery/<节点英文名>/；成品 pptx 最终复制到用户 Desktop。
4. 渲染环境：python3 + python-pptx（缺失时 pip3 install python-pptx）；预览目检用 LibreOffice 转 PDF + PyMuPDF 出 PNG。
5. 参考骨架 templates/generator_skeleton.py 提供 rect/circle/text/pill/hbar_rows/header/footer/divider/page_bg/set_alpha 全套 helper，direct_generator 必须 import 或内联复用，禁止从零裸写 add_shape。
6. 单页装饰 ≤3、元素最小 0.5in、坐标红线 x+w≤13.343 / y+h≤7.51、字号全局最小 14pt（上限已删，见已知矛盾1裁决记录）。
7. 每个 Agent 输出严格遵循七大模块：身份→分层技能→输入契约→输出契约→工作流程→自检清单→调度配置；每个 Agent 内容最后单独一行固定标注：下一节点：XXX。
8. 统一资产粒度规范。
9. 区分内置工具/外部业务工具黑白名单。
10. LLM主动越权调用外部工具定义为重度阻断缺陷。
11. 全部交付物输出标准Markdown文档（代码类交付物除外：02_generator.py 与 output.pptx 按各自格式）。
12. 所有节点输出严格遵守固定模板格式。
13. 禁止使用三重反引号包裹交付物模板内容。
14. 流程图仅使用文本占位注释。
15. 每个交付物的文件名必须非空且在流程内唯一。
16. 每个节点的「名称」必须是英文标识符，禁止使用中文名称；非末节点的「下游节点」必须显式填写，且只能指向本文档已定义节点的「名称」；终点节点统一填“暂无”。
17. 「工具列表」只能来自「可用工具白名单」；「所需全局状态」无需求时省略整段。
18. 「下游节点」单一名称为串行；用英文逗号分隔多个名称表示并行分叉——这些节点全部执行完毕后，它们共同指向的汇聚节点才会就绪。
19. 终点约束源唯一：全部节点交付物字段以「终点约束源」节为唯一 Schema 标准，禁止发明约束源之外的字段。
20. 产物序号化命名：中间交付物按节点序前缀命名（01_ 设计产物、02_ 生产产物、03_ 交付产物），终点交付物 output.pptx 与交付单独立命名。
21. 回流为 Agent 级行为：自检/门禁不通过时，由当前节点主动重做「回流节点」指定上游节点的工作并覆盖重写其交付物后重检；禁止把回流目标写入「下游节点」，下游节点仅填正向链路（引擎仅消费下游节点链路推进）。
22. 质检环节必须对照终点约束源逐字段校验，并按「质量门禁」输出量化评分与阻断指标计数。

## 终止条件

condition: 【PPT渲染目检交付专员】输出 output.pptx + 03_delivery.md 且目检全页 PASS、四大阻断指标归零、成品已复制到 Desktop 后，流水线闭环结束。
message: 🎉 ppt-direct流水线全部节点执行完毕，成品与完整交付清单已生成于运行目录

## Agent 节点固定输出模板

Agent 节点0：模板逆向解析专员（仅复刻分支执行；直出分支按输入分流强制跳过）
节点角色：parser
名称：template_reverse_parser
身份/角色描述：仅复刻分支执行。按 schema v2 单模板解析规范轻量执行：Glob 锁定 page_*.png 顺序 → Read 大模型视觉逐页识别 → 提取色板/color_usage_map、canvas_grid、layout_metrics 区间、font_spec、layer_stack、页面序列过滤 sequence_filter（business_sequence 为下游唯一依据），输出 00_extracted_scheme.json 作为 design_planner 复刻模式的唯一设计约束源。
技能列表：
- 大模型视觉逐页识别（Read 直接读图，禁止外部 OCR）
- 色板与色彩使用场景绑定提取
- 版式量化参数区间化提取
- 页面序列过滤（剔除模板预览冗余页）
业务规则列表：
- 输入为模板截图目录时直接用；输入为 .pptx/.potx/.ppt 文件时先按「输入分流」节转换为 page_XX.png 再解析。
- 单模板解析：不执行 Phase B（训练数据整合/素材下载），不做批量遍历。
- page_type 枚举白名单：cover/toc/chapter/content/summary；ending 归入 cover，枚举外类型映射最近邻并记录 _extract_meta.sequence_adaptation。
- sequence_filter.business_sequence 必须剔除预览冗余页，作为下游页面序列唯一依据。
- 色值必须绑定使用场景（color_usage_map），比例参数一律输出 [min,max] 区间。
- 解析字段契约以本节点交付物1 的 schema v2 清单为权威（canvas_grid/color_usage_map/layout_metrics/font_spec/layer_stack/sequence_filter/decoration_spec），本节点只做单模板轻量执行。
输入契约：
输入源：用户提供的模板截图目录或 PPT 文件（转换后落 output_delivery/template_reverse_parser/shots/）
输入校验规则：
- page_*.png 数量 ≥3 且序号连续；不足或断档时记录缺失页码并提示用户，禁止凭空编造页面
- **机械门禁（判定禁止手写）**：逐页视觉识别前必须跑 `python3 ~/.claude/skills/ppt-direct/scripts/ppt_shots_check.py --shots-dir <模板截图目录>` 照抄输出；exit 0 才允许开解析，exit 1 照抄 JSON violations（缺失页码清单）提示用户补齐后重新执行
交付物列表：
【交付物1】
文件名：00_extracted_scheme.json
模板内容：（schema v2 JSON：canvas_grid/color_usage_map/layout_metrics/font_spec/layer_stack/sequence_filter/decoration_spec）
【交付物2】
文件名：00_parse_report.md
模板内容：
# 模板逆向解析报告
## 模板概况（页数/风格/主题/颜色三级分类）
## 页面序列过滤结论（business_sequence + 剔除清单及理由）
## 关键设计参数摘要（色板/字体/版式区间）
## 复刻风险提示（黑白名单冲突/序列-意象不匹配等）
自检清单：
阻断级：
- page_*.png 缺失未记录 / page_type 越出枚举白名单 / business_sequence 为空
优化级：
- 比例参数未区间化 / color_usage_map 缺场景绑定
回流节点：暂无
工具列表：
- Read
- Write
- Bash
下游节点：design_planner
下一节点：design_planner

Agent 节点1：PPT设计决策师
节点角色：coordinator
名称：design_planner
身份/角色描述：双模式。**原创模式（直出分支）**：读取用户需求，直接完成全部设计决策——主题色板（深底主色+浅底+主色+点缀+涨跌双色+正文+弱化八色）、页面序列（cover/toc/chapter×N/content×N/summary）、每页版式意图、数据可视化形式（矢量数据条/大数字卡/徽章）、Hero 焦点数字选型。**复刻模式（复刻分支）**：读取 00_extracted_scheme.json + 用户文案，色板/字体/版式区间全部从方案继承（binding，禁止自创），只做文案→business_sequence 的页面映射与数据填充决策。两种模式都输出 01_design_spec.md 作为 direct_generator 的唯一设计依据。
技能列表：
- 需求结构化拆解
- 色板与视觉节奏决策（原创）/ 方案继承与映射（复刻）
- 数据可视化选型
- Hero 焦点提炼
业务规则列表：
- 必须读取用户原始输入 input_source/input.txt；复刻模式必须额外读取 output_delivery/template_reverse_parser/00_extracted_scheme.json。
- 输出色板必须包含八色 HEX：深底 navy、浅底 bg、主色 primary、点缀 accent、涨 up、降 down、正文 ink、弱化 mute。**复刻模式：八色从 color_usage_map 映射得出，逐色标注来源条目，禁止自创 HEX。**
- 页面序列必须覆盖 cover 与 summary；chapter/content 按内容板块拆分；明确标注每页 Hero 焦点（大数字/核心结论）。**复刻模式：页面序列以 scheme 的 business_sequence 为骨架，文案板块数与骨架页数不一致时在 01_design_spec.md 中写明增删页适配理由。**
- 数据页必须逐页给出：指标名、基线值、新值、变化文案与方向（up/down）。
- 不写代码、不生成坐标级排版（坐标由 direct_generator 按骨架 helper 决定；复刻模式需给出落在 scheme 版式区间内的版式意图）。
输入契约：
输入源：input_source/input.txt（用户原始需求）；复刻模式追加 output_delivery/template_reverse_parser/00_extracted_scheme.json
输入校验规则：
- 需求文本非空；未指定页数默认按内容板块数自适应（5~15 页）；复刻模式校验 scheme 含 business_sequence 与 color_usage_map，缺失则回流 template_reverse_parser 补解析
交付物列表：
【交付物1】
文件名：01_design_spec.md
模板内容：
# PPT 设计规范
## 0 模式标注（原创｜复刻；复刻须注明来源 scheme 路径）
## 1 色板（八色 HEX：navy/bg/primary/accent/up/down/ink/mute；复刻模式逐色标注 scheme 来源）
## 2 页面序列与版式意图（逐页：page_type + 版式描述 + Hero 焦点；复刻模式标注对应 business_sequence 页码）
## 3 数据映射表（指标/基线值/新值/变化文案/方向 up|down）
## 4 Hero 焦点清单（≥1 个 ≥88pt 大数字）
自检清单：
阻断级：
- 八色缺失 / 页面序列缺 cover 或 summary / 数据页缺变化方向 / 复刻模式自创色板未标来源
优化级：
- Hero 焦点不足 1 个 / 色板无深浅交替节奏（复刻模式以 scheme 节奏为准，不强制）
回流节点：template_reverse_parser（仅复刻模式 scheme 缺字段时；原创模式填 暂无）
工具列表：
- Read
- Write
下游节点：direct_generator
下一节点：direct_generator

Agent 节点2：PPT直出生成工程师
节点角色：producer
名称：direct_generator
身份/角色描述：读取 01_design_spec.md，复用 templates/generator_skeleton.py 的 helper 函数族，手写完整 python-pptx 生成脚本 02_generator.py——全部图形矢量化（rect/circle/text/pill/hbar_rows），深浅底交替、大数字 Hero、涨跌徽章、章节大序号全部落地，脚本可独立运行直出 pptx。
技能列表：
- python-pptx 矢量绘制
- 骨架 helper 复用
- 设计规范代码化
业务规则列表：
- 必须读取 01_design_spec.md 与 templates/generator_skeleton.py。
- 必须 import 或内联复用 skeleton 的 helper 函数（rect/circle/text/pill/hbar_rows/header/footer/divider/page_bg/set_alpha），禁止从零裸写 add_shape。
- 数据条/徽章/卡片全部为原生 shape，禁止插入位图图表；中文字体统一 Microsoft YaHei（run 级 rPr 同步设置 a:ea typeface）。
- 坐标红线：x+w≤13.343、y+h≤7.51；元素最小 0.5in；字号全局最小 14pt（2026-08-17 用户裁定删除 title/body 上限，Hero 大字号为设计哲学强制项）。
- 脚本入口固定输出路径变量 OUT，默认输出 output_delivery/render_delivery/output.pptx。
- 脚本必须可 python3 02_generator.py 独立运行成功，禁止伪代码。
输入契约：
输入源：output_delivery/design_planner/01_design_spec.md + templates/generator_skeleton.py
输入校验规则：
- 01_design_spec.md 八色齐全、页面序列完整、数据映射逐页闭合
交付物列表：
【交付物1】
文件名：02_generator.py
模板内容：（完整可运行的 python-pptx 脚本，以可运行性为准，不使用文本模板约束）
自检清单：
阻断级：
- 脚本语法错误 / 裸写 add_shape / 位图图表 / 坐标越界 / 缺 ea typeface
优化级：
- helper 复用率 <80% / 装饰密度 >3 每页
回流节点：design_planner（设计规范缺失或矛盾时）
工具列表：
- Read
- Write
下游节点：render_delivery
下一节点：render_delivery

Agent 节点3：PPT渲染目检交付专员
节点角色：delivery
名称：render_delivery
身份/角色描述：执行 02_generator.py 生成 output.pptx；LibreOffice 转 PDF + PyMuPDF 出 PNG，Read 逐页目检（溢出/遮挡/对比度/版式失衡），不通过回流 direct_generator 修脚本重跑；通过后复制成品到用户 Desktop，输出 03_delivery.md 流程闭环。
技能列表：
- 脚本执行与报错定位
- 渲染预览目检
- 质量门禁评分
- 成品交付归档
业务规则列表：
- 固定渲染命令：cd output_delivery/render_delivery && python3 ../direct_generator/02_generator.py；失败按报错修复脚本（计为 direct_generator 回流）。
- **设计溯源一致性机械门禁（判定禁止手写）**：output.pptx 生成后、逐页目检前必须跑检查器照抄输出：复刻分支 `python3 ~/.claude/skills/ppt-direct/scripts/design_trace_check.py --pptx output_delivery/render_delivery/output.pptx --design-spec output_delivery/design_planner/01_design_spec.md --extracted-scheme output_delivery/template_reverse_parser/00_extracted_scheme.json`；直出分支同命令但省略 `--extracted-scheme`。机考三项：色板 binding 差集（自创 HEX 违例）/ 最大字号 ≥88pt Hero / 涨跌徽章固定色号存在性。**exit 2 即打回 02 重修（回流 direct_generator 覆盖重写 02_generator.py 重跑），禁止带违例进入目检；修复后重新执行，exit 0 才允许继续。**
- 预览命令：soffice --headless --convert-to pdf 转 PDF，PyMuPDF 按 70dpi 出 PNG，Read 读图逐页目检。
- 目检项：文字溢出、元素遮挡、对比度、越界、装饰密度、Hero 焦点是否突出；全部通过才允许交付。
- 对照 artifact_spec 逐字段校验，并按质量门禁输出四维评分与阻断指标计数；评分 <80 回流重做，≥95 方可收敛交付。
- 目检不通过 → 回流 direct_generator 覆盖重写 02_generator.py 并重跑，目检轮次上限 3 轮。
- 交付动作：cp output.pptx ~/Desktop/<主题命名>.pptx；03_delivery.md 记录页数、文件大小、目检结论、门禁评分、设计溯源。
- 交付感知（2026-08-14 复盘补丁，强制）：cp 完成后必须执行 open -R ~/Desktop/<主题命名>.pptx 在 Finder 中亮显成品（用户明确要求不打开时除外）；最终回复首行必须单独显眼给出成品完整路径（如「📂 成品：~/Desktop/xxx.pptx」），禁止只埋在正文或表格中——仅文字提及路径不构成有效交付。
- **交付实证机械门禁（判定禁止手写）**：cp 完成后、输出交付回执前必须跑 `python3 ~/.claude/skills/ppt-direct/scripts/pptx_check.py --pptx output_delivery/render_delivery/output.pptx --desktop-copy ~/Desktop/<主题命名>.pptx` 照抄输出；exit 1 即成品未落 Desktop 或存在矢量合规违例，回补 cp 后重新执行。注：open -R 亮显为 Finder 交互动作、无落盘产物可机械核验，留软层自觉执行，不进机考。
输入契约：
输入源：output_delivery/direct_generator/02_generator.py
输入校验规则：
- 脚本可运行且产出 output.pptx；pptx 页数与 01_design_spec.md 页面序列一致
交付物列表：
【交付物1】
文件名：output.pptx
模板内容：（二进制成品，以 artifact_spec 为准）
【交付物2】
文件名：03_delivery.md
模板内容：
# PPT 直出交付单
## 成品（路径/页数/大小）
## 目检结论（逐页 PASS / 问题+处置）
## 门禁评分（四维×25 与阻断指标计数）
## 设计溯源（色板/版式/Hero 对应 01_design_spec 条目）
自检清单：
阻断级：
- pptx 未生成 / 目检存在未处置缺陷 / 未复制到 Desktop / 阻断指标非零
优化级：
- 门禁评分 80~95 之间需注明让步理由（体积阈值已删，见已知矛盾2裁决记录）
回流节点：direct_generator
工具列表：
- Read
- Write
- Bash
下游节点：暂无
下一节点：暂无（终点）

### 节点中英文映射表
0. 模板逆向解析专员 → template_reverse_parser（仅复刻分支）
1. PPT设计决策师 → design_planner
2. PPT直出生成工程师 → direct_generator
3. PPT渲染目检交付专员 → render_delivery

### 已知矛盾与适用边界（2026-08-16 复盘实证增补）
1. **双闸不可同真（2026-08-17 用户裁定已闭环：删上限保 Hero）**：design_trace_check 要求 Hero 最大字号 ≥88pt，pptx_check 原字号上限 title 54pt / body 24pt 与之矛盾——用户裁定**删除 pptx_check 字号上限**（保留全局最小 14pt 下限），Hero ≥88pt 闸保留。pptx_check.py 与本文件三处条款已同步，此矛盾自 2026-08-17 起不再存在，遇旧文档引用以上条款为准。
2. **体积下限与纯矢量红线互斥（2026-08-17 用户裁定已闭环：不对 file_size 设定门槛）**：pptx_check 原体积区间 100KB~50MB 与纯矢量禁位图红线互斥——用户裁定**删除 file_size 门槛整项**，体积仅作观测值记录（03_delivery.md 软层如实登记大小）。pptx_check.py 已同步，此矛盾自 2026-08-17 起不再存在。
3. **纯图片诉求不适用本流水线**：用户要的是"图片/PNG/卡片图"（非可编辑 PPT）时，走 HTML→无头浏览器截图直出（实证路径：playwright 截图，分钟级），禁止命中本流水线。

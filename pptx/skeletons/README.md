# PPT 复刻母版体系（skeletons/）

> 一次分析，永久复用：被复刻的 pptx 放进母版 → 机械解剖出龙骨包 → 以后出片 = 文案 + 主色调 + 素材图，秒级生成。

## 三件套

| 件 | 路径 | 职责 |
|---|---|---|
| 母版分析器 | `forge_skeleton.py` | pptx → XML 元素实测 + 像素聚类 + 素材分类 → 生成龙骨包 |
| 通用引擎 | `_engine/deck_engine.py` | 数据驱动渲染：主色调统一迁移 / 插图槽位 / 文案注入 / 幽灵回声巨字等烘焙工艺 |
| 龙骨包 | `<name>/`（首例 `jianghong_forged/`） | skeleton.json（龙骨+色调分布）+ extracted_scheme.json（v2.1 兼容层）+ assets/ + build_deck.py + examples/ + README |

## 用法

```bash
# ① 复刻新模板（一次性）
python3 forge_skeleton.py --pptx 模板.pptx --name my_skel --out skeletons/my_skel [--style 中国风]

# ② 日常出片（每次）
cd skeletons/my_skel
python3 build_deck.py --content 文案.json --out 输出.pptx        # 主色 auto 按文案推导
python3 build_deck.py --content 文案.json --primary "#2B5EA7"    # 指定主色，全套统一迁移
python3 build_deck.py --content 文案.json --fetch "蓝色的中国风"  # 自动下载素材图
```

## 可变仅 4 类（其余全部固定）
主色调 / 插图（目录竖条·hero·封底） / 背景图（纸底·顶枝） / 文案。
固定：每页排版、元素坐标、蒙版位置、目录设计、字号字体、装饰工艺——全部实测自原稿 XML，写死于 skeleton.json。

## 与 /pptx 主链关系
龙骨包内 `extracted_scheme.json` 即 ppt-extract schema v2.1 兼容层（六 binding 字段机械推导），后续节点03 载入改造可直接消费（另案推进）。

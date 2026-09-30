#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
deck_engine.py — 龙骨包通用渲染引擎（数据驱动，由 forge_skeleton.py 母版生成的 skeleton.json 驱动）
固定：页面骨架/元素坐标/字号字体/装饰工艺/蒙版位置（全部来自 skeleton.json 实测值）
可变：主色调（--primary/--secondary/--accent 或 auto 由文案推导）、插图/背景图（槽位）、文案（--content）
渲染终端：技能根目录下的 render_ppt.py（相对本文件 ../../render_ppt.py 解析，可用 PPT_RENDERER 环境变量覆盖）
"""
import argparse, colorsys, json, os, subprocess, sys, tempfile
from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageEnhance

RENDERER = os.environ.get("PPT_RENDERER") or os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "render_ppt.py")
SW_IN, SH_IN = 13.333, 7.5
FONT_PATHS = {
    "Xingkai SC": ["/System/Library/AssetsV2/com_apple_MobileAsset_Font8/13b8ce423f920875b28b551f9406bf1014e0a656.asset/AssetData/Xingkai.ttc"],
    "STXingkai": ["/System/Library/AssetsV2/com_apple_MobileAsset_Font8/13b8ce423f920875b28b551f9406bf1014e0a656.asset/AssetData/Xingkai.ttc"],
    "Kaiti SC": ["/System/Library/AssetsV2/com_apple_MobileAsset_Font8/88d6cc32a907955efa1d014207889413890573be.asset/AssetData/Kaiti.ttc"],
}
def bake_font(name):
    for k, v in FONT_PATHS.items():
        if k in (name or ""):
            for p in v:
                if os.path.isfile(p):
                    return p
    for p in FONT_PATHS["Kaiti SC"]:
        if os.path.isfile(p):
            return p
    raise SystemExit("无可用烘焙字体")

def hex2rgb(h):
    h = str(h).lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4)) if len(h) == 6 else (0, 0, 0)

# ══ 主色调统一：红色系素材 hue 迁移 ════════════════════════════
def hue_of(hexc):
    r, g, b = [v / 255 for v in hex2rgb(hexc)]
    return colorsys.rgb_to_hsv(r, g, b)

def is_red_family(hexc):
    h, s, v = hue_of(hexc)
    return s > 0.25 and (h < 0.07 or h > 0.93)

def hue_shift_image(src, dst, target_hex, src_hue=0.0):
    """把素材中红色系像素的色相旋转到目标主色（保留明度/饱和度关系），透明通道不动"""
    import numpy as np
    im = Image.open(src).convert("RGBA")
    arr = np.asarray(im).astype(np.float32)
    rgb = arr[..., :3] / 255.0
    mx, mn = rgb.max(-1), rgb.min(-1)
    diff = mx - mn + 1e-6
    h = np.zeros_like(mx)
    r, g, b = rgb[..., 0], rgb[..., 1], rgb[..., 2]
    idx = (mx == r); h[idx] = ((g - b)[idx] / diff[idx]) % 6
    idx = (mx == g); h[idx] = (b - r)[idx] / diff[idx] + 2
    idx = (mx == b); h[idx] = (r - g)[idx] / diff[idx] + 4
    h = h / 6.0
    s = np.where(mx < 1e-6, 0, diff / (mx + 1e-6))
    th, ts, tv = hue_of(target_hex)
    delta = (th - src_hue) % 1.0
    mask = ((h < 0.09) | (h > 0.91)) & (s > 0.25)          # 只动红色系
    h2 = np.where(mask, (h + delta) % 1.0, h)
    s2 = np.where(mask, np.clip(s * 0.7 + ts * 0.3, 0, 1), s)
    # hsv→rgb
    i = np.floor(h2 * 6).astype(np.int32) % 6
    f = h2 * 6 - np.floor(h2 * 6)
    v2 = mx
    p = v2 * (1 - s2); q = v2 * (1 - f * s2); t = v2 * (1 - (1 - f) * s2)
    out = np.zeros_like(rgb)
    for c, (rr, gg, bb) in enumerate([(v2, t, p), (q, v2, p), (p, v2, t), (p, q, v2), (t, p, v2), (v2, p, q)]):
        m = i == c
        out[..., 0][m], out[..., 1][m], out[..., 2][m] = rr[m], gg[m], bb[m]
    arr[..., :3] = out * 255
    im_out = Image.fromarray(arr.astype("uint8"), "RGBA")
    if dst.lower().endswith((".jpg", ".jpeg")):
        im_out = im_out.convert("RGB")
    im_out.save(dst)
    return dst

# ══ 程序化工艺件（通用工艺，零原稿素材复用） ═══════════════════
def gen_paper(color, out, w=2667, h=1500):
    """宣纸底：底色+纤维噪点+轻晕影（程序化，非原稿图）"""
    import numpy as np
    rng = np.random.default_rng(42)
    base = np.zeros((h, w, 3), dtype=np.float32)
    base[:] = hex2rgb(color)
    noise = rng.normal(0, 4.0, (h, w, 1))
    fiber = rng.normal(0, 2.2, (h, w, 1))
    im_arr = np.clip(base + noise + fiber, 0, 255).astype("uint8")
    im = Image.fromarray(im_arr, "RGB")
    # 轻晕影
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).rounded_rectangle([-w*0.1, -h*0.1, w*1.1, h*1.1], radius=int(h*0.5), fill=18)
    mask = mask.filter(ImageFilter.GaussianBlur(200))
    dark = Image.new("RGB", (w, h), tuple(max(0, c-18) for c in hex2rgb(color)))
    im = Image.composite(dark, im, mask)
    im.save(out)
    return out

def gen_ink_blot(color, out, S=220):
    """墨点章：14 椭圆簇+边缘溅点（工艺配方参数化，零原稿）"""
    import random
    random.seed(7)
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    cx, cy = S/2, S/2
    for _ in range(14):
        dx, dy = random.uniform(-9, 9), random.uniform(-9, 9)
        r = random.uniform(S*0.20, S*0.30)
        d.ellipse([cx-r+dx, cy-r+dy, cx+r+dx, cy+r+dy], fill=color)
    for _ in range(26):
        import math
        a = random.uniform(0, 6.283); rr = random.uniform(S*0.32, S*0.46)
        r = random.uniform(1.5, 4.5)
        x, y = cx + rr*math.cos(a), cy + rr*math.sin(a)
        d.ellipse([x-r, y-r, x+r, y+r], fill=color)
    im = im.filter(ImageFilter.GaussianBlur(1.2))
    im.save(out)
    return out

def gen_arc(color, out, S=132):
    """笔触小弧：分段椭圆沿弧排布，中粗两细"""
    import math, random
    random.seed(3)
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    for i in range(9):
        t = i / 8
        ang = math.radians(200 - 140*t)
        rr = S*0.34
        x, y = S/2 + rr*math.cos(ang), S/2 - rr*math.sin(ang)
        w = 6 - 3.5*abs(t-0.5)*2
        d.ellipse([x-w, y-w*0.55, x+w, y+w*0.55], fill=color)
    im = im.filter(ImageFilter.GaussianBlur(0.8))
    im.save(out)
    return out

def gen_cloud(color, out, w=440, h=200):
    """简笔祥云：重叠圆+尾弧（程序化通用形）"""
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    for cx, cy, r in [(0.28, 0.55, 0.22), (0.45, 0.42, 0.26), (0.62, 0.52, 0.2), (0.76, 0.6, 0.14)]:
        d.ellipse([w*(cx-r), h*(cy-r*1.6), w*(cx+r), h*(cy+r*0.9)], fill=color)
    d.rounded_rectangle([w*0.16, h*0.52, w*0.84, h*0.78], radius=int(h*0.13), fill=color)
    im = im.filter(ImageFilter.GaussianBlur(2.5))
    im.save(out)
    return out

def gen_band(color_a, color_b, out, w=2340, h=300):
    """横带兜底：主→辅双色柔渐变带（无素材时的中性兜底，打印警告）"""
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    ca, cb = hex2rgb(color_a), hex2rgb(color_b)
    d = ImageDraw.Draw(im)
    for x in range(w):
        t = x / w
        c = tuple(int(ca[i]*(1-t) + cb[i]*t) for i in range(3))
        d.line([(x, 0), (x, h)], fill=c + (200,))
    mask = Image.new("L", (w, h), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, int(h*0.18), w, int(h*0.82)], radius=int(h*0.3), fill=255)
    mask = mask.filter(ImageFilter.GaussianBlur(18))
    im.putalpha(mask.point(lambda v: int(v*0.55)))
    im.save(out)
    return out

def veil(im_rgba, paper_hex, alpha=0.30):
    """纸色纱层：素材上罩一层纸底色的半透明纱，使其沉进纸底（样本融合工艺）"""
    veil_im = Image.new("RGBA", im_rgba.size, tuple(hex2rgb(paper_hex)) + (int(alpha*255),))
    out = im_rgba.copy()
    out.alpha_composite(veil_im)
    return out

def apply_band_mask(src, box, out, dpi=200):
    """真实横带素材 → 笔刷柔边扣窗（对齐模板横带的不规则边缘观感）"""
    tw, th = max(200, int(box[2]*dpi)), max(80, int(box[3]*dpi))
    im = crop_cover_fit(Image.open(src).convert("RGB"), tw, th)
    mask = Image.new("L", (tw, th), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, int(th*0.10), tw, int(th*0.90)], radius=int(th*0.42), fill=255)
    mask = mask.filter(ImageFilter.GaussianBlur(max(6, int(th*0.08))))
    o = Image.new("RGBA", (tw, th), (0, 0, 0, 0))
    o.paste(im, (0, 0), mask)
    o.save(out)
    return out

def gen_strip(color, out, w=590, h=1500):
    """目录竖条兜底：主色竖向深渐变"""
    im = Image.new("RGB", (w, h))
    d = ImageDraw.Draw(im)
    r, g, b = hex2rgb(color)
    for y in range(h):
        t = y / h
        f = 0.45 + 0.35*t
        d.line([(0, y), (w, y)], fill=(int(r*f), int(g*f), int(b*f)))
    im.save(out)
    return out

# ══ 烘焙工艺库 ═══════════════════════════════════════════════
def draw_giant(base, text, region, font_name, ink, ghost_rgba, dpi=200):
    """幽灵回声巨字：region=(x,y,w,h) inch；4字=2x2错落 / 3字=1+2 / 其他=单行居中"""
    x, y, w, h = [v * dpi for v in region]
    n = len(text)
    if n == 4:
        cs = min(w / 2.4, h / 2.1); pos = [(0, 0), (1.02, 0.02), (0.30, 0.98), (1.32, 1.0)]
    elif n == 3:
        cs = min(w / 2.4, h / 2.1); pos = [(0.15, 0), (1.30, 0.06), (0.62, 1.02)]
    else:
        cs = min(w / max(n, 1) * 1.15, h)
        pos = [(i, 0) for i in range(n)]
    f = ImageFont.truetype(bake_font(font_name), int(cs))
    echo = Image.new("RGBA", base.size, (0, 0, 0, 0))
    de = ImageDraw.Draw(echo)
    fe = ImageFont.truetype(bake_font(font_name), int(cs * 1.06))
    off = int(0.28 * dpi)
    for ch, (cx, cy) in zip(text, pos):
        de.text((x + cx * cs + off, y + cy * cs * 0.96 + off), ch, font=fe, fill=ghost_rgba)
    base.alpha_composite(echo.filter(ImageFilter.GaussianBlur(2.2)))
    d = ImageDraw.Draw(base)
    for ch, (cx, cy) in zip(text, pos):
        d.text((x + cx * cs, y + cy * cs * 0.96), ch, font=f, fill=ink)

def draw_giant_cutout(base, art_rgb, text, region, font_name, ghost_rgba, dpi=200):
    """蒙版扣窗巨字（样本工艺）：幽灵回声垫底 → 字形即窗，窗内透「暗化+墨色混合」的画"""
    from PIL import ImageEnhance
    x, y, w, h = [v * dpi for v in region]
    n = len(text)
    if n == 4:
        cs = min(w / 2.4, h / 2.1); pos = [(0, 0), (1.02, 0.02), (0.30, 0.98), (1.32, 1.0)]
    elif n == 3:
        cs = min(w / 2.4, h / 2.1); pos = [(0.15, 0), (1.30, 0.06), (0.62, 1.02)]
    else:
        cs = min(w / max(n, 1) * 1.15, h); pos = [(i, 0) for i in range(n)]
    # 幽灵回声垫层（纸灰，样本同款）
    echo = Image.new("RGBA", base.size, (0, 0, 0, 0))
    de = ImageDraw.Draw(echo)
    fe = ImageFont.truetype(bake_font(font_name), int(cs * 1.06))
    off = int(0.28 * dpi)
    for ch, (cx, cy) in zip(text, pos):
        de.text((x + cx * cs + off, y + cy * cs * 0.96 + off), ch, font=fe, fill=ghost_rgba)
    base.alpha_composite(echo.filter(ImageFilter.GaussianBlur(2.2)))
    # 字形蒙版
    mask = Image.new("L", base.size, 0)
    dm = ImageDraw.Draw(mask)
    fm = ImageFont.truetype(bake_font(font_name), int(cs))
    for ch, (cx, cy) in zip(text, pos):
        dm.text((x + cx * cs, y + cy * cs * 0.96), ch, font=fm, fill=255)
    # 窗内画：画底暗化 0.5 + 墨色 0.35 混合（透画肌理）
    dark = ImageEnhance.Brightness(art_rgb).enhance(0.50).convert("RGBA")
    ink_tint = Image.new("RGBA", base.size, (26, 22, 18, 90))
    dark.alpha_composite(ink_tint)
    base.paste(dark, (0, 0), mask)

def crop_cover_fit(im, tw, th):
    sc = max(tw / im.width, th / im.height)
    im = im.resize((max(tw, int(im.width * sc)), max(th, int(im.height * sc))), Image.LANCZOS)
    return im.crop(((im.width - tw)//2, (im.height - th)//2, (im.width - tw)//2 + tw, (im.height - th)//2 + th))

def bake_strip(src, numeral, label, out, darken, dpi=200):
    """目录竖条：源图裁竖幅+暗化+巨字序号+底部衬带白标签。源图比例不限，按 skeleton 实测竖条比"""
    im = Image.open(src).convert("RGB")
    SW = 590; SH = int(SW * 1500 / 590)
    im = crop_cover_fit(im, SW, SH)
    im = ImageEnhance.Brightness(im).enhance(darken)
    strip = im.convert("RGBA")
    lay = Image.new("RGBA", (SW, SH), (0, 0, 0, 0))
    ImageDraw.Draw(lay).text((SW/2, SH*0.42), numeral, font=ImageFont.truetype(bake_font("Xingkai SC"), 520),
                             fill=(246, 241, 231, 150), anchor="mm")
    strip.alpha_composite(lay)
    strip.alpha_composite(Image.new("RGBA", (SW, 220), (30, 22, 16, 120)), (0, SH - 260))
    ImageDraw.Draw(strip).text((SW/2, SH - 150), label, font=ImageFont.truetype(bake_font("Kaiti SC"), 72),
                               fill="#F6F1E7", anchor="mm")
    strip.convert("RGB").save(out)

def bake_soft_ill(src, tw, th, radius, blur, out):
    ill = crop_cover_fit(Image.open(src).convert("RGB"), tw, th)
    mask = Image.new("L", (tw, th), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, tw, th], radius=radius, fill=255)
    mask = mask.filter(ImageFilter.GaussianBlur(blur))
    o = Image.new("RGBA", (tw, th), (0, 0, 0, 0))
    o.paste(ill, (0, 0), mask)
    o.save(out)

# ══ on_color 采样 ════════════════════════════════════════════
class Sampler:
    def __init__(self):
        self._c = {}
    def sample(self, path, box):
        if path not in self._c:
            self._c[path] = Image.open(path).convert("RGB")
        im = self._c[path]
        sx, sy = im.width / SW_IN, im.height / SH_IN
        x0, y0 = max(0, int(box[0]*sx)), max(0, int(box[1]*sy))
        x1, y1 = min(im.width, int((box[0]+box[2])*sx)), min(im.height, int((box[1]+box[3])*sy))
        region = im.crop((x0, y0, max(x0+8, x1), max(y0+8, y1))).resize((8, 8))
        px = list(region.getdata())
        return "#%02X%02X%02X" % tuple(sum(p[i] for p in px)//64 for i in range(3))

# ══ 主色自动推导（文案关键词 → 主色调） ════════════════════════
AUTO_COLOR_MAP = [
    (["科技", "代码", "智能", "数据", "AI", "插件", "引擎", "安全", "防护", "云", "互联网", "数字"], "#2B5EA7"),
    (["党政", "党建", "红", "庆典", "婚", "喜庆", "春节", "国潮"], "#A8322A"),
    (["自然", "环保", "茶", "绿", "生态", "农", "健康", "中医"], "#3E7C4F"),
    (["金融", "财富", "金", "高端", "奢侈", "商务"], "#9C7A2E"),
    (["医疗", "卫生", "清", "水"], "#2A9D8F"),
    (["教育", "培训", "校园", "课程", "文化", "书"], "#7A5C3E"),
]
def auto_primary(content):
    blob = json.dumps(content, ensure_ascii=False)
    for words, color in AUTO_COLOR_MAP:
        if any(w in blob for w in words):
            return color
    return "#5B6B7A"   # 无关键词命中：中性青灰（绝不默认照搬模板色）

def derive_palette(primary_hex, struct=None):
    """独立配色生成：只继承模板的色彩结构（各角色饱和度/明度/色相偏移关系=血统），
    色相由新内容主色决定——结构参数来自母版实测 color_structure，缺省用通用中式浅底关系"""
    h, s, v = hue_of(primary_hex)
    def hsl(hh, ss, vv):
        r, g, b = colorsys.hsv_to_rgb(hh % 1.0, max(0.0, min(1.0, ss)), max(0.0, min(1.0, vv)))
        return "#%02X%02X%02X" % (int(r*255), int(g*255), int(b*255))
    rel = struct or {}
    def role(name, dflt_s, dflt_v, dflt_hoff=0.0):
        cfg = rel.get(name, {})
        return hsl(h + cfg.get("hue_offset", dflt_hoff), cfg.get("s", dflt_s), cfg.get("v", dflt_v))
    return {
        "primary": primary_hex,
        "secondary": role("secondary", min(s * 0.58, 0.5), v * 0.85, 0.05),
        "accent": role("accent", min(s * 1.05, 1.0), min(v * 1.15, 0.85)),
        "text_dark": role("text_dark", 0.35, 0.16),
        "body_text": role("body_text", 0.32, 0.20),
        "paper": role("paper", 0.07, 0.95),
        "card_bg": role("card_bg", 0.08, 0.955),
        "cream": role("cream", 0.06, 0.965),
        "ghost": role("ghost", 0.07, 0.54),
        "ink": role("ink", 0.30, 0.20),
        "cloud": role("cloud", 0.06, 0.72),
    }

def B(x, y, w, h):
    return {"x_inch": round(x, 3), "y_inch": round(y, 3), "width_inch": round(w, 3), "height_inch": round(h, 3)}

def _lum(hexc):
    f = lambda c: (c/255/12.92) if c/255 <= 0.03928 else ((c/255+0.055)/1.055)**2.4
    r, g, b = hex2rgb(hexc)
    return 0.2126*f(r) + 0.7152*f(g) + 0.0722*f(b)

def _contrast(c1, c2):
    l1, l2 = sorted([_lum(c1), _lum(c2)], reverse=True)
    return (l1 + 0.05) / (l2 + 0.05)

def _overflow(text, box_w_in, box_h_in, size_pt):
    """与 render_ppt.py estimate_text_overflow 同口径：CJK 宽≈字号，Latin≈0.55 字号，行高 1.25"""
    width_pt, height_pt = box_w_in * 72.0, box_h_in * 72.0
    lines, cur = 1, 0.0
    for ch in str(text):
        if ch == "\n":
            lines, cur = lines + 1, 0.0
            continue
        cur += size_pt * (1.0 if ord(ch) > 0x2E7F else 0.55)
        if cur > width_pt:
            lines, cur = lines + 1, size_pt * (1.0 if ord(ch) > 0x2E7F else 0.55)
    return lines * size_pt * 1.25 > height_pt

def fit_size(text, box, size_pt, floor=13):
    """实测框内自适配字号（防溢出，下限 floor）"""
    while size_pt > floor and _overflow(text, box[2], box[3], size_pt):
        size_pt -= 1
    return size_pt

# ══ 引擎主装配方 ═════════════════════════════════════════════
_MISSING = object()

def _valid_assets(d):
    """plan 素材校验：路径不存在跳过+stderr 警告，不中断"""
    out = {}
    for k, v in (d or {}).items():
        if isinstance(v, list):
            ok = [x for x in v if os.path.isfile(x)]
            for x in v:
                if not os.path.isfile(x):
                    print(f"[plan] 路径不存在，跳过该条: {k} <- {x}", file=sys.stderr)
            if ok:
                out[k] = ok
        elif os.path.isfile(str(v)):
            out[k] = v
        else:
            print(f"[plan] 路径不存在，跳过该条: {k} <- {v}", file=sys.stderr)
    return out

def build(skel_dir, content, palette, images, work, plan=None):
    skel = json.load(open(os.path.join(skel_dir, "skeleton.json"), encoding="utf-8"))
    plan = plan or {}
    images = {**images, **_valid_assets(plan.get("global"))}     # 全局槽位覆盖
    page_plans = plan.get("pages", {})
    ctx = {"band_veil": 0.22}                                    # 横带过纱默认
    slots = {s["id"]: s for s in skel.get("asset_slots", [])}
    baked = os.path.join(work, "baked"); os.makedirs(baked, exist_ok=True)
    # ── 独立配色：色相来自内容主色，结构来自母版实测（零照搬模板色值） ──
    smp = Sampler()
    primary = palette["primary"]
    pal = derive_palette(primary, skel.get("color_structure"))
    if palette.get("secondary"): pal["secondary"] = palette["secondary"]
    if palette.get("accent"): pal["accent"] = palette["accent"]
    INK = pal["text_dark"]; BODY = pal["body_text"]; CREAM = pal["cream"]; CARD = pal["card_bg"]
    INK_BLOT = pal["ink"]; CLOUD_C = pal["cloud"]; PAPER_C = pal["paper"]
    secondary = pal["secondary"]
    warns = []
    ghost_cfg = skel["fonts"]["giant"].get("ghost_echo", {"offset_inch": 0.28, "scale": 1.06, "alpha": 170})
    ghost_rgba = tuple(hex2rgb(pal["ghost"])) + (ghost_cfg.get("alpha", 170),)
    giant_font = skel["fonts"]["giant"]["family"]

    # ── 槽位解析：用户供图 → 程序化工艺件/兜底（零原稿素材复用） ──
    _cache = {}
    def resolve(slot_id):
        if images.get(slot_id):
            s = slots.get(slot_id, {})
            if s.get("kind") == "band":
                key = f"mask_{slot_id}_{os.path.basename(str(images[slot_id]))}_{ctx['band_veil']}"   # 缓存键含源文件+alpha（页级不同图不串）
                if key not in _cache:
                    out = os.path.join(baked, f"slot_{slot_id}_{int(ctx['band_veil']*100)}_{os.path.basename(str(images[slot_id]))[:12]}.png")
                    apply_band_mask(images[slot_id], s.get("box", [0, 0, 11.7, 1.5]), out)
                    veil(Image.open(out), PAPER_C, ctx["band_veil"]).save(out)   # 横带过纱沉底（page_plan 可页级覆盖）
                    _cache[key] = out
                return _cache[key]
            return images[slot_id]
        if slot_id in _cache:
            return _cache[slot_id]
        s = slots.get(slot_id, {})
        kind = s.get("kind", slot_id)
        out = os.path.join(baked, f"slot_{slot_id}.png")
        if kind == "paper":
            gen_paper(PAPER_C, out)
        elif kind == "ink_blot":
            gen_ink_blot(INK_BLOT, out)
        elif kind == "arc":
            gen_arc(pal["ghost"], out)
        elif kind == "cloud":
            gen_cloud(CLOUD_C, out)
        elif kind == "band":
            warns.append(f"槽位 {slot_id} 无素材，已用渐变横带兜底（建议 --fetch 或显式供图）")
            gen_band(primary, secondary, out)
        elif kind == "strip_src":
            warns.append(f"目录竖条 {slot_id} 无素材，已用主色渐变兜底")
            gen_strip(primary, out)
        elif kind in ("hero_ill", "summary_ill"):
            warns.append(f"插画槽位 {slot_id} 无素材，已用柔色面板兜底")
            gen_band(primary, PAPER_C, out, w=1150, h=980)
        elif kind in ("cover_art", "summary_art"):
            return None                                   # 满版画底：无素材回退纸底（装配层处理）
        elif kind == "branch":
            warns.append("顶枝槽位无素材，本页跳过顶枝（建议供图或 --fetch）")
            return None
        else:
            return None
        _cache[slot_id] = out
        return out

    paper = resolve("paper")
    branch_path = images.get("branch") or resolve("branch")
    branch_im = None
    if branch_path:
        branch_im = crop_cover_fit(Image.open(branch_path).convert("RGB"), 1517, 470)
        _bm = Image.new("L", (1517, 470), 0)
        ImageDraw.Draw(_bm).rounded_rectangle([0, 0, 1517, 470], radius=120, fill=255)
        _bm = _bm.filter(ImageFilter.GaussianBlur(40))          # 顶枝柔边渐隐，与纸底融合
        branch_im = Image.composite(branch_im, Image.open(paper).convert("RGB").crop((1150, 0, 2667, 470)), _bm)

    def bg_paper():
        return {"type": "image", "path": paper, "fallback_color": PAPER_C, "overlay_alpha": 0}

    def txt(role, text, box, size, color, font, align="left", bold=False, on=None, bg=None):
        on = on or smp.sample(bg or paper, (box["x_inch"], box["y_inch"], box["width_inch"], box["height_inch"]))
        # 采样闭环：对比度不达标自动换 text_dark ↔ text_light 更优者（画底上文字防糊）
        need = 3.0 if (size >= 18 or bold) else 4.5
        if _contrast(color, on) < need:
            alt = max((INK, CREAM), key=lambda c: _contrast(c, on))
            if _contrast(alt, on) > _contrast(color, on):
                color = alt
        return {"type": "text", "role": role, "text": text, "box": box,
                "style": {"font": font, "size_pt": size, "bold": bold, "color": color, "align": align}, "on_color": on}

    def img(path, box, role=None, fit="stretch"):
        el = {"type": "image", "path": path, "box": box, "fit": fit}
        if role: el["role"] = role
        return el

    def seal_square(box, ch):
        """方印章：shape+白字（随主色自动统一，无需烘焙）"""
        return [{"type": "shape", "shape": "round_rect", "role": "card", "color": primary, "radius_pt": 12, "box": box},
                txt("keyword", ch, dict(box), 16, CREAM, "Kaiti SC", align="center", on=primary)]

    def seal_ink(box, ch):
        """墨点章：固定墨褐 PNG + 白字（墨色系不随主色）"""
        return [img(resolve("ink_blot"), box),
                txt("keyword", ch, dict(box), 16, CREAM, "Kaiti SC", align="center", on=INK_BLOT)]

    def seal_flower(box, num):
        """序号章：圆形 shape+数字（随辅色统一）"""
        return [{"type": "shape", "shape": "circle", "role": "card", "color": secondary, "box": box},
                txt("keyword", str(num), dict(box), 18, CREAM, "Kaiti SC", align="center", on=secondary)]

    P = skel["page_skeletons"]
    slides = []

    # ── 封面（giant_bg 工艺） ──
    cv = content["cover"]; ps = P["cover"]
    cover_bg = os.path.join(baked, "cover_bg.png")
    if images.get("cover_art"):                           # 满版画底 + 纸色纱 0.55（样本封面工艺）
        base = crop_cover_fit(Image.open(images["cover_art"]).convert("RGB"), 2667, 1500).convert("RGBA")
        base = veil(base, PAPER_C, 0.55)
    else:
        base = Image.open(paper).convert("RGBA")
    if branch_im is not None and ps.get("branch", True) and not images.get("cover_art"):
        base.paste(branch_im, tuple(ps.get("branch_xy", [1150, 0])))     # 满版画底下顶枝让位
    if images.get("cover_art"):
        draw_giant_cutout(base, crop_cover_fit(Image.open(images["cover_art"]).convert("RGB"), 2667, 1500),
                          cv["giant"], ps["giant_region"], giant_font, ghost_rgba)   # 扣窗透画
    else:
        draw_giant(base, cv["giant"], ps["giant_region"], giant_font, INK, ghost_rgba)
    base.convert("RGB").save(cover_bg)
    els = []
    if cv.get("subtitle"):
        e = ps["subtitle"]; els.append(txt("subtitle", cv["subtitle"], B(*e["box"]), e["size"], BODY, e["font"], bg=cover_bg))
    if cv.get("caption"):
        e = ps["caption"]; bx = list(e["box"])
        if cv.get("subtitle"):                       # 防叠守卫：落款永在副标之下
            sb = ps["subtitle"]["box"]
            bx[1] = max(bx[1], sb[1] + sb[3] + 0.06)
        bx[1] = min(bx[1], 7.5 - bx[3])
        els.append(txt("body", cv["caption"], B(*bx), e["size"], primary, e["font"], align="center", bg=cover_bg))
    els += seal_square(B(*ps["seal"]["box"]), cv.get("seal", "闸"))
    slides.append({"page_no": 1, "page_type": "cover",
                   "background": {"type": "image", "path": cover_bg, "fallback_color": PAPER_C, "overlay_alpha": 0},
                   "elements": els})

    # ── 目录（strips 工艺） ──
    ps = P["toc"]; toc = content["toc"]
    numerals = skel.get("toc_numerals", "壹贰叁肆")
    strip_srcs = images.get("toc_strips") or [resolve(f"toc_strip_{i}") for i in range(1, 5)]
    t = ps["title"]
    els = [txt("title", toc.get("title", "目录"), B(*t["box"]), t["size"], INK, t["font"], align="center", bold=True)]
    for i, sb in enumerate(ps["strip_boxes"]):
        sp = os.path.join(baked, f"toc_strip_{i+1}.png")
        bake_strip(strip_srcs[i % len(strip_srcs)], numerals[i], toc["labels"][i], sp, ps.get("strip_darken", 0.62))
        els.append(img(sp, B(*sb)))
    slides.append({"page_no": 2, "page_type": "toc", "background": bg_paper(), "elements": els})

    # ── 内容页 ──
    for pi, page in enumerate(content["pages"], start=3):
        pp = page_plans.get(str(pi), {})
        p_assets = _valid_assets(pp.get("assets"))
        saved = {k: images.get(k, _MISSING) for k in p_assets}
        images.update(p_assets)                                  # 局部覆盖（本页装配前）
        ctx["band_veil"] = pp.get("fusion", {}).get("band_veil", 0.22)
        variant = page["variant"]; ps = P[variant]; items = page["items"]
        els = []
        # 页标题（标题艺术条 → 文本渲染，坐标/字号实测）
        t = ps["title"]
        els.append(txt("title", page["title"], B(*t["box"]), t["size"], INK, t["font"], bold=True))
        if ps.get("arc"):
            els.append(img(resolve("arc"), B(*ps["arc"]["box"]), role="decoration"))
        # 横带
        if ps.get("band"):
            els.append(img(resolve(ps["band"]["asset"]), B(*ps["band"]["box"])))
        g = ps.get("group", {})                         # 组骨架：seal/head/body 相对坐标
        n = len(items)
        positions = ps.get("positions") or []           # 实测组原点列表（模板组数）
        if n <= len(positions):
            pos_use = positions[:n]
        elif variant == "hero_blocks":                  # 角块超出 4 角 → 底部居中（锚定低带上沿防叠）
            band_top = ps["band"]["box"][1] if ps.get("band") else 6.6
            ill_bot = ps["ill"]["box"][1] + ps["ill"]["box"][3] if ps.get("ill") else 5.1
            extra_y = min(max(ill_bot + 0.12, 4.6), band_top - 1.15)
            pos_use = positions + [[4.35, round(extra_y, 2)]] * (n - len(positions))
        elif positions:                                 # 组数超模板：3列错落网格（横带下方起排，防越界）
            band_bottom = (ps["band"]["box"][1] + ps["band"]["box"][3] + 0.15) if ps.get("band") else 2.2
            pos_use = []
            for i in range(n):
                r, c = divmod(i, 3)
                pos_use.append([0.95 + c * 4.0 + (1.0 if r % 2 else 0), band_bottom + r * 1.85])
        else:
            pos_use = []
        if variant in ("seal_blocks", "plum_blocks"):
            for i, it in enumerate(items):
                x, y = pos_use[i][0], pos_use[i][1]
                body_h = min(g["body"][3], 7.45 - (y + g["body"][1]))   # 画布下缘硬钳制
                els += seal_ink(B(x + g["seal"][0], y + g["seal"][1], g["seal"][2], g["seal"][3]), it.get("seal", "·"))
                els.append(txt("keyword", it["head"], B(x + g["head"][0], y + g["head"][1], g["head"][2], g["head"][3]),
                               g["head_size"], INK, t["font"], bold=True))
                els.append(txt("body", it["body"], B(x + g["body"][0], y + g["body"][1], g["body"][2], body_h),
                               g["body_size"], BODY, skel["fonts"]["body"]["family"]))
        elif variant in ("cards", "scroll_cards"):
            cw, ch, y0 = ps["card"]["w"], ps["card"]["h"], ps["card"]["y"]
            gap = ps["card"]["gap"]
            cw2 = min(cw, (11.73 - (n - 1) * gap) / n)
            x_start = (SW_IN - (cw2 * n + gap * (n - 1))) / 2
            for i, it in enumerate(items):
                x = x_start + i * (cw2 + gap)
                els.append({"type": "shape", "shape": "round_rect", "role": "card", "color": CARD,
                            "radius_pt": ps["card"].get("radius_pt", 10), "box": B(x, y0, cw2, ch)})
                els += seal_flower(B(x + cw2/2 - ps["seal_d"]/2, y0 - ps["seal_d"]/2 + ps.get("seal_dy", 0.05), ps["seal_d"], ps["seal_d"]), i + 1)
                els.append(txt("keyword", it["head"], B(x + 0.15, y0 + ps["head_dy"], cw2 - 0.3, 0.5),
                               ps["head_size"], INK, t["font"], align="center", bold=True, on=CARD))
                els.append(txt("body", it["body"], B(x + 0.16, y0 + ps["body_dy"], cw2 - 0.32, ps.get("body_h", 1.35)),
                               g["body_size"], BODY, skel["fonts"]["body"]["family"], align="center", on=CARD))
        elif variant == "hero_blocks":
            hero_src = images.get("hero_ill") or resolve("hero_ill")
            hp = os.path.join(baked, "hero_ill.png")
            bake_soft_ill(hero_src, int(ps["ill"]["box"][2]*200), int(ps["ill"]["box"][3]*200), 70, 22, hp)
            els.append(img(hp, B(*ps["ill"]["box"])))
            for i, (it, (x, y)) in enumerate(zip(items, pos_use)):
                els.append(txt("keyword", it["head"], B(x, y, g["head"][2], g["head"][3]), g["head_size"], INK, t["font"], bold=True))
                els.append(txt("body", it["body"], B(x, y + g["body_dy"], g["body"][2], g["body"][3]), g["body_size"], BODY, skel["fonts"]["body"]["family"]))
        else:
            raise SystemExit(f"未知 variant: {variant}")
        for deco in ps.get("clouds", []):               # 祥云角饰（装饰计数由骨架保证 ≤3）
            els.append(img(resolve(deco["asset"]), B(*deco["box"]), role="decoration"))
        slides.append({"page_no": pi, "page_type": "content", "background": bg_paper(), "elements": els})
        for k, v in saved.items():                               # 局部覆盖恢复
            if v is _MISSING:
                images.pop(k, None)
            else:
                images[k] = v
        ctx["band_veil"] = 0.22

    # ── 封底（giant_bg + 右侧插画槽） ──
    sm = content["summary"]; ps = P["summary"]
    summ_bg = os.path.join(baked, "summary_bg.png")
    if images.get("summary_art"):                       # 封底满版画底 + 纱
        base = crop_cover_fit(Image.open(images["summary_art"]).convert("RGB"), 2667, 1500).convert("RGBA")
        base = veil(base, PAPER_C, 0.55)
    else:
        base = Image.open(paper).convert("RGBA")
    if branch_im is not None and ps.get("branch", True): base.paste(branch_im, tuple(ps.get("branch_xy", [1150, 0])))
    s_ill = None if images.get("summary_art") else (images.get("summary_ill") or resolve("summary_ill"))  # 满版画底下插画让位
    ib = ps["ill"]["box"]
    if s_ill is None:
        ill = None
    else:
        ill = crop_cover_fit(Image.open(s_ill).convert("RGB"), int(ib[2]*200), int(ib[3]*200))
    mask = Image.new("L", (int(ib[2]*200), int(ib[3]*200)), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, int(ib[2]*200), int(ib[3]*200)], radius=90, fill=255)
    mask = mask.filter(ImageFilter.GaussianBlur(26))
    if ill is not None:
        base.paste(ill, (int(ib[0]*200), int(ib[1]*200)), mask)
    if images.get("summary_art"):
        draw_giant_cutout(base, crop_cover_fit(Image.open(images["summary_art"]).convert("RGB"), 2667, 1500),
                          sm["giant"], ps["giant_region"], giant_font, ghost_rgba)
    else:
        draw_giant(base, sm["giant"], ps["giant_region"], giant_font, INK, ghost_rgba)
    base.convert("RGB").save(summ_bg)
    els = []
    e = ps["body"]; els.append(txt("body", sm["body"], B(*e["box"]), fit_size(sm["body"], e["box"], e["size"]), BODY, e["font"], bg=summ_bg))
    e = ps["caption"]; bx = list(e["box"])          # 防叠守卫：落款永在副文之下
    bb = ps["body"]["box"]
    bx[1] = min(max(bx[1], bb[1] + bb[3] + 0.05), 7.5 - bx[3])
    els.append(txt("body", sm["caption"], B(*bx), e["size"], primary, e["font"], align="right", bg=summ_bg))
    els += seal_square(B(*ps["seal"]["box"]), sm.get("seal", "定"))
    slides.append({"page_no": len(slides)+1, "page_type": "summary",
                   "background": {"type": "image", "path": summ_bg, "fallback_color": PAPER_C, "overlay_alpha": 0},
                   "elements": els})

    ri = {"meta": {"slide_width_inch": SW_IN, "slide_height_inch": SH_IN, "default_font": skel["fonts"]["body"]["family"]},
          "theme": {"colors": {"primary": primary, "secondary": [secondary, pal["cloud"]], "accent": pal["accent"],
                               "text_dark": INK, "text_light": CREAM, "background": PAPER_C},
                    "fonts": {"title": skel["fonts"]["page_title"]["family"], "subtitle": skel["fonts"]["caption"]["family"],
                              "body": skel["fonts"]["body"]["family"]}},
          "asset_base_path": "/", "slides": slides}
    for w in warns:
        print("[slot]", w, file=sys.stderr)
    ri_path = os.path.join(work, "render_input.json")
    with open(ri_path, "w", encoding="utf-8") as f:
        json.dump(ri, f, ensure_ascii=False, indent=1)
    return ri_path

def main(argv=None, skeleton_dir=None):
    ap = argparse.ArgumentParser(description="龙骨包通用引擎")
    ap.add_argument("--skeleton", default=skeleton_dir, help="龙骨包目录（含 skeleton.json）")
    ap.add_argument("--content"); ap.add_argument("--demo", action="store_true")
    ap.add_argument("--primary", default="auto", help="主色调 HEX 或 auto（由文案关键词推导）")
    ap.add_argument("--secondary", default=None); ap.add_argument("--accent", default=None)
    ap.add_argument("--toc-images"); ap.add_argument("--hero-ill"); ap.add_argument("--summary-ill")
    ap.add_argument("--paper"); ap.add_argument("--branch")
    ap.add_argument("--fetch", default=None, help="素材检索词：自动下载 4竖条+hero+封底插图")
    ap.add_argument("--plan", default=None, help="page_plan.json：逐页设计产物（global/pages 槽位覆盖+fusion 参数）")
    ap.add_argument("--out", default=None); ap.add_argument("--work", default=None)
    args = ap.parse_args(argv)
    skel_dir = args.skeleton or os.path.dirname(os.path.abspath(__file__))
    skel = json.load(open(os.path.join(skel_dir, "skeleton.json"), encoding="utf-8"))
    if args.demo:
        args.content = os.path.join(skel_dir, "examples", "content.example.json")
    if not args.content:
        ap.error("--content 或 --demo 必填")
    content = json.load(open(args.content, encoding="utf-8"))
    primary = auto_primary(content) if args.primary in (None, "auto") else args.primary
    palette = {"primary": primary, "secondary": args.secondary, "accent": args.accent}
    images = {}
    if args.fetch:
        n_bands = len([s for s in skel.get("asset_slots", []) if s.get("kind") == "band"])
        # asset spider removed in public release; supply your own asset fetcher
        got = []
        if len(got) >= 4: images["toc_strips"] = got[:4]
        if len(got) >= 5: images["hero_ill"] = got[4]
        if len(got) >= 6: images["summary_ill"] = got[5]
        for bi in range(n_bands):                        # 横带槽位接续分配
            gi = 6 + bi
            if len(got) > gi: images[f"band_{bi}"] = got[gi]
        # 满版画底与横带同族（素材同族规则）：复用横带首图，规避拼图/水印尾图；顶枝默认弃用（硬贴违和）
        if n_bands and len(got) > 6:
            images["cover_art"] = images.get("cover_art") or got[6]
            images["summary_art"] = images.get("summary_art") or got[6]
    if args.toc_images: images["toc_strips"] = [p.strip() for p in args.toc_images.split(",")][:4]
    for k, v in [("hero_ill", args.hero_ill), ("summary_ill", args.summary_ill), ("paper", args.paper), ("branch", args.branch)]:
        if v: images[k] = v
    work = args.work or tempfile.mkdtemp(prefix="deck_")
    os.makedirs(work, exist_ok=True)
    plan = None
    if args.plan:
        plan = json.load(open(args.plan, encoding="utf-8"))
    ri = build(skel_dir, content, palette, images, work, plan=plan)
    out = args.out or os.path.join(os.getcwd(), "deck_output.pptx")
    r = subprocess.run([sys.executable, RENDERER, ri], capture_output=True, text=True)
    print(r.stdout.strip())
    if r.returncode != 0:
        print(r.stderr[-500:], file=sys.stderr); sys.exit(r.returncode)
    subprocess.run(["cp", os.path.join(work, "output.pptx"), out], check=True)
    print(json.dumps({"status": "DELIVERED", "pptx": out, "primary": palette["primary"], "work": work}, ensure_ascii=False))

if __name__ == "__main__":
    main()

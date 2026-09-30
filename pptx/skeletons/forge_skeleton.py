#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
forge_skeleton.py — PPT 复刻母版脚本（生成器）
输入：被复刻的 .pptx
分析：①逐页 XML 解剖（元素类型/坐标/字号/字体/颜色，真文本 vs 烘焙图）
     ②逐页像素聚类（主色调占比/明暗基调） ③素材分类（横带/角饰/印章/插画/画底）
     ④巨字烘焙区暗色簇定位 ⑤纸底自动裁切
输出：龙骨包目录 = skeleton.json（龙骨+主色调+图片需求） + assets/ + build_deck.py（复刻脚本） + examples/content.example.json + README.md
之后出片只需：文案 JSON + 主色调(auto可推导) + 素材图（可 --fetch 自动下载）→ build_deck.py 一键生成

用法：
  python3 forge_skeleton.py --pptx /path/模板.pptx --name my_skel --out /path/outdir
"""
import argparse, colorsys, json, os, re, shutil, subprocess, sys, zipfile
import numpy as np
from PIL import Image

EMU = 914400
CANVAS_W, CANVAS_H = 13.333, 7.5
ENGINE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_engine")
SOFFICE = os.environ.get("SOFFICE_PATH") or shutil.which("soffice") or "/Applications/LibreOffice.app/Contents/MacOS/soffice"

def hexs(c):
    return "#%02X%02X%02X" % tuple(int(v) for v in c)

def hex2rgb(h):
    h = str(h).lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4)) if len(h) == 6 else (0, 0, 0)

def hue_dist_red(hexc):
    r, g, b = [v/255 for v in hex2rgb(hexc)]
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    dh = min(h, 1 - h)
    return dh, s, v

def lum(hexc):
    r, g, b = hex2rgb(hexc)
    f = lambda c: (c/255/12.92) if c/255 <= 0.03928 else ((c/255+0.055)/1.055)**2.4
    return 0.2126*f(r) + 0.7152*f(g) + 0.0722*f(b)

# ══ 1. pptx → 分页 PNG + media 解包 ══════════════════════════
def rasterize(pptx, work):
    out = os.path.join(work, "pages"); os.makedirs(out, exist_ok=True)
    pdf = None
    for attempt in range(3):                            # soffice profile 锁竞争重试
        profile = os.path.join(work, f".lo_profile_{attempt}")
        subprocess.run([SOFFICE, f"-env:UserInstallation=file://{profile}",
                        "--headless", "--convert-to", "pdf", "--outdir", out, pptx],
                       capture_output=True, timeout=300)
        pdf = next((f for f in os.listdir(out) if f.endswith(".pdf")), None)
        if pdf:
            break
    if not pdf:
        raise SystemExit("soffice 转换 PDF 失败（3 次重试）")
    import fitz
    doc = fitz.open(os.path.join(out, pdf))
    paths = []
    for i, page in enumerate(doc, 1):
        p = os.path.join(out, f"page_{i:02d}.png")
        page.get_pixmap(dpi=110).save(p)
        paths.append(p)
    os.remove(os.path.join(out, pdf))
    return paths

def unzip_media(pptx, work):
    dst = os.path.join(work, "media"); os.makedirs(dst, exist_ok=True)
    with zipfile.ZipFile(pptx) as z:
        for n in z.namelist():
            if n.startswith("ppt/media/") and not n.endswith("/"):
                with open(os.path.join(dst, os.path.basename(n)), "wb") as f:
                    f.write(z.read(n))
    return dst

# ══ 2. XML 解剖 ══════════════════════════════════════════════
def anatomy(pptx):
    from pptx import Presentation
    prs = Presentation(pptx)
    slides = []
    for s in prs.slides:
        els = []
        for sh in s.shapes:
            box = [round(sh.left/EMU, 3), round(sh.top/EMU, 3), round(sh.width/EMU, 3), round(sh.height/EMU, 3)]
            el = {"box": box}
            st = str(sh.shape_type)
            if "PICTURE" in st:
                el["type"] = "picture"
                try:
                    el["image"] = sh.image.filename or ""
                    el["blob_size"] = len(sh.image.blob)
                except Exception:
                    pass
            elif "TEXT_BOX" in st and sh.has_text_frame and sh.text_frame.text.strip():
                el["type"] = "text"
                el["text"] = sh.text_frame.text.strip()
                for para in sh.text_frame.paragraphs:
                    if para.runs:
                        r = para.runs[0]
                        el["font"] = r.font.name or ""
                        el["size"] = r.font.size.pt if r.font.size else 14
                        el["bold"] = bool(r.font.bold)
                        try: el["color"] = str(r.font.color.rgb)
                        except Exception: el["color"] = "2C1E12"
                        break
            elif "AUTO_SHAPE" in st:
                el["type"] = "shape"
                try:
                    from pptx.enum.dml import MSO_FILL_TYPE
                    if sh.fill.type == MSO_FILL_TYPE.SOLID:
                        el["fill"] = str(sh.fill.fore_color.rgb)
                except Exception:
                    pass
            else:
                continue
            els.append(el)
        slides.append(els)
    return slides

# ══ 3. 像素分析（主色调占比 + 明暗基调 + 巨字区 + 纸底裁切） ═══
def pixel_profile(page_png, k=6):
    from sklearn.cluster import MiniBatchKMeans
    im = Image.open(page_png).convert("RGB").resize((320, 180))
    px = np.asarray(im).reshape(-1, 3).astype(float)
    km = MiniBatchKMeans(n_clusters=k, random_state=42, batch_size=4096).fit(px)
    counts = np.bincount(km.labels_, minlength=k)
    order = np.argsort(-counts)
    dist = [{"hex": hexs(km.cluster_centers_[o]), "ratio": round(float(counts[o])/counts.sum(), 3)} for o in order]
    mean_lum = float(np.mean(0.2126*px[:, 0]/255 + 0.7152*px[:, 1]/255 + 0.0722*px[:, 2]/255))
    return {"distribution": dist, "mean_lum": round(mean_lum, 3)}

def detect_giant_region(bg_png):
    """全出血画底上的烘焙巨字暗色簇定位 → [x,y,w,h] inch（膨胀合并笔画成簇，排除贴顶边装饰簇）"""
    im = Image.open(bg_png).convert("L")
    arr = np.asarray(im).astype(float)/255
    mask = arr < 0.32
    from scipy import ndimage
    mask = ndimage.binary_dilation(mask, iterations=18)   # 合并同字/同题笔画
    lab, n = ndimage.label(mask)
    if n == 0:
        return [0.8, 1.0, 5.0, 4.5]
    H, W = arr.shape
    cands = []
    for c in range(1, n+1):
        ys, xs = np.where(lab == c)
        if len(xs) < 200:
            continue
        cands.append((len(xs), xs.min(), xs.max(), ys.min(), ys.max()))
    if not cands:
        return [0.8, 1.0, 5.0, 4.5]
    cands.sort(reverse=True)
    inner = [c for c in cands if c[3] > H*0.05]            # 排除贴顶边（枝桠/横带类装饰）
    _, x0, x1, y0, y1 = (inner or cands)[0]
    pad = 0.3
    xi0, yi0 = max(0, x0/W*CANVAS_W-pad), max(0, y0/H*CANVAS_H-pad)
    xi1, yi1 = min(CANVAS_W, x1/W*CANVAS_W+pad), min(CANVAS_H, y1/H*CANVAS_H+pad)
    return [round(xi0, 2), round(yi0, 2), round(xi1-xi0, 2), round(yi1-yi0, 2)]

def crop_paper_base(bg_png, out):
    """最大低方差空白区 → 纸底"""
    im = Image.open(bg_png).convert("RGB")
    W, H = im.size
    arr = np.asarray(im.resize((W//4, H//4))).astype(float)
    best, bs = None, 1e18
    wh, ww = arr.shape[0]//3, arr.shape[1]//3
    for y in range(0, arr.shape[0]-wh, wh//2):
        for x in range(0, arr.shape[1]-ww, ww//2):
            s = arr[y:y+wh, x:x+ww].std()
            if s < bs:
                bs, best = s, (x*4, y*4, ww*4, wh*4)
    x, y, w, h = best
    im.crop((x, y, x+w, y+h)).resize((2667, 1500), Image.LANCZOS).save(out)

# ══ 4. 素材槽位分类（零拷贝铁律：只记规格，原稿媒体一个文件都不进包） ════════
def classify_pictures(slides):
    """横带(宽>60%页宽,高<35%) / 角饰(角落小图) / 墨点章(小方图+单字白文本重叠) / 笔触弧 / 画底(全页，另处理)"""
    assets = {"bands": [], "clouds": [], "ink_blot": None, "flower_seal": None, "arc": None, "branch": None}
    seen = {}
    for els in slides:
        texts = [e for e in els if e["type"] == "text"]
        for e in els:
            if e["type"] != "picture":
                continue
            x, y, w, h = e["box"]
            key = (e.get("image"), e.get("blob_size"))
            if key in seen:
                continue
            seen[key] = e
            if w >= CANVAS_W*0.9 and h >= CANVAS_H*0.9:
                continue                                    # 画底单独处理
            if w > CANVAS_W*0.6 and h < CANVAS_H*0.35:
                bi = len(assets["bands"])
                assets["bands"].append({"key": f"band_{bi}", "box": e["box"]})
            elif 0.4 <= w <= 1.0 and abs(w-h) < 0.25:
                overlap = any(t["type"] == "text" and len(t["text"]) == 1 and t.get("color") == "FFFFFF"
                              and abs(t["box"][0]-x) < 0.3 and abs(t["box"][1]-y) < 0.3 for t in texts)
                if overlap and assets["ink_blot"] is None:
                    assets["ink_blot"] = {"box": e["box"]}
                elif assets["arc"] is None and y < 1.2:
                    assets["arc"] = {"box": e["box"]}
            elif 1.0 <= w <= 3.0 and h < 1.3 and (x < 1.0 or x > 10.0):
                idx = len(assets["clouds"])
                assets["clouds"].append({"key": f"cloud_{idx+1}", "box": e["box"]})
    return assets

def find_media(media_dir, el):
    """按 blob 大小匹配 media 文件"""
    for f in sorted(os.listdir(media_dir)):
        p = os.path.join(media_dir, f)
        if os.path.getsize(p) == el.get("blob_size"):
            return p
    return None

# ══ 5. 页面配方分类 ══════════════════════════════════════════
def classify_page(els, page_index=0):
    pics = [e for e in els if e["type"] == "picture"]
    texts = [e for e in els if e["type"] == "text"]
    shapes = [e for e in els if e["type"] == "shape"]
    full_bg = [p for p in pics if p["box"][2] >= CANVAS_W*0.9 and p["box"][3] >= CANVAS_H*0.9]
    strips = [p for p in pics if p["box"][3] > CANVAS_H*0.7 and p["box"][3]/max(p["box"][2], 0.1) > 2.0 and p["box"][2] < 5]
    cards = [s for s in shapes if 1.5 < s["box"][2] < 4.5 and 1.5 < s["box"][3] < 4.5]
    heads = [t for t in texts if t.get("bold") and 15 <= t.get("size", 0) <= 22 and 2 <= len(t["text"]) <= 6]
    seal_chars = [t for t in texts if len(t["text"]) == 1 and t.get("color") == "FFFFFF"]
    big_ill = [p for p in pics if p["box"][2] > 3 and p["box"][3] > 3 and p not in full_bg]
    if len(strips) >= 3:
        return "toc"
    if len(cards) >= 3:
        return "scroll_cards"
    if full_bg and len(heads) >= 4:
        return "hero_blocks"
    if full_bg and big_ill:
        return "summary"
    if full_bg and page_index == 0:
        return "cover"          # 首页全出血画底 → 封面（含印章字）
    if full_bg:
        return "chapter"        # 后续全出血画底+章字 → 章页（记录不工程化）
    if len(seal_chars) >= 2 or len(heads) >= 2:
        return "seal_blocks"
    return "custom"

# ══ 主流程 ═══════════════════════════════════════════════════
def forge(pptx, name, out_dir, style_label=None):
    work = os.path.join(out_dir, ".forge_work"); os.makedirs(work, exist_ok=True)
    os.makedirs(os.path.join(out_dir, "examples"), exist_ok=True)

    pages_png = rasterize(pptx, work)
    media_dir = unzip_media(pptx, work)          # 仅用于分析，收尾即焚，零文件入包
    slides = anatomy(pptx)
    profiles = [pixel_profile(p) for p in pages_png]
    pic_assets = classify_pictures(slides)

    # ── 色彩结构提取（只取 s/v/色相偏移=血统关系；实测 HEX 一律不写入包） ──
    all_text_colors, shape_fills = [], []
    for els in slides:
        for e in els:
            if e["type"] == "text" and e.get("color"):
                all_text_colors.append("#" + e["color"])
            if e["type"] == "shape" and e.get("fill"):
                shape_fills.append(("#" + e["fill"], e["box"]))
    darks = [c for c in all_text_colors if lum(c) < 0.15]
    reds = [c for c in set(all_text_colors) if hue_dist_red(c)[0] < 0.06 and hue_dist_red(c)[1] > 0.4]
    primary_m = reds[0] if reds else None
    text_dark_m = max(set(darks), key=darks.count) if darks else None
    light_pages = [p for p in profiles if p["mean_lum"] > 0.75]
    paper_m = light_pages[0]["distribution"][0]["hex"] if light_pages else None
    card_fills = [c for c, b in shape_fills if 1.5 < b[2] < 4.5 and 1.5 < b[3] < 4.5]
    # 注：方印章 shape fill 属 primary 载体（非墨色），不纳入 ink 角色；墨点章是图片无 XML fill，ink 走引擎默认深墨结构
    def _hsv(hexc):
        r, g, b = [v/255 for v in hex2rgb(hexc)]
        return colorsys.rgb_to_hsv(r, g, b)
    h0 = _hsv(primary_m)[0] if primary_m else 0.0
    def _struct(hexc):
        if not hexc:
            return None
        h, s, v = _hsv(hexc)
        return {"s": round(s, 3), "v": round(v, 3), "hue_offset": round((h - h0) % 1.0, 3)}
    color_structure = {}
    for k, m in [("text_dark", text_dark_m), ("body_text", text_dark_m),
                 ("paper", paper_m), ("card_bg", card_fills[0] if card_fills else None)]:
        st = _struct(m)
        if st:
            color_structure[k] = st

    # ── 画底分析（仅用于巨字区定位；纸底/顶枝一律不裁不入包） ──
    full_bg_pages = {}
    for i, els in enumerate(slides):
        for e in els:
            if e["type"] == "picture" and e["box"][2] >= CANVAS_W*0.9 and e["box"][3] >= CANVAS_H*0.9:
                src = find_media(media_dir, e)
                if src:
                    full_bg_pages[i] = src
    giant_regions = {i: detect_giant_region(src) for i, src in full_bg_pages.items()}

    # ── 逐页配方生成 ──
    skels = {}
    content_pages, example_pages = [], []
    toc_info, cover_info, summary_info = None, None, None
    for i, els in enumerate(slides):
        recipe = classify_page(els, page_index=i)
        texts = [e for e in els if e["type"] == "text"]
        pics = [e for e in els if e["type"] == "picture"]
        shapes = [e for e in els if e["type"] == "shape"]
        def first(pred, default=None):
            for t in texts:
                if pred(t): return t
            return default
        if recipe == "toc":
            title = first(lambda t: t.get("bold") and t.get("size", 0) >= 30, {"box": [5.67, 0.38, 2.0, 0.85], "size": 40, "font": "Songti SC"})
            strip_boxes = sorted([p["box"] for p in pics if p["box"][3] > CANVAS_H*0.7 and p["box"][2] < 5], key=lambda b: b[0])
            toc_info = {"title": {"box": title["box"], "size": min(54, int(title.get("size", 40))), "font": norm_font(title.get("font"))},
                        "strip_boxes": strip_boxes[:4], "strip_darken": 0.62}
        elif recipe == "cover":
            sub = first(lambda t: 18 <= t.get("size", 0) <= 22 and t["box"][1] > 5, None)
            cap = first(lambda t: t.get("color") and hue_dist_red("#"+t["color"])[1] > 0.4 and t["box"][1] > 5.5, None)
            seal = first(lambda t: len(t["text"]) == 1 and t.get("color") == "FFFFFF", None)
            cover_info = {"giant_region": giant_regions.get(i, [0.9, 1.1, 5.2, 4.6]),
                          "subtitle": {"box": (sub or {"box": [0.92, 6.30, 6.5, 0.5]})["box"], "size": int((sub or {}).get("size", 18)), "font": norm_font((sub or {}).get("font"))},
                          "caption": {"box": (cap or {"box": [4.67, 6.93, 4.0, 0.5]})["box"], "size": int((cap or {}).get("size", 16)), "font": norm_font((cap or {}).get("font") or "Kaiti SC")},
                          "seal": {"box": (seal or {"box": [12.3, 6.35, 0.6, 0.6]})["box"]}, "branch": True, "branch_xy": [1150, 0],
                          "_cover_texts": {"subtitle": (sub or {}).get("text", ""), "caption": (cap or {}).get("text", ""), "seal": (seal or {}).get("text", "印")}}
        elif recipe == "summary":
            body = first(lambda t: 14 <= t.get("size", 0) <= 22 and len(t["text"]) > 8, None)
            cap = first(lambda t: t.get("color") and hue_dist_red("#"+t["color"])[1] > 0.4, None)
            seal = first(lambda t: len(t["text"]) == 1 and t.get("color") == "FFFFFF", None)
            ill = max([p for p in pics if p["box"][2] > 3 and p["box"][3] > 3 and p["box"][2] < 12], key=lambda p: p["box"][2]*p["box"][3], default=None)
            summary_info = {"giant_region": giant_regions.get(i, [0.8, 0.9, 4.8, 4.8]),
                            "ill": {"box": (ill or {"box": [6.9, 1.2, 5.6, 4.2]})["box"]},
                            "body": {"box": (body or {"box": [6.95, 6.02, 5.85, 1.0]})["box"], "size": min(20, int((body or {}).get("size", 14))), "font": norm_font((body or {}).get("font"))},
                            "caption": {"box": (cap or {"box": [9.35, 7.0, 3.2, 0.5]})["box"], "size": int((cap or {}).get("size", 16)), "font": norm_font((cap or {}).get("font") or "Kaiti SC")},
                            "seal": {"box": (seal or {"box": [0.85, 6.30, 0.6, 0.6]})["box"]}, "branch": True, "branch_xy": [1150, 0],
                            "_texts": {"body": (body or {}).get("text", ""), "caption": (cap or {}).get("text", ""), "seal": (seal or {}).get("text", "印")}}
        elif recipe in ("seal_blocks", "scroll_cards", "hero_blocks"):
            title_pic = next((p for p in pics if 3 < p["box"][2] < 8 and 0.3 < p["box"][3] < 1.2 and p["box"][1] < 1.2), None)
            tb = title_pic["box"] if title_pic else [0.9, 0.32, 8.5, 0.95]
            title = {"box": [tb[0], tb[1], max(tb[2], 5.4), max(tb[3], 0.72)],   # 文本渲染防溢出：box 适配 36pt 单行
                     "size": 36, "font": "Songti SC"}
            band = next((p for p in pics if p["box"][2] > CANVAS_W*0.6 and p["box"][3] < CANVAS_H*0.35), None)
            band_info = None
            if band:
                ba = next((a for a in pic_assets["bands"] if abs(a["box"][1]-band["box"][1]) < 0.3), None)
                if ba is None and pic_assets["bands"]:
                    ba = pic_assets["bands"][0]
                if ba:
                    band_info = {"asset": ba["key"], "box": band["box"]}
            arc = next((p for p in pics if 0.4 <= p["box"][2] <= 1.0 and p["box"][1] < 1.2 and p is not title_pic), None)
            clouds = []
            for ci, ca in enumerate(pic_assets["clouds"][:2]):
                clouds.append({"asset": f"cloud_{ci+1}", "box": [11.85, 0.42, 1.3, 0.68] if ci == 0 else [0.35, 6.35, 1.15, 0.6]})
            heads = sorted([t for t in texts if t.get("bold") and 14 <= t.get("size", 0) <= 22 and 2 <= len(t["text"]) <= 8], key=lambda t: (t["box"][1], t["box"][0]))
            bodies = sorted([t for t in texts if not t.get("bold") and len(t["text"]) > 6], key=lambda t: (t["box"][1], t["box"][0]))
            seal_chars = sorted([t for t in texts if len(t["text"]) == 1 and t.get("color") == "FFFFFF"], key=lambda t: (t["box"][1], t["box"][0]))
            head_size = int(heads[0]["size"]) if heads else 18
            body_size = int(bodies[0]["size"]) if bodies else 14
            items_ex = []
            if recipe == "scroll_cards":
                cards = sorted([s for s in shapes if 1.5 < s["box"][2] < 4.5 and 1.5 < s["box"][3] < 4.5], key=lambda s: s["box"][0])
                c0 = cards[0]["box"]
                gap = round(cards[1]["box"][0] - cards[0]["box"][0] - c0[2], 2) if len(cards) > 1 else 0.25
                seal_pics = sorted([p for p in pics if 0.4 <= p["box"][2] <= 1.0 and abs(p["box"][2]-p["box"][3]) < 0.2 and p["box"][1] > 2], key=lambda p: p["box"][0])
                head0 = heads[0]["box"] if heads else [c0[0]+0.15, c0[1]+0.62, c0[2]-0.3, 0.5]
                body0 = bodies[0]["box"] if bodies else [c0[0]+0.16, c0[1]+1.22, c0[2]-0.32, 1.35]
                skel_page = {"title": title, "arc": {"box": arc["box"]} if arc else None, "band": band_info,
                             "card": {"w": c0[2], "h": c0[3], "y": c0[1], "gap": gap, "radius_pt": 10},
                             "seal_d": seal_pics[0]["box"][2] if seal_pics else 0.75,
                             "seal_dy": (seal_pics[0]["box"][1] - c0[1] + (seal_pics[0]["box"][2])/2) if seal_pics else 0.05,
                             "head_dy": round(head0[1] - c0[1], 2), "body_dy": round(body0[1] - c0[1], 2),
                             "body_h": body0[3], "head_size": head_size, "group": {"body_size": body_size},
                             "clouds": clouds[:1]}
                for j, hd in enumerate(heads):
                    items_ex.append({"head": hd["text"], "body": bodies[j]["text"] if j < len(bodies) else ""})
            elif recipe == "hero_blocks":
                pos = [[h["box"][0], h["box"][1]] for h in heads]
                h0, b0 = (heads[0], bodies[0]) if heads and bodies else (None, None)
                skel_page = {"title": title, "band": band_info,
                             "ill": {"box": [4.22, 1.55, 4.9, 3.6]},
                             "positions": pos,
                             "group": {"head": [0, 0, 2.7, 0.5], "head_size": head_size,
                                       "body": [0, 0, 2.55, 1.35], "body_size": body_size,
                                       "body_dy": round((b0["box"][1] - h0["box"][1]), 2) if h0 and b0 else 0.65},
                             "clouds": clouds[:1]}
                for j, hd in enumerate(heads):
                    items_ex.append({"head": hd["text"], "body": bodies[j]["text"] if j < len(bodies) else ""})
            else:  # seal_blocks
                pos = [[s["box"][0], s["box"][1]] for s in seal_chars]
                h0, b0 = (heads[0], bodies[0]) if heads and bodies else (None, None)
                s0 = seal_chars[0] if seal_chars else None
                skel_page = {"title": title, "arc": {"box": arc["box"]} if arc else None, "band": band_info,
                             "positions": pos,
                             "group": {"seal": [0, 0, s0["box"][2], s0["box"][3]] if s0 else [0, 0, 0.8, 0.8],
                                       "head": [round(h0["box"][0]-s0["box"][0], 2), round(h0["box"][1]-s0["box"][1], 2), h0["box"][2], h0["box"][3]] if h0 and s0 else [1.0, 0.08, 2.5, 0.55],
                                       "body": [round(b0["box"][0]-s0["box"][0], 2), round(b0["box"][1]-s0["box"][1], 2), b0["box"][2], b0["box"][3]] if b0 and s0 else [0, 0.92, 3.45, 1.25],
                                       "head_size": head_size, "body_size": body_size},
                             "clouds": clouds[:2]}
                for j, sc in enumerate(seal_chars):
                    items_ex.append({"seal": sc["text"], "head": heads[j]["text"] if j < len(heads) else "",
                                     "body": bodies[j]["text"] if j < len(bodies) else ""})
            # 引擎可消费的 variant 名归一
            variant = recipe
            skels.setdefault(variant, skel_page)
            content_pages.append({"page_index": i+1, "variant": variant})
            example_pages.append({"variant": variant, "title": "页面标题 · 占位", "items": items_ex})
        else:
            content_pages.append({"page_index": i+1, "variant": "custom", "note": "母版未工程化，已记录"})

    # ── 素材槽位规格（零原稿素材：只落规格与检索词，不出任何媒体文件） ──
    CQ = "<内容主色>色的<风格>"
    asset_slots = [
        {"id": "paper", "kind": "paper", "spec": "程序化宣纸底 2667x1500（底色由内容主色派生+纤维噪点+晕影），可被 --paper 替换"},
        {"id": "branch", "kind": "branch", "spec": "顶部枝桠装饰，烘焙区 1517x470px 贴顶右", "suggested_query": f"{CQ} 枝桠 横版"},
        {"id": "ink_blot", "kind": "ink_blot", "box": (pic_assets["ink_blot"] or {"box": [0.9, 2.55, 0.65, 0.65]})["box"], "spec": "墨点章（程序化椭圆簇+溅点）"},
        {"id": "arc", "kind": "arc", "box": (pic_assets["arc"] or {"box": [0.55, 0.72, 0.5, 0.5]})["box"], "spec": "标题旁笔触小弧（程序化分段弧）"},
    ]
    for c in pic_assets["clouds"]:
        asset_slots.append({"id": c["key"], "kind": "cloud", "box": c["box"], "suggested_query": f"{CQ} 祥云 装饰", "spec": "默认程序化简笔祥云"})
    for b in pic_assets["bands"]:
        asset_slots.append({"id": b["key"], "kind": "band", "box": b["box"], "suggested_query": f"{CQ} 长卷 横带", "spec": "页面横带，宽>60%页宽"})
    if toc_info:
        for j, sb in enumerate(toc_info["strip_boxes"], 1):
            asset_slots.append({"id": f"toc_strip_{j}", "kind": "strip_src", "box": sb,
                                "suggested_query": f"{CQ} 竖条 画心", "spec": "目录竖条源图（脚本裁竖幅+暗化+巨字序号+白标签）"})
    hero_box = skels.get("hero_blocks", {}).get("ill", {}).get("box", [4.22, 1.55, 4.9, 3.6])
    asset_slots.append({"id": "hero_ill", "kind": "hero_ill", "box": hero_box, "suggested_query": f"{CQ} 主题插画 横版"})
    asset_slots.append({"id": "summary_ill", "kind": "summary_ill",
                        "box": (summary_info or {}).get("ill", {}).get("box", [6.9, 1.2, 5.6, 4.2]),
                        "suggested_query": f"{CQ} 氛围插画 横版"})

    page_skels = {"cover": {k: v for k, v in (cover_info or {}).items() if not k.startswith("_")},
                  "toc": toc_info, "summary": {k: v for k, v in (summary_info or {}).items() if not k.startswith("_")}}
    page_skels.update(skels)

    def dehex(dist):
        """主色调占比去 HEX 化：只留 ratio + lum/sat/val 数值（色彩结构，不带色值身份）"""
        out = []
        for d in dist:
            r, g, b = [v/255 for v in hex2rgb(d["hex"])]
            h, s, v = colorsys.rgb_to_hsv(r, g, b)
            out.append({"ratio": d["ratio"], "lum": round(lum(d["hex"]), 3), "sat": round(s, 3), "val": round(v, 3)})
        return out

    skel = {
        "skeleton_id": name, "source_template": os.path.basename(pptx),
        "canvas": {"width_inch": CANVAS_W, "height_inch": CANVAS_H, "safe_margin_inch": 0.8, "min_element_inch": 0.5, "bake_dpi": 200},
        "color_structure": color_structure,
        "palette": {"note": "色角色结构见顶层 color_structure（s/v/色相偏移）；HEX 一律不落盘",
                    "page_color_distribution": {f"page_{i+1:02d}": dehex(p["distribution"][:4]) for i, p in enumerate(profiles)},
                    "page_mean_lum": {f"page_{i+1:02d}": p["mean_lum"] for i, p in enumerate(profiles)}},
        "fonts": {"giant": {"family": "Xingkai SC", "mode": "PIL烘焙",
                            "ghost_echo": {"offset_inch": 0.28, "scale": 1.06, "alpha": 170}},
                  "page_title": {"family": "Songti SC", "size_pt": 36, "bold": True},
                  "body": {"family": "Songti SC", "size_pt": 14},
                  "caption": {"family": "Kaiti SC", "size_pt": 16, "color_role": "primary"}},
        "asset_slots": asset_slots,
        "page_skeletons": page_skels,
        "toc_numerals": "壹贰叁肆",
        "rules": {"decorations_per_page_max": 3, "contrast": {"normal": 4.5, "large": 3.0},
                  "giant_title_layout": "4字=2x2错落；3字=1+2错落",
                  "zero_reuse": "复刻=骨架/血统/年龄：原稿素材与色值零入包，素材按 asset_slots 检索词现下载，色板由内容主色派生"}
    }
    with open(os.path.join(out_dir, "skeleton.json"), "w", encoding="utf-8") as f:
        json.dump(skel, f, ensure_ascii=False, indent=1)

    # ── extracted_scheme v2.1 兼容层（ppt-extract 精华并入，色值全派生占位） ──
    if style_label is None:
        style_label = "中国风" if (pic_assets["ink_blot"] or any("Kaiti" in (e.get("font") or "") or "Xingkai" in (e.get("font") or "")
                                   for els in slides for e in els)) else "未标定"
    compat = build_scheme_compat(name, style_label, page_skels, content_pages, profiles, pic_assets)
    with open(os.path.join(out_dir, "extracted_scheme.json"), "w", encoding="utf-8") as f:
        json.dump(compat, f, ensure_ascii=False, indent=1)

    # ── 复刻脚本（薄封装 → 通用引擎） ──
    with open(os.path.join(out_dir, "build_deck.py"), "w", encoding="utf-8") as f:
        f.write("""#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# 复刻脚本（由 forge_skeleton.py 母版生成）：龙骨/主色调/素材需求全部固话在 skeleton.json
# 用法：python3 build_deck.py --content 文案.json [--primary auto|#HEX] [--fetch 检索词] [--out 输出.pptx]
import os, sys
sys.path.insert(0, %r)
from deck_engine import main
if __name__ == "__main__":
    main(skeleton_dir=os.path.dirname(os.path.abspath(__file__)))
""" % ENGINE_DIR)
    os.chmod(os.path.join(out_dir, "build_deck.py"), 0o755)

    # ── 示例文案（原稿提取；烘焙巨字为占位，需人工填） ──
    ct = (cover_info or {}).get("_cover_texts", {})
    st = (summary_info or {}).get("_texts", {})
    example = {
        "cover": {"giant": "四字巨字", "subtitle": ct.get("subtitle", ""), "caption": ct.get("caption", ""), "seal": ct.get("seal", "印")},
        "toc": {"title": "目录", "labels": ["篇章一", "篇章二", "篇章三", "篇章四"]},
        "pages": example_pages,
        "summary": {"giant": "四字巨字", "body": st.get("body", ""), "caption": st.get("caption", ""), "seal": st.get("seal", "印")},
        "_note": "giant/labels 在原稿中烘焙于图片内，母版无法 OCR，已用占位符；正文类文本均为原稿实测提取"
    }
    with open(os.path.join(out_dir, "examples", "content.example.json"), "w", encoding="utf-8") as f:
        json.dump(example, f, ensure_ascii=False, indent=2)

    # ── README ──
    with open(os.path.join(out_dir, "README.md"), "w", encoding="utf-8") as f:
        f.write(f"""# 龙骨包 {name}（母版 forge_skeleton.py 生成）

源模板：{os.path.basename(pptx)}（{len(slides)} 页）

## 快速出片
```bash
python3 build_deck.py --demo                          # 示例一键出片
python3 build_deck.py --content 文案.json --out 我的.pptx
python3 build_deck.py --content 文案.json --primary "#2B5EA7"   # 换主色调（全 deck 统一迁移）
python3 build_deck.py --content 文案.json --fetch "蓝色的中国风"   # 自动下载素材
```

## 可变项（仅 4 类）
1. 主色调 --primary（auto=按文案关键词推导；全套色板由内容主色派生——skeleton.json 只存色彩结构 color_structure，零模板色值）
2. 插图：--toc-images a,b,c,d / --hero-ill / --summary-ill（槽位规格与检索词见 skeleton.json asset_slots；零原稿素材，一律现下载/供图）
3. 背景图：--paper / --branch（缺省程序化宣纸底/跳过顶枝）
4. 文案：--content JSON（schema 见 examples/content.example.json；内容页 variant 可选 {"、".join(sorted(skels)) or "见 skeleton.json"}）

## 固定项
页面骨架/元素坐标/字号字体/装饰工艺/蒙版位置 —— 全部实测自原稿 XML，写死于 skeleton.json。
""")
    shutil.rmtree(work, ignore_errors=True)
    print(json.dumps({"status": "FORGED", "package": out_dir, "pages": len(slides),
                      "recipes": sorted(skels), "color_structure_keys": sorted(color_structure)}, ensure_ascii=False))

def norm_font(name):
    if not name:
        return "Songti SC"
    if "Xingkai" in name or "Kaiti" in name or "Songti" in name:
        return name
    return "Songti SC"

# ══ 6. extracted_scheme v2.1 兼容层（ppt-extract 精华·机械推导） ════════════
def color_name_of(primary_hex):
    """primary HEX → 中文颜色主基调（hue 机械映射）"""
    r, g, b = [v/255 for v in hex2rgb(primary_hex)]
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    if s < 0.12:
        return "白" if v > 0.8 else ("灰" if v > 0.35 else "黑")
    deg = h * 360
    for name, lo, hi in [("红", 345, 360), ("红", 0, 15), ("橙", 15, 45), ("黄", 45, 70),
                          ("绿", 70, 165), ("青", 165, 195), ("蓝", 195, 255), ("紫", 255, 290),
                          ("粉", 290, 345)]:
        if lo <= deg < hi:
            return name
    return "红"

def build_scheme_compat(name, style_label, page_skels, content_pages, profiles, pic_assets):
    """extracted_scheme.json schema v2.1 兼容视图：结构/版式/字体家族为母版实测；
    全部色值角色写 derive:<role> 占位（由内容主色派生，零模板色值）"""
    seq_map = {"cover": "cover", "toc": "toc", "summary": "summary",
               "seal_blocks": "content", "plum_blocks": "content", "scroll_cards": "content",
               "cards": "content", "hero_blocks": "content"}
    seq = ["cover", "toc"] + [seq_map.get(p["variant"], "content") for p in content_pages] + ["summary"]
    titles = [ps["title"]["size"] for k, ps in page_skels.items() if isinstance(ps, dict) and "title" in ps]
    body_size = 14
    for k, ps in page_skels.items():
        if isinstance(ps, dict) and ps.get("group", {}).get("body_size"):
            body_size = ps["group"]["body_size"]; break
    t_size = titles[0] if titles else 36
    light = all(p["mean_lum"] > 0.5 for i, p in enumerate(profiles))
    CQ = "<内容主色>色的<风格>"
    role_usage = {"primary": "印章/落款/强调点缀", "secondary": "序号章底", "accent": "小面积强调",
                  "text_dark": "巨字/页面标题/正文", "text_light": "章内白字/深色画心标签", "background": "页面底色"}
    compat = {
        "schema_version": "v2.1",
        "style_label": style_label,
        "color_name": "由内容派生",
        "template_id": f"{style_label}_<内容主色>_{name}",
        "fixed_theme_color": {k: f"derive:{k}" for k in
                              ("primary", "secondary", "accent", "text_dark", "text_light", "background")},
        "color_usage_map": role_usage,
        "color_usage_rules": {k: {"allow": [v], "ban": []} for k, v in role_usage.items()},
        "layout_metrics": {
            "margin_ratio": [0.055, 0.065],
            "fill_max_ratio": [0.62, 0.72],
            "decoration_quota": [2, 3],
            "title_body_scale": [round(t_size/body_size*0.95, 2), round(t_size/body_size*1.05, 2)],
            "line_height": [1.3, 1.5],
            "paragraph_spacing": [0.18, 0.3]},
        "font_spec": {
            "title": {"family": page_skels.get("cover", {}).get("subtitle", {}).get("font", "Songti SC"),
                      "size_pt": t_size, "color": "derive:text_dark",
                      "screen_ratio": {"width": [round(page_skels.get("cover", {}).get("giant_region", [0,0,5,5])[2]/CANVAS_W*0.9, 2),
                                                round(page_skels.get("cover", {}).get("giant_region", [0,0,5,5])[2]/CANVAS_W*1.1, 2)]},
                      "shadow_layer": {"applicable": True, "mode": "ghost_echo_bake",
                                       "note": "母版实测：巨字烘焙于画底，幽灵回声=偏移0.28in/放大1.06/纸灰alpha170"}},
            "seal": {"family": "Kaiti SC", "color": "derive:text_light"},
            "body": {"family": "Songti SC", "size_pt": [body_size, body_size+4], "color": "derive:text_dark"},
            "caption": {"family": "Kaiti SC", "color": "derive:primary"}},
        "layer_stack": {
            "order": ["paper_base", "illustration", "decoration", "text"],
            "_boundary_note": "浅底" if light else "含深色页",
            "source": "母版由 XML z-order + 像素明暗实测推导"},
        "asset_policy": {
            "whitelist": [a for a, present in [("横带", bool(pic_assets["bands"])), ("祥云角饰", bool(pic_assets["clouds"])),
                                                ("墨点章", bool(pic_assets["ink_blot"])), ("序号章", True), ("方印章", True),
                                                ("目录竖条画心", "toc" in page_skels), ("书法巨字烘焙", True)] if present],
            "blacklist": [],
            "card_spec": page_skels.get("scroll_cards", {}).get("card") or None,
            "max_per_page": 3},
        "decoration_spec": {
            "render_priority": "illustration_first",
            "elements": ([{"name": f"祥云_{i+1}", "asset_path": None, "placement": ["角落"]}
                          for i, c in enumerate(pic_assets["clouds"])]
                         + ([{"name": "笔触弧", "asset_path": None, "placement": ["标题左侧"]}] if pic_assets["arc"] else []))},
        "imagery_map": {
            "applicable": True,
            "mapping": [{"page_type": "toc", "suggested_query": CQ},
                        {"page_type": "content", "suggested_query": f"{CQ} 横带"},
                        {"page_type": "summary", "suggested_query": f"{CQ} 插画"}]},
        "business_page_sequence": seq,
        "_compat_note": "本文件由 forge_skeleton 母版机械推导生成（ppt-extract schema v2.1 精华字段兼容层）；色值全为 derive 占位——由内容主色派生，零模板色值；素材 asset_path 置 None——按槽位检索词现下载，零原稿素材"
    }
    return compat

def main():
    ap = argparse.ArgumentParser(description="PPT 复刻母版脚本")
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--name", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--style", default=None, help="风格标签（如 中国风/科技未来）；缺省机械推断")
    args = ap.parse_args()
    forge(args.pptx, args.name, args.out, style_label=args.style)

if __name__ == "__main__":
    main()

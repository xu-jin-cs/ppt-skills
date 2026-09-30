#!/usr/bin/env python3
"""render_ppt.py — PPT 自动化工作流终端渲染器（全链路唯一 Schema 约束源）
用法: python3 render_ppt.py <07_render_input.json>
产物: <输入同目录>/output.pptx + output.quality_report.json，stdout 打印执行结果 JSON
"""
import json
import math
import os
import sys

# ══ 全局固定常量（下游所有 Agent 唯一数值标准，禁止推测） ══════════════════
SLIDE_WIDTH_INCH = 13.333
SLIDE_HEIGHT_INCH = 7.5
SAFE_MARGIN_INCH = 0.8
MIN_ELEMENT_SIZE_INCH = 0.5
MIN_FONT_SIZE_PT = 14
MAX_TITLE_SIZE_PT = 54
MAX_SUBTITLE_SIZE_PT = 28
MAX_BODY_SIZE_PT = 24
DEFAULT_FONT = "Microsoft YaHei"
FALLBACK_FONTS = ["Microsoft YaHei", "PingFang SC", "Arial"]
BG_IMAGE_OVERLAY_COLOR = "#000000"
BG_IMAGE_OVERLAY_ALPHA = 0.35
MAX_DECORATIONS_PER_PAGE = 3
CONTRAST_MIN_NORMAL = 4.5
CONTRAST_MIN_LARGE = 3.0
FORBIDDEN_PREVIEW_BG = "ppt_previews"
VALID_PAGE_TYPES = ["cover", "toc", "chapter", "content", "summary"]
VALID_BG_TYPES = ["solid", "gradient", "image"]
VALID_ELEMENT_TYPES = ["text", "image", "shape", "line"]
VALID_ALIGNMENTS = ["left", "center", "right"]
VALID_FIT_MODES = ["cover", "contain", "stretch"]
VALID_ROLES = ["title", "subtitle", "body", "keyword", "decoration"]
ROLE_SIZE_RANGE = {"title": (36, MAX_TITLE_SIZE_PT), "subtitle": (20, MAX_SUBTITLE_SIZE_PT), "body": (MIN_FONT_SIZE_PT, MAX_BODY_SIZE_PT), "keyword": (MIN_FONT_SIZE_PT, MAX_BODY_SIZE_PT)}
TOP_LEVEL_KEYS = ["meta", "theme", "asset_base_path", "slides"]
SLIDE_KEYS = ["page_no", "page_type", "background", "elements"]
THEME_COLOR_KEYS = ["primary", "secondary", "accent", "text_dark", "text_light", "background"]
THEME_FONT_KEYS = ["title", "subtitle", "body"]

# ══ 纯函数工具 ════════════════════════════════════════════════════════════
def hex_to_rgb(h):
    h = str(h).lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4)) if len(h) == 6 else (0, 0, 0)

def _lum_channel(c):
    c = c / 255.0
    return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

def luminance(hex_color):
    r, g, b = hex_to_rgb(hex_color)
    return 0.2126 * _lum_channel(r) + 0.7152 * _lum_channel(g) + 0.0722 * _lum_channel(b)

def contrast_ratio(c1, c2):
    l1, l2 = sorted([luminance(c1), luminance(c2)], reverse=True)
    return (l1 + 0.05) / (l2 + 0.05)

def is_large_text(size_pt, bold):
    return size_pt >= 18 or bold

def check_contrast(fg, bg, size_pt, bold):
    return contrast_ratio(fg, bg) >= (CONTRAST_MIN_LARGE if is_large_text(size_pt, bold) else CONTRAST_MIN_NORMAL)

def clamp_font_size(role, size_pt):
    lo, hi = ROLE_SIZE_RANGE.get(role, (MIN_FONT_SIZE_PT, MAX_BODY_SIZE_PT))
    return max(lo, min(hi, size_pt)), (size_pt != max(lo, min(hi, size_pt)))

def estimate_text_overflow(text, box_w_in, box_h_in, size_pt):
    """CJK 宽≈字号，Latin≈0.55 字号；行高 1.25 倍。返回是否溢出。"""
    width_pt, height_pt = box_w_in * 72.0, box_h_in * 72.0
    char_w = lambda ch: size_pt * (1.0 if ord(ch) > 0x2E7F else 0.55)
    lines, cur = 1, 0.0
    for ch in str(text):
        if ch == "\n":
            lines, cur = lines + 1, 0.0
            continue
        cur += char_w(ch)
        if cur > width_pt:
            lines, cur = lines + 1, char_w(ch)
    return lines * size_pt * 1.25 > height_pt

# ══ Schema 校验（复刻给主线 07 的同一套规则） ══════════════════════════════
def validate_input(data):
    errors = []
    for k in TOP_LEVEL_KEYS:
        if k not in data:
            errors.append(f"missing top-level key: {k}")
    if errors:
        return errors
    meta, theme = data["meta"], data["theme"]
    if abs(meta.get("slide_width_inch", 0) - SLIDE_WIDTH_INCH) > 0.01 or abs(meta.get("slide_height_inch", 0) - SLIDE_HEIGHT_INCH) > 0.01:
        errors.append("meta canvas size mismatch")
    colors = theme.get("colors", {})
    for k in THEME_COLOR_KEYS:
        if k not in colors:
            errors.append(f"theme.colors missing: {k}")
    if not isinstance(colors.get("secondary"), list) or len(colors.get("secondary", [])) < 2:
        errors.append("theme.colors.secondary must be list with >=2 colors")
    for k in THEME_FONT_KEYS:
        if k not in theme.get("fonts", {}):
            errors.append(f"theme.fonts missing: {k}")
    if not data.get("asset_base_path"):
        errors.append("asset_base_path empty")
    slides = data.get("slides")
    if not isinstance(slides, list) or not slides:
        return errors + ["slides must be non-empty list"]
    for i, s in enumerate(slides):
        for k in SLIDE_KEYS:
            if k not in s:
                errors.append(f"slides[{i}] missing: {k}")
        pt = s.get("page_type")
        if pt not in VALID_PAGE_TYPES:
            errors.append(f"slides[{i}] illegal page_type: {pt}")
        bg = s.get("background", {})
        if bg.get("type") not in VALID_BG_TYPES:
            errors.append(f"slides[{i}] illegal background.type: {bg.get('type')}")
        if bg.get("type") == "image" and FORBIDDEN_PREVIEW_BG in str(bg.get("path", "")):
            errors.append(f"slides[{i}] forbidden ppt_previews as background")
        deco = 0
        for j, el in enumerate(s.get("elements", [])):
            et = el.get("type")
            if et not in VALID_ELEMENT_TYPES:
                errors.append(f"slides[{i}].elements[{j}] illegal type: {et}")
                continue
            box = el.get("box", {})
            w, h = box.get("width_inch", 0), box.get("height_inch", 0)
            x, y = box.get("x_inch", 0), box.get("y_inch", 0)
            if et != "line" and (w < MIN_ELEMENT_SIZE_INCH or h < MIN_ELEMENT_SIZE_INCH):
                errors.append(f"slides[{i}].elements[{j}] box too small: {w}x{h}")
            if x + w > SLIDE_WIDTH_INCH + 0.01 or y + h > SLIDE_HEIGHT_INCH + 0.01 or x < -0.01 or y < -0.01:
                if not (bg.get("type") == "image" and et == "image" and x == 0 and y == 0):
                    errors.append(f"slides[{i}].elements[{j}] out of canvas: x={x} y={y} w={w} h={h}")
            if et == "text":
                st = el.get("style", {})
                if st.get("align", "left") not in VALID_ALIGNMENTS:
                    errors.append(f"slides[{i}].elements[{j}] illegal align")
            if et == "image":
                if el.get("fit", "cover") not in VALID_FIT_MODES:
                    errors.append(f"slides[{i}].elements[{j}] illegal fit")
                if FORBIDDEN_PREVIEW_BG in str(el.get("path", "")):
                    errors.append(f"slides[{i}].elements[{j}] forbidden ppt_previews image")
            if el.get("role") == "decoration" or et == "line" or (et == "shape" and el.get("role") != "card"):
                deco += 1
        if deco > MAX_DECORATIONS_PER_PAGE:
            errors.append(f"slides[{i}] too many decorations: {deco} > {MAX_DECORATIONS_PER_PAGE}")
    return errors

# ══ 渲染器 ════════════════════════════════════════════════════════════════
def _set_alpha(fill_owner, alpha):
    """给 solidFill 注入透明度（python-pptx 无原生 API，走 XML）。"""
    from pptx.oxml.ns import qn
    srgb = fill_owner.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
    a = srgb.makeelement(qn("a:alpha"), {"val": str(int(alpha * 100000))})
    srgb.append(a)

def render_background(slide, bg, base_path, report, page_no):
    from pptx.util import Emu
    from pptx.dml.color import RGBColor
    full_w, full_h = Emu(int(SLIDE_WIDTH_INCH * 914400)), Emu(int(SLIDE_HEIGHT_INCH * 914400))
    shp = slide.shapes.add_shape(1, 0, 0, full_w, full_h)  # 1=rectangle
    shp.line.fill.background()
    shp.shadow.inherit = False
    btype = bg.get("type")
    if btype == "solid":
        shp.fill.solid()
        shp.fill.fore_color.rgb = RGBColor(*hex_to_rgb(bg.get("color", "#FFFFFF")))
    elif btype == "gradient":
        try:
            shp.fill.gradient()
            stops = shp.fill.gradient_stops
            c = bg.get("colors", ["#FFFFFF", "#EEEEEE"])
            stops[0].color.rgb, stops[0].position = RGBColor(*hex_to_rgb(c[0])), 0.0
            stops[1].color.rgb, stops[1].position = RGBColor(*hex_to_rgb(c[-1])), 1.0
            shp.fill.gradient_angle = bg.get("angle", 135)
        except Exception:
            shp.fill.solid()
            shp.fill.fore_color.rgb = RGBColor(*hex_to_rgb(bg.get("colors", ["#FFFFFF"])[0]))
    elif btype == "image":
        path = os.path.join(base_path, bg.get("path", ""))
        if os.path.exists(path):
            slide.shapes._spTree.remove(shp._element)
            pic = slide.shapes.add_picture(path, 0, 0, full_w, full_h)
            _apply_cover_crop(pic, path, SLIDE_WIDTH_INCH, SLIDE_HEIGHT_INCH)
            ov = slide.shapes.add_shape(1, 0, 0, full_w, full_h)
            ov.line.fill.background()
            ov.shadow.inherit = False
            ov.fill.solid()
            ov.fill.fore_color.rgb = RGBColor(*hex_to_rgb(bg.get("overlay_color", BG_IMAGE_OVERLAY_COLOR)))
            _set_alpha(ov, bg.get("overlay_alpha", BG_IMAGE_OVERLAY_ALPHA))
            return
        report["missing_images"].append({"page_no": page_no, "path": bg.get("path")})
        shp.fill.solid()
        shp.fill.fore_color.rgb = RGBColor(*hex_to_rgb(bg.get("fallback_color", "#333333")))

def _apply_cover_crop(pic, path, box_w_in, box_h_in):
    try:
        from PIL import Image
        with Image.open(path) as im:
            iw, ih = im.size
        img_asp, box_asp = iw / ih, box_w_in / box_h_in
        if img_asp > box_asp:
            pic.crop_left = pic.crop_right = (1 - box_asp / img_asp) / 2
        elif img_asp < box_asp:
            pic.crop_top = pic.crop_bottom = (1 - img_asp / box_asp) / 2
    except Exception:
        pass

def render_text(slide, el, theme, report, page_no):
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    box = el["box"]
    st = el.get("style", {})
    role = el.get("role", "body")
    size, clamped = clamp_font_size(role, st.get("size_pt", 18))
    if clamped:
        report["font_size_clamped"].append({"page_no": page_no, "text": str(el.get("text", ""))[:20], "from": st.get("size_pt"), "to": size})
    if estimate_text_overflow(el.get("text", ""), box["width_inch"], box["height_inch"], size):
        report["overflow"].append({"page_no": page_no, "text": str(el.get("text", ""))[:30]})
    bg_hex = el.get("on_color", "#FFFFFF")
    fg = st.get("color", "#000000")
    if not check_contrast(fg, bg_hex, size, st.get("bold", False)):
        report["contrast_failures"].append({"page_no": page_no, "text": str(el.get("text", ""))[:20], "fg": fg, "bg": bg_hex, "ratio": round(contrast_ratio(fg, bg_hex), 2)})
    tb = slide.shapes.add_textbox(Inches(box["x_inch"]), Inches(box["y_inch"]), Inches(box["width_inch"]), Inches(box["height_inch"]))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE if role in ("title", "subtitle") else MSO_ANCHOR.TOP
    lines = str(el.get("text", "")).split("\n")
    for idx, line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}[st.get("align", "left")]
        run = p.add_run()
        run.text = line
        f = run.font
        f.name = st.get("font", theme.get("fonts", {}).get(role if role in THEME_FONT_KEYS else "body", DEFAULT_FONT))
        f.size = Pt(size)
        f.bold = st.get("bold", role == "title")
        f.color.rgb = RGBColor(*hex_to_rgb(fg))
        from pptx.oxml.ns import qn
        rPr = run._r.get_or_add_rPr()
        ea = rPr.makeelement(qn("a:ea"), {"typeface": f.name})
        rPr.append(ea)

def render_image(slide, el, base_path, report, page_no):
    from pptx.util import Inches
    box = el["box"]
    path = os.path.join(base_path, el.get("path", ""))
    if not os.path.exists(path):
        report["missing_images"].append({"page_no": page_no, "path": el.get("path")})
        fb = slide.shapes.add_shape(1, Inches(box["x_inch"]), Inches(box["y_inch"]), Inches(box["width_inch"]), Inches(box["height_inch"]))
        from pptx.dml.color import RGBColor
        fb.line.fill.background()
        fb.shadow.inherit = False
        fb.fill.solid()
        fb.fill.fore_color.rgb = RGBColor(*hex_to_rgb(el.get("fallback_color", "#CCCCCC")))
        return
    pic = slide.shapes.add_picture(path, Inches(box["x_inch"]), Inches(box["y_inch"]), Inches(box["width_inch"]), Inches(box["height_inch"]))
    fit = el.get("fit", "cover")
    if fit == "cover":
        _apply_cover_crop(pic, path, box["width_inch"], box["height_inch"])

def render_shape(slide, el, report, page_no):
    from pptx.util import Inches
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    box = el["box"]
    kind = {"rect": MSO_SHAPE.RECTANGLE, "round_rect": MSO_SHAPE.ROUNDED_RECTANGLE, "circle": MSO_SHAPE.OVAL}.get(el.get("shape", "round_rect"), MSO_SHAPE.ROUNDED_RECTANGLE)
    shp = slide.shapes.add_shape(kind, Inches(box["x_inch"]), Inches(box["y_inch"]), Inches(box["width_inch"]), Inches(box["height_inch"]))
    shp.shadow.inherit = False
    if el.get("shape") == "round_rect":
        try:
            shp.adjustments[0] = min(0.5, el.get("radius_pt", 10) / 100.0)
        except Exception:
            pass
    if el.get("no_fill"):
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = RGBColor(*hex_to_rgb(el.get("color", "#FFFFFF")))
        if "alpha" in el:
            _set_alpha(shp, el["alpha"])
    if el.get("line_color"):
        shp.line.color.rgb = RGBColor(*hex_to_rgb(el["line_color"]))
        shp.line.width = Inches(0.02)
    else:
        shp.line.fill.background()

def render_line(slide, el):
    from pptx.util import Inches
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_CONNECTOR
    box = el["box"]
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(box["x_inch"]), Inches(box["y_inch"]), Inches(box["x_inch"] + box["width_inch"]), Inches(box["y_inch"] + box["height_inch"]))
    ln.line.color.rgb = RGBColor(*hex_to_rgb(el.get("color", "#000000")))
    ln.line.width = Inches(el.get("weight_inch", 0.02))

# ══ 主流程 ════════════════════════════════════════════════════════════════
def render(data, out_pptx):
    from pptx import Presentation
    from pptx.util import Inches
    report = {"overflow": [], "missing_images": [], "contrast_failures": [], "too_many_decorations": [], "font_size_clamped": []}
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(SLIDE_WIDTH_INCH), Inches(SLIDE_HEIGHT_INCH)
    blank = prs.slide_layouts[6]
    theme, base = data["theme"], data["asset_base_path"]
    for s in data["slides"]:
        page_no = s.get("page_no", 0)
        slide = prs.slides.add_slide(blank)
        render_background(slide, s["background"], base, report, page_no)
        deco = sum(1 for e in s["elements"] if e.get("role") == "decoration" or e.get("type") == "line" or (e.get("type") == "shape" and e.get("role") != "card"))
        if deco > MAX_DECORATIONS_PER_PAGE:
            report["too_many_decorations"].append({"page_no": page_no, "count": deco})
        for el in s["elements"]:
            {"text": render_text, "image": render_image, "shape": render_shape, "line": render_line}[el["type"]](
                slide, el, *( (theme, report, page_no) if el["type"] == "text" else (base, report, page_no) if el["type"] == "image" else (report, page_no) if el["type"] == "shape" else () ))
    prs.save(out_pptx)
    return report

def main():
    if len(sys.argv) < 2:
        print(json.dumps({"status": "error", "message": "usage: python3 render_ppt.py <input.json>"}, ensure_ascii=False))
        sys.exit(1)
    in_path = sys.argv[1]
    data = json.loads(open(in_path, encoding="utf-8").read())
    run_dir = os.path.dirname(os.path.abspath(in_path))
    errors = validate_input(data)
    if errors:
        result = {"status": "schema_error", "errors": errors}
        open(os.path.join(run_dir, "output.quality_report.json"), "w", encoding="utf-8").write(json.dumps(result, ensure_ascii=False, indent=2))
        print(json.dumps(result, ensure_ascii=False))
        sys.exit(2)
    out_pptx = os.path.join(run_dir, "output.pptx")
    try:
        rep = render(data, out_pptx)
    except Exception as e:
        print(json.dumps({"status": "render_error", "message": str(e)}, ensure_ascii=False))
        sys.exit(3)
    blocking = len(rep["overflow"]) + len(rep["missing_images"]) + len(rep["contrast_failures"]) + len(rep["too_many_decorations"])
    quality = {"status": "PASS" if blocking == 0 else "FAIL", "overflow_count": len(rep["overflow"]), "missing_images_count": len(rep["missing_images"]), "contrast_failure_count": len(rep["contrast_failures"]), "too_many_decorations_count": len(rep["too_many_decorations"]), "font_size_clamped_count": len(rep["font_size_clamped"]), "blocking_total": blocking, "details": rep, "output_pptx": out_pptx}
    open(os.path.join(run_dir, "output.quality_report.json"), "w", encoding="utf-8").write(json.dumps(quality, ensure_ascii=False, indent=2))
    print(json.dumps({"status": quality["status"], "pptx": out_pptx, "blocking_total": blocking, "font_size_clamped_count": quality["font_size_clamped_count"]}, ensure_ascii=False))
    sys.exit(0 if blocking == 0 else 4)

if __name__ == "__main__":
    main()

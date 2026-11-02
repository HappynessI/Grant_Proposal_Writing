import argparse
from io import BytesIO
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.shapes.group import GroupShape
from pptx.shapes.picture import Picture

from deck_ops import parse_pages

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
FONT_CANDIDATES = [
    "/System/Library/Fonts/STHeiti Medium.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "C:/Windows/Fonts/msyh.ttc",
]
FONT_PATH = next(p for p in FONT_CANDIDATES if Path(p).exists())
PX_PER_IN = 120
EMU = 914400


def theme_colors(prs):
    # 读取主题配色，用于还原 schemeClr 填充
    master = prs.slide_masters[0]
    for rel in master.part.rels.values():
        if rel.reltype.endswith("/theme"):
            from lxml import etree

            root = etree.fromstring(rel.target_part.blob)
            scheme = root.find(f".//{{{A_NS}}}clrScheme")
            colors = {}
            for node in scheme:
                child = node[0]
                value = child.get("val") if child.tag.endswith("srgbClr") else child.get("lastClr")
                colors[node.tag.split("}")[1]] = value
            colors.update(tx1=colors["dk1"], bg1=colors["lt1"], tx2=colors["dk2"], bg2=colors["lt2"])
            return colors
    raise ValueError("母版没有主题")


def to_rgb(clr, theme):
    if clr is None:
        return None
    value = clr.get("val") if clr.tag.endswith("srgbClr") else theme.get(clr.get("val"))
    if value is None or len(value) != 6:
        return None
    return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))


def solid(node, theme):
    if node is None:
        return None
    fill = node.find(f"{{{A_NS}}}solidFill")
    return to_rgb(fill[0], theme) if fill is not None and len(fill) else None


def draw_shape(canvas, draw, s, tf, theme, prs):
    x, y, w, h = tf(s.left, s.top, s.width, s.height)
    box = [x, y, x + max(w, 1), y + max(h, 1)]
    if isinstance(s, Picture):
        if s._pic.blipFill.blip.rEmbed and s.image.ext not in ("wmf", "emf", "svg"):
            img = Image.open(BytesIO(s.image.blob)).convert("RGB")
            iw, ih = img.size
            img = img.crop((int(iw * s.crop_left), int(ih * s.crop_top), int(iw * (1 - s.crop_right)), int(ih * (1 - s.crop_bottom))))
            canvas.paste(img.resize((max(int(w), 1), max(int(h), 1))), (int(x), int(y)))
        else:
            draw.rectangle(box, outline=(150, 150, 150))
        return
    sp_pr = s._element.find("{http://schemas.openxmlformats.org/presentationml/2006/main}spPr")
    fill, line = None, None
    if sp_pr is not None:
        fill = solid(sp_pr, theme)
        ln = sp_pr.find(f"{{{A_NS}}}ln")
        line = solid(ln, theme) if ln is not None else None
    if fill or line:
        draw.rectangle(box, fill=fill, outline=line)
    if not s.has_text_frame or not s.text_frame.text.strip():
        return
    cy = y + 3
    color = (35, 35, 35)
    for p in s.text_frame.paragraphs:
        text = "".join(r.text for r in p.runs)
        size = next((r.font.size.pt for r in p.runs if r.font.size), 18)
        run_color = next((r.font._element.find(f"{{{A_NS}}}solidFill") for r in p.runs if r.font._element is not None), None)
        if run_color is not None and len(run_color):
            color = to_rgb(run_color[0], theme) or color
        px = max(int(size / 72 * PX_PER_IN), 8)
        font = ImageFont.truetype(FONT_PATH, px)
        line_text = ""
        for ch in text:
            if draw.textlength(line_text + ch, font=font) > w - 12:
                draw.text((x + 6, cy), line_text, fill=color, font=font)
                cy += px * 1.2
                line_text = ch
            else:
                line_text += ch
        draw.text((x + 6, cy), line_text, fill=color, font=font)
        cy += px * 1.2


def draw_tree(canvas, draw, shapes, tf, theme, prs):
    for s in shapes:
        if isinstance(s, GroupShape):
            xfrm = s._element.grpSpPr.find(f"{{{A_NS}}}xfrm")
            off, ext = xfrm.find(f"{{{A_NS}}}off"), xfrm.find(f"{{{A_NS}}}ext")
            ch_off, ch_ext = xfrm.find(f"{{{A_NS}}}chOff"), xfrm.find(f"{{{A_NS}}}chExt")
            sx = int(ext.get("cx")) / max(int(ch_ext.get("cx")), 1)
            sy = int(ext.get("cy")) / max(int(ch_ext.get("cy")), 1)

            def child_tf(l, t, w, h, sx=sx, sy=sy, off=off, ch_off=ch_off, parent=tf):
                return parent(int(off.get("x")) + (l - int(ch_off.get("x"))) * sx, int(off.get("y")) + (t - int(ch_off.get("y"))) * sy, w * sx, h * sy)

            draw_tree(canvas, draw, s.shapes, child_tf, theme, prs)
        elif s.left is not None:
            draw_shape(canvas, draw, s, tf, theme, prs)


def render(prs, slide, theme, out):
    size = (int(prs.slide_width / EMU * PX_PER_IN), int(prs.slide_height / EMU * PX_PER_IN))
    canvas = Image.new("RGB", size, "white")
    draw = ImageDraw.Draw(canvas)

    def tf(l, t, w, h):
        k = PX_PER_IN / EMU
        return l * k, t * k, w * k, h * k

    draw_tree(canvas, draw, slide.slide_layout.slide_master.shapes, tf, theme, prs)
    draw_tree(canvas, draw, slide.shapes, tf, theme, prs)
    canvas.save(out)


def main():
    parser = argparse.ArgumentParser(description="绘制 PPT 版面草图，用于检查位置和遮挡；字体和颜色只是近似效果")
    parser.add_argument("deck")
    parser.add_argument("pages", help="页码，例如 3,5-8")
    parser.add_argument("out", type=Path)
    args = parser.parse_args()
    args.out.mkdir(parents=True, exist_ok=True)
    prs = Presentation(args.deck)
    theme = theme_colors(prs)
    pages = parse_pages(args.pages)
    for i, slide in enumerate(prs.slides, 1):
        if i in pages:
            path = args.out / f"page_{i:03d}.png"
            render(prs, slide, theme, path)
            print(path)


if __name__ == "__main__":
    main()

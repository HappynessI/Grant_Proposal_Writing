import argparse
from io import BytesIO
from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.shapes.group import GroupShape
from pptx.shapes.picture import Picture

from deck_ops import parse_pages, walk

EMU = 914400


def picture_ok(s):
    # 跳过链接图片和 Pillow 打不开的矢量格式
    return bool(s._pic.blipFill.blip.rEmbed) and s.image.ext not in ("wmf", "emf", "svg")


def outline(prs, pages):
    for i, slide in enumerate(prs.slides, 1):
        if pages and i not in pages:
            continue
        print(f"===== 第 {i} 页 =====")
        for s in walk(slide.shapes):
            if isinstance(s, Picture):
                print("  [图片]")
            if s.has_text_frame and s.text_frame.text.strip():
                print("  " + s.text_frame.text.strip().replace("\n", " / "))
            if getattr(s, "has_table", False) and s.has_table:
                for r in s.table.rows:
                    print("  " + " | ".join(c.text.strip() for c in r.cells))


def _dump_tree(shapes, depth):
    for s in shapes:
        extra = ""
        if isinstance(s, Picture) and picture_ok(s):
            extra = f" img={s.image.size} crop=({s.crop_left:.2f},{s.crop_top:.2f},{s.crop_right:.2f},{s.crop_bottom:.2f})"
        if s.has_text_frame and s.text_frame.text.strip():
            sizes = sorted({r.font.size.pt for p in s.text_frame.paragraphs for r in p.runs if r.font.size})
            extra = f" font={sizes} text={s.text_frame.text.strip()[:70]!r}"
        print(f"{'  ' * depth}[{s.shape_id}] {type(s).__name__} '{s.name}' x={s.left / EMU:.2f} y={s.top / EMU:.2f} w={s.width / EMU:.2f} h={s.height / EMU:.2f}{extra}")
        if isinstance(s, GroupShape):
            _dump_tree(s.shapes, depth + 1)


def shapes(prs, pages):
    for i, slide in enumerate(prs.slides, 1):
        if i in pages:
            print(f"===== 第 {i} 页 layout={slide.slide_layout.name}")
            _dump_tree(slide.shapes, 0)


def images(prs, pages, out):
    # 导出原始分辨率图片，并为每页生成带编号和尺寸的预览图，方便选图
    out.mkdir(parents=True, exist_ok=True)
    for i, slide in enumerate(prs.slides, 1):
        if i not in pages:
            continue
        thumbs = []
        for n, s in enumerate(x for x in walk(slide.shapes) if isinstance(x, Picture)):
            if not picture_ok(s):
                continue
            img = Image.open(BytesIO(s.image.blob)).convert("RGB")
            if img.width < 200 or s.width >= prs.slide_width * 0.95:
                continue
            name = f"s{i}_{n}"
            img.save(out / f"{name}.png")
            t = img.copy()
            t.thumbnail((380, 300))
            thumbs.append((name, img.size, t))
        if not thumbs:
            print(f"第 {i} 页没有可导出的图片；若页面有视频或动图，请解压 ppt/media 后截取画面")
            continue
        rows = (len(thumbs) + 3) // 4
        sheet = Image.new("RGB", (1600, rows * 330), "white")
        draw = ImageDraw.Draw(sheet)
        for k, (name, size, t) in enumerate(thumbs):
            x, y = (k % 4) * 400, (k // 4) * 330
            sheet.paste(t, (x, y + 20))
            draw.text((x, y + 2), f"{name} {size[0]}x{size[1]}", fill="red")
        sheet.save(out / f"sheet_{i}.png")
        print(out / f"sheet_{i}.png")


def main():
    parser = argparse.ArgumentParser(description="读取 PPT 的文字、形状和图片")
    parser.add_argument("deck")
    parser.add_argument("mode", choices=["outline", "shapes", "images"])
    parser.add_argument("--pages", help="页码，例如 3,5-8；outline 省略时读取全部页")
    parser.add_argument("--out", type=Path, help="images 模式的输出目录")
    args = parser.parse_args()
    prs = Presentation(args.deck)
    pages = parse_pages(args.pages) if args.pages else set()
    if args.mode == "outline":
        outline(prs, pages)
    elif args.mode == "shapes":
        shapes(prs, pages)
    else:
        images(prs, pages, args.out)


if __name__ == "__main__":
    main()

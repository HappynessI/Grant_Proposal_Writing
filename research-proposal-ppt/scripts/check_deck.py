import argparse
import unicodedata

from pptx import Presentation
from pptx.shapes.group import GroupShape
from pptx.shapes.picture import Picture
from pptx.util import Pt

from deck_ops import REVIEW_MARK, parse_pages, walk

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def char_width(ch):
    # 以汉字宽度为 1：全角字符算 1，半角字母数字算 0.55
    return 1.0 if unicodedata.east_asian_width(ch) in ("W", "F") else 0.55


def text_width(text):
    return sum(char_width(c) for c in text)


def count_text(text):
    print(f"字符数 {len(text)}，折合汉字宽度 {text_width(text):.1f}：{text}")


def run_size(para, default):
    sizes = [r.font.size.pt for r in para.runs if r.font.size]
    return max(sizes) if sizes else default


def insets(s):
    body = s._element.find(f".//{{{A_NS}}}bodyPr")
    get = lambda k, d: int(body.get(k, d)) if body is not None else d
    return get("lIns", 91440) + get("rIns", 91440), get("tIns", 45720) + get("bIns", 45720)


def estimate(s, scale_x=1.0, scale_y=1.0):
    # 逐段按框宽估算行数，返回 (总高度, 各段行宽填充率)
    hor, ver = insets(s)
    usable = max(s.width * scale_x - hor, 1)
    height, fills = ver, []
    for p in s.text_frame.paragraphs:
        size = run_size(p, 18.0)
        width_units = text_width("".join(r.text for r in p.runs))
        per_line = usable / Pt(size)
        lines = max(1, int(width_units // per_line) + (width_units % per_line > 0))
        height += lines * Pt(size) * 1.2
        if width_units:
            fills.append((lines, (width_units % per_line or per_line) / per_line))
    return height, fills


def group_scales(slide):
    # 组合内形状使用组合自己的坐标系，估算文字时要换算回页面尺寸
    scales = {}

    def visit(shapes, sx, sy):
        for s in shapes:
            if isinstance(s, GroupShape):
                xfrm = s._element.grpSpPr.find(f"{{{A_NS}}}xfrm")
                ext = xfrm.find(f"{{{A_NS}}}ext")
                ch = xfrm.find(f"{{{A_NS}}}chExt")
                nx = sx * int(ext.get("cx")) / max(int(ch.get("cx")), 1)
                ny = sy * int(ext.get("cy")) / max(int(ch.get("cy")), 1)
                visit(s.shapes, nx, ny)
            else:
                scales[s.shape_id] = (sx, sy)

    visit(slide.shapes, 1.0, 1.0)
    return scales


def check_slide(prs, slide, no, ban):
    issues = []
    scales = group_scales(slide)
    top_level = {s.shape_id for s in slide.shapes}
    for s in walk(slide.shapes):
        if isinstance(s, GroupShape):
            continue
        sx, sy = scales.get(s.shape_id, (1.0, 1.0))
        if s.shape_id in top_level:
            if s.left < -Pt(2) or s.top < -Pt(2) or s.left + s.width > prs.slide_width + Pt(2) or s.top + s.height > prs.slide_height + Pt(2):
                issues.append(f"[{s.shape_id}] {s.name} 超出页面")
        if isinstance(s, Picture) and s._pic.blipFill.blip.rEmbed and s.image.ext not in ("wmf", "emf", "svg"):
            iw, ih = s.image.size
            shown = (iw * (1 - s.crop_left - s.crop_right)) / max(ih * (1 - s.crop_top - s.crop_bottom), 1)
            frame = (s.width * sx) / max(s.height * sy, 1)
            if abs(shown - frame) / frame > 0.05:
                issues.append(f"[{s.shape_id}] {s.name} 图片变形：内容比例 {shown:.2f}，图框比例 {frame:.2f}")
            if iw / max(s.width * sx / 914400, 0.01) < 96:
                issues.append(f"[{s.shape_id}] {s.name} 分辨率偏低：{iw}px 显示 {s.width * sx / 914400:.2f} 英寸")
        if not s.has_text_frame or not s.text_frame.text.strip():
            continue
        text = s.text_frame.text.strip()
        for word in ban:
            if word in text:
                issues.append(f"[{s.shape_id}] 包含禁用词“{word}”：{text[:40]}")
        if text == REVIEW_MARK:
            issues.append("存在待删除的审查标记")
            continue
        height, fills = estimate(s, sx, sy)
        if height > s.height * sy * 1.1 and s.height * sy > 91440:
            issues.append(f"[{s.shape_id}] 文字可能溢出：估算高 {height / 914400:.2f} 英寸，框高 {s.height * sy / 914400:.2f} 英寸：{text[:30]}")
        # 关键句框：宽度接近整页且 21pt 左右
        para = s.text_frame.paragraphs[0]
        if s.width * sx > prs.slide_width * 0.8 and 18 <= run_size(para, 0) <= 24 and fills:
            lines, last = fills[0]
            if lines > 2:
                issues.append(f"[{s.shape_id}] 关键句超过两行：{text[:30]}")
            elif lines == 2 and last < 0.7:
                issues.append(f"[{s.shape_id}] 关键句第二行只占 {last:.0%}，应缩为一行或写满两行：{text[:30]}")
    return issues


def main():
    parser = argparse.ArgumentParser(description="估算检查 PPT 的文字溢出、图片变形、残留文字和审查标记")
    parser.add_argument("deck", nargs="?")
    parser.add_argument("--pages", help="页码，例如 20-37；省略时检查全部页")
    parser.add_argument("--ban", default="", help="禁用词，用英文逗号分隔，例如 旧客户名,老师姓名,团队")
    parser.add_argument("--count", help="只统计一句话的字数和宽度")
    args = parser.parse_args()
    if args.count:
        count_text(args.count)
        return
    prs = Presentation(args.deck)
    pages = parse_pages(args.pages) if args.pages else set(range(1, len(prs.slides) + 1))
    ban = [w for w in args.ban.split(",") if w]
    total = 0
    for i, slide in enumerate(prs.slides, 1):
        if i in pages:
            for issue in check_slide(prs, slide, i, ban):
                print(f"第 {i} 页 {issue}")
                total += 1
    print(f"检查 {len(pages)} 页，发现 {total} 个问题。结果是估算值，字体和换行仍以 WPS 或 PowerPoint 打开为准。")


if __name__ == "__main__":
    main()

import argparse
import copy
from io import BytesIO
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.shapes.group import GroupShape
from pptx.util import Inches, Pt

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
REVIEW_MARK = "【待审查】本页已修改，审查后删除"


def walk(shapes):
    for s in shapes:
        yield s
        if isinstance(s, GroupShape):
            yield from walk(s.shapes)


def shape(slide, shape_id):
    found = [s for s in walk(slide.shapes) if s.shape_id == shape_id]
    if len(found) != 1:
        raise ValueError(f"shape_id={shape_id} 匹配到 {len(found)} 个形状")
    return found[0]


def set_text(slide, shape_id, text):
    # 只保留第一段第一个 run 的格式，用于标题、关键句和单段说明
    tf = shape(slide, shape_id).text_frame
    paras = tf.paragraphs
    runs = paras[0].runs
    if not runs:
        raise ValueError(f"shape_id={shape_id} 没有可继承格式的文字")
    runs[0].text = text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)
    for p in paras[1:]:
        p._p.getparent().remove(p._p)


def set_lines(slide, shape_id, lines):
    # 多段文本框：有文字的段数必须相同，每段保留原有格式和开头空格，空段不动
    paras = [p for p in shape(slide, shape_id).text_frame.paragraphs if p.runs]
    if len(paras) != len(lines):
        raise ValueError(f"shape_id={shape_id} 原有 {len(paras)} 段文字，传入 {len(lines)} 段")
    for p, line in zip(paras, lines):
        runs = p.runs
        old = "".join(r.text for r in runs)
        runs[0].text = old[: len(old) - len(old.lstrip())] + line
        for r in runs[1:]:
            r._r.getparent().remove(r._r)


def _image_bytes(img):
    return img if isinstance(img, bytes) else Path(img).read_bytes()


def put_image(slide, shape_id, img, box=None):
    # 在图框内等比缩放并居中，完整保留图片内容
    pic = shape(slide, shape_id)
    data = _image_bytes(img)
    iw, ih = Image.open(BytesIO(data)).size
    left, top, w, h = box or (pic.left, pic.top, pic.width, pic.height)
    if w / h > iw / ih:
        nh, nw = h, int(h * iw / ih)
    else:
        nw, nh = w, int(w * ih / iw)
    _, rid = slide.part.get_or_add_image_part(BytesIO(data))
    pic._element.blipFill.blip.rEmbed = rid
    pic.crop_left = pic.crop_right = pic.crop_top = pic.crop_bottom = 0
    pic.left, pic.top = int(left + (w - nw) / 2), int(top + (h - nh) / 2)
    pic.width, pic.height = nw, nh
    prune_unused_images(slide)


def fill_image(slide, shape_id, img):
    # 铺满原图框并居中裁切，图框大小不变
    pic = shape(slide, shape_id)
    data = _image_bytes(img)
    iw, ih = Image.open(BytesIO(data)).size
    _, rid = slide.part.get_or_add_image_part(BytesIO(data))
    pic._element.blipFill.blip.rEmbed = rid
    frame, image = pic.width / pic.height, iw / ih
    pic.crop_left = pic.crop_right = max(0.0, (1 - frame / image) / 2)
    pic.crop_top = pic.crop_bottom = max(0.0, (1 - image / frame) / 2)
    prune_unused_images(slide)


def remove_shape(slide, shape_id):
    el = shape(slide, shape_id)._element
    el.getparent().remove(el)
    strip_timing(slide)
    prune_unused_images(slide)


def prune_unused_images(slide):
    # 换图或删图后旧图仍留在关系表里，文件不会变小，旧内容也还在文件中，需要显式清除
    used = {
        value
        for node in slide._element.iter()
        for attr, value in node.attrib.items()
        if attr.startswith(f"{{{R_NS}}}")
    }
    for rid, rel in list(slide.part.rels.items()):
        if rel.reltype == RT.IMAGE and rid not in used:
            slide.part.rels.pop(rid)


def strip_timing(slide):
    # 动画引用已删除的形状时 PowerPoint 会提示修复，删除形状后一并清除动画
    timing = slide._element.find(f"{{{P_NS}}}timing")
    if timing is not None:
        slide._element.remove(timing)


def copy_shape(src_slide, shape_id, dst_slide):
    # 跨页复制形状或图片，并重新指向目标页的图片关系
    el = copy.deepcopy(shape(src_slide, shape_id)._element)
    for node in el.iter():
        for attr, value in list(node.attrib.items()):
            if attr.startswith(f"{{{R_NS}}}"):
                rel = src_slide.part.rels[value]
                node.attrib[attr] = dst_slide.part.relate_to(rel.target_part, rel.reltype)
    used = [s.shape_id for s in walk(dst_slide.shapes)]
    next(el.iter(f"{{{P_NS}}}cNvPr")).set("id", str(max(used + [1]) + 1))
    dst_slide.shapes._spTree.append(el)
    return dst_slide.shapes[-1]


def duplicate_slide(prs, src):
    # 新页追加在末尾，用 reorder 调整位置；动画和备注不复制
    new = prs.slides.add_slide(src.slide_layout)
    for ph in list(new.placeholders):
        ph._element.getparent().remove(ph._element)
    rid_map = {}
    for rid, rel in src.part.rels.items():
        if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
            continue
        target = rel.target_ref if rel.is_external else rel.target_part
        rid_map[rid] = new.part.relate_to(target, rel.reltype, is_external=rel.is_external)
    bg = src._element.cSld.find(f"{{{P_NS}}}bg")
    if bg is not None:
        new._element.cSld.insert(0, copy.deepcopy(bg))
    tree = new.shapes._spTree
    for el in src.shapes._spTree.iterchildren():
        if not (el.tag.endswith("}nvGrpSpPr") or el.tag.endswith("}grpSpPr")):
            tree.append(copy.deepcopy(el))
    for node in tree.iter():
        for attr, value in list(node.attrib.items()):
            if attr.startswith(f"{{{R_NS}}}") and value in rid_map:
                node.attrib[attr] = rid_map[value]
    return new


def delete_slides(prs, slides):
    lst = prs.slides._sldIdLst
    by_id = {el.id: el for el in lst}
    for s in slides:
        el = by_id[s.slide_id]
        prs.part.drop_rel(el.rId)
        lst.remove(el)
    _sync_sections(prs)


def reorder(prs, slides):
    # slides 必须列出全部保留的页面
    lst = prs.slides._sldIdLst
    by_id = {el.id: el for el in lst}
    if {s.slide_id for s in slides} != set(by_id):
        raise ValueError("新顺序必须正好包含当前全部页面")
    for s in slides:
        el = by_id[s.slide_id]
        lst.remove(el)
        lst.append(el)


def _sync_sections(prs):
    # 分节信息里残留已删页面时 PowerPoint 会提示修复
    ids = {el.id for el in prs.slides._sldIdLst}
    for node in list(prs.part._element.iter(f"{{{P14_NS}}}sldId")):
        if int(node.get("id")) not in ids:
            node.getparent().remove(node)


def add_review_mark(slide):
    tb = slide.shapes.add_textbox(Inches(0.15), Inches(7.08), Inches(3.4), Inches(0.32))
    run = tb.text_frame.paragraphs[0].add_run()
    run.text = REVIEW_MARK
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0xE0, 0x1E, 0x1E)


def clear_review_marks(prs):
    removed = []
    for i, slide in enumerate(prs.slides, 1):
        for s in list(slide.shapes):
            if s.has_text_frame and s.text_frame.text == REVIEW_MARK:
                s._element.getparent().remove(s._element)
                removed.append(i)
    return removed


def keep_only(prs, page_numbers):
    slides = list(prs.slides)
    delete_slides(prs, [s for i, s in enumerate(slides, 1) if i not in page_numbers])


def parse_pages(spec):
    pages = set()
    for part in spec.split(","):
        a, _, b = part.partition("-")
        pages.update(range(int(a), int(b or a) + 1))
    return pages


def main():
    parser = argparse.ArgumentParser(description="PPT 页面操作的命令行入口；复杂修改请 import 本文件写脚本")
    sub = parser.add_subparsers(dest="cmd", required=True)
    clear = sub.add_parser("clear-marks", help="删除全部审查标记并覆盖保存")
    clear.add_argument("deck")
    subset = sub.add_parser("subset", help="只保留指定页并另存，例如 20-37,41")
    subset.add_argument("deck")
    subset.add_argument("pages")
    subset.add_argument("output")
    args = parser.parse_args()
    prs = Presentation(args.deck)
    if args.cmd == "clear-marks":
        print("已删除审查标记的页码：", clear_review_marks(prs))
        prs.save(args.deck)
    else:
        keep_only(prs, parse_pages(args.pages))
        prs.save(args.output)
        print("已保存", args.output, "共", len(prs.slides), "页")


if __name__ == "__main__":
    main()

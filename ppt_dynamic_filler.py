#!/usr/bin/env python3
"""
按「列表嵌套字典」动态拼装 PPT：从模板中按占位符键集合匹配版式页，
复制到文末并替换，最后删除全部模板页。

复制页会对 r:embed / r:link / r:id 等关系 ID 做重映射，避免仅 deepcopy XML 时图片/媒体引用指向错误。

依赖模板中每一类版式页的 {占位符} 键集合唯一，且与 JSON 每条 dict 的键集合一致（可通过 key_aliases 映射）。
"""

from __future__ import annotations

import re
from copy import deepcopy
from typing import Any, Dict, FrozenSet, Iterable, List, Mapping, Optional, Tuple

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import nsuri
from pptx.slide import Slide

from ppt_filler import fill_slide_placeholders, iter_shapes_with_text_frame

# 关系命名空间下的属性（r:embed、r:link、r:id 等）值形如 rIdN，需映射到新幻灯片部件的 rels
_REL_NS_PREFIX = "{" + nsuri("r") + "}"

# 示例 JSON 若使用「目录占位符N」，可映射到模板真实键名「目录内容占位符N」
DEFAULT_KEY_ALIASES: Dict[str, str] = {
    "目录占位符1": "目录内容占位符1",
    "目录占位符2": "目录内容占位符2",
    "目录占位符3": "目录内容占位符3",
}


def _collect_r_ids_in_element(root: Any) -> List[str]:
    """收集元素树中所有关系命名空间属性上的 rId* 引用（去重且保持顺序）。"""
    seen: set[str] = set()
    ordered: List[str] = []
    for el in root.iter():
        for attr_name, attr_val in el.attrib.items():
            if not attr_name.startswith(_REL_NS_PREFIX):
                continue
            if not attr_val or not attr_val.startswith("rId"):
                continue
            if attr_val not in seen:
                seen.add(attr_val)
                ordered.append(attr_val)
    return ordered


def _remap_copied_relationships(source_part: Any, dest_part: Any, root: Any) -> None:
    """
    deepcopy 的形状仍携带源页上的 rId，与目标 slide 部件的 _rels 不一致会导致图片/图表丢失。
    按源部件解析每个旧 rId，在目标部件上 relate_to 同一目标，再把 XML 中的 rId 替换为新值。
    """
    old_ids = _collect_r_ids_in_element(root)
    id_map: Dict[str, str] = {}
    for old in old_ids:
        if old not in source_part.rels:
            continue
        rel = source_part.rels[old]
        if rel.is_external:
            new_id = dest_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            new_id = dest_part.relate_to(rel.target_part, rel.reltype)
        id_map[old] = new_id
    if not id_map:
        return
    for el in root.iter():
        for attr_name, attr_val in list(el.attrib.items()):
            if attr_val in id_map:
                el.set(attr_name, id_map[attr_val])


def duplicate_slide(pres: Presentation, slide: Slide) -> Slide:
    """复制一页幻灯片（deepcopy 形状并修复 r:embed 等关系，避免图片/图表引用断裂）。"""
    layout = slide.slide_layout
    new_slide = pres.slides.add_slide(layout)
    for shape in list(new_slide.shapes):
        shape.element.getparent().remove(shape.element)
    for shape in slide.shapes:
        new_el = deepcopy(shape.element)
        new_slide.shapes._spTree.insert_element_before(new_el, "p:extLst")
    _remap_copied_relationships(slide.part, new_slide.part, new_slide.element)
    return new_slide


def delete_slide(pres: Presentation, slide: Slide) -> None:
    """从演示文稿中删除指定页（先移 sldId 再 drop 关系）。"""
    r_id = None
    for rel in pres.part.rels.values():
        if rel.target_part == slide.part:
            r_id = rel.rId
            break
    if not r_id:
        raise ValueError("无法在演示文稿关系中定位该幻灯片")
    sid = slide.slide_id
    sld_id_lst = pres.slides._sldIdLst
    for sld_id in list(sld_id_lst):
        if sld_id.id == sid:
            sld_id_lst.remove(sld_id)
            break
    else:
        raise ValueError("无法在 sldIdLst 中定位该幻灯片")
    pres.part.drop_rel(r_id)


def placeholder_keys_on_slide(slide: Slide) -> FrozenSet[str]:
    """收集该页全部 {key} 中的 key（含组合形状内文本框）。"""
    texts: List[str] = []
    for shape in iter_shapes_with_text_frame(slide.shapes):
        if shape.text:
            texts.append(shape.text)
    full = "".join(texts)
    return frozenset(re.findall(r"\{([^}]+)\}", full))


def _normalize_item(
    item: Mapping[str, Any], key_aliases: Optional[Mapping[str, str]]
) -> Dict[str, Any]:
    """将输入 dict 的键按别名映射为模板中的键名。"""
    aliases = key_aliases or {}
    return {aliases.get(k, k): v for k, v in item.items()}


def _build_template_keymap(
    pres: Presentation, template_count: int
) -> Dict[FrozenSet[str], Slide]:
    """前 template_count 页为模板页：键集合 -> 代表 Slide（后者用于 duplicate）。"""
    keymap: Dict[FrozenSet[str], Slide] = {}
    for i in range(template_count):
        slide = pres.slides[i]
        ks = placeholder_keys_on_slide(slide)
        if not ks:
            continue
        if ks in keymap:
            raise ValueError(
                f"模板第 {i + 1} 页与前面某页的占位符键集合重复: {sorted(ks)}"
            )
        keymap[ks] = slide
    return keymap


def generate_dynamic_ppt(
    template_path: str,
    output_path: str,
    pages: Iterable[Mapping[str, Any]],
    *,
    key_aliases: Optional[Mapping[str, str]] = None,
) -> Tuple[int, int]:
    """
    按 pages 顺序生成 PPT：每页 dict 的键集合需匹配某一模板页的占位符键集合。

    Args:
        template_path: 含「模板页」的 pptx（通常全部版式页连续排在文件前部）
        output_path: 输出路径
        pages: 列表/可迭代，每项为占位符 -> 值
        key_aliases: 输入键名 -> 模板键名，会与 DEFAULT_KEY_ALIASES 合并（传入优先覆盖）

    Returns:
        (生成的内容页数量, 累计替换占位符次数)
    """
    merged_aliases: Dict[str, str] = {**DEFAULT_KEY_ALIASES}
    if key_aliases:
        merged_aliases.update(dict(key_aliases))

    prs = Presentation(template_path)
    template_count = len(prs.slides)
    if template_count == 0:
        raise ValueError("模板无幻灯片")

    keymap = _build_template_keymap(prs, template_count)
    total_replacements = 0
    n_out = 0

    for raw in pages:
        item = _normalize_item(raw, merged_aliases)
        ks = frozenset(item.keys())
        if ks not in keymap:
            raise KeyError(
                "无匹配模板页，键集合为 "
                f"{sorted(ks)}；"
                f"可用模板键集合示例: {[sorted(k) for k in list(keymap.keys())[:3]]}..."
            )
        src = keymap[ks]
        new_slide = duplicate_slide(prs, src)
        total_replacements += fill_slide_placeholders(new_slide, item)
        n_out += 1

    for _ in range(template_count):
        delete_slide(prs, prs.slides[0])

    prs.save(output_path)
    return n_out, total_replacements

if __name__ == "__main__":
    generate_dynamic_ppt(
        template_path="templates/述职报告1.pptx",
        output_path="output3.pptx",
        pages=[
    {
        "主标题": "2026年第一季度个人总结"
    },
    {
        "目录内容占位符1": "个人情况回顾",
        "目录内容占位符2": "存在问题与不足",
        "目录内容占位符3": "下一步个人计划"
    },
    {
        "章节序号": "01",
        "章节标题": "个人情况回顾"
    },
    {
        "页面标题（1框）": "一季度个人总体情况",
        "小标题占位符1": "学习与成长成效",
        "内容占位符1": "本季度我坚持以积极进取的态度投入学习生活，认真贯彻学校各项要求，围绕学业目标、品德修养、综合能力提升扎实推进自我提升。累计参与主题学习活动3次，学习心得分享会1次，专题学习2次，各项学习活动参与率达95%以上。"
    },
    {
        "章节序号": "02",
        "章节标题": "存在问题与不足"
    },
    {
        "页面标题（2框）": "当前个人短板分析",
        "小标题占位符1": "理论学习方面",
        "内容占位符1": "部分知识学习深度不够，存在学用脱节现象；学习方法较为单一，自主学习的创新性和主动性有待提升。",
        "小标题占位符2": "自我提升方面",
        "内容占位符2": "能力提升进度偏慢；个别时候自我要求不够严格，服务同学、奉献集体的意识需进一步增强。"
    },
    {
        "章节序号": "03",
        "章节标题": "下一步工作计划"
    },
    {
        "页面标题（3框）": "二季度个人重点计划",
        "小标题占位符1": "深化知识学习",
        "内容占位符1": "制定详细学习计划，开展\"书香个人\"建设，推动知识学习入脑入心、学以致用。",
        "小标题占位符2": "强化自我管理",
        "内容占位符2": "规范日常学习生活习惯，推进个人能力标准化提升，争做优秀学生榜样。",
        "小标题占位符3": "服务集体与实践",
        "内容占位符3": "积极参与\"优秀标兵\"创建活动，推动个人发展与集体事务深度融合，以高标准要求促进全面成长。"
    },
    {
        "页面标题（4框）": "保障措施与自我要求",
        "小标题占位符1": "压实自我责任",
        "内容占位符1": "严格履行个人主体责任，主动落实各项学习与提升任务。",
        "小标题占位符2": "完善自我考核",
        "内容占位符2": "建立个人提升清单，实行台账管理，定期自我检查。",
        "小标题占位符3": "加强总结反思",
        "内容占位符3": "及时总结学习生活中的好经验好做法，营造比学赶超的良好氛围。",
        "小标题占位符4": "注重成果转化",
        "内容占位符4": "把学习成果转化为实践能力，以实际表现和成绩检验个人提升成效。"
    },
    {
        "结束语": "谢谢！"
    }
]
    )
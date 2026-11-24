"""Inspect a bid or apply a source-bound plan as native Word revisions and comments."""

import argparse
import copy
import hashlib
import io
import json
import os
from pathlib import Path
import posixpath
import re
import sys
import tempfile
import unicodedata
import uuid
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime, timezone
from itertools import combinations

from lxml import etree as ET


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
OFF = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
COMMENTS_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
DOC = "word/document.xml"
RELS = "word/_rels/document.xml.rels"
TYPES = "[Content_Types].xml"
SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def w(name):
    return f"{{{W}}}{name}"


def parse(data):
    parser = ET.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    root = ET.fromstring(data, parser)
    if root.getroottree().docinfo.doctype:
        raise ValueError("不支持含 DTD 的 XML。")
    return root


def xml_bytes(root):
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def child_value(node, *names):
    if node is None:
        return ""
    for name in names:
        node = node.find(w(name))
        if node is None:
            return ""
    return node.get(w("val"), "")


def text_node(name, text):
    elem = ET.Element(w(name))
    elem.text = text
    if text[:1].isspace() or text[-1:].isspace():
        elem.set(SPACE, "preserve")
    return elem


def xml_string(value):
    if not isinstance(value, str):
        raise ValueError("计划中的正文和标识必须是字符串。")
    probe = ET.Element("probe")
    probe.text = value
    return value


def atomic_create(path, write):
    """Publish a fully written file without allowing an overwrite race."""
    path = Path(path)
    if os.path.lexists(path):
        raise ValueError(f"输出已存在：{path}")
    with tempfile.NamedTemporaryFile(dir=path.parent, prefix=".bid-polish-", delete=False) as temp:
        temp_path = Path(temp.name)
    try:
        write(temp_path)
        if os.name == "nt":
            os.rename(temp_path, path)
        else:
            os.link(temp_path, path)
    finally:
        temp_path.unlink(missing_ok=True)


@dataclass
class Span:
    start: int
    end: int
    node: object
    reason: str
    epoch: int


@dataclass
class Paragraph:
    id: str
    elem: object
    style: str
    section: str
    in_table: bool
    lock: str
    starts_safe: bool
    ends_safe: bool = False
    text: str = ""
    spans: list = field(default_factory=list)
    text_chunks: list = field(default_factory=list)
    length: int = 0

    def append(self, text, node, reason, epoch):
        if text:
            self.text_chunks.append(text)
            self.spans.append(Span(self.length, self.length + len(text), node, reason, epoch))
            self.length += len(text)

    def finish(self):
        self.text = "".join(self.text_chunks)
        self.text_chunks.clear()

    def public(self):
        flags = sorted({s.reason for s in self.spans if s.reason})
        return {"id": self.id, "text": self.text, "style": self.style, "section": self.section,
                "in_table": self.in_table, "protected": bool(self.lock), "reason": self.lock,
                "protected_spans": flags}


CONTAINERS = {
    w(name) for name in ("sdt", "ins", "del", "moveFrom", "moveTo", "txbxContent",
                         "customXml", "smartTag", "fldSimple", "hyperlink")
}
RANGE_TYPES = ("bookmark", "commentRange", "moveFromRange", "moveToRange",
               "customXmlInsRange", "customXmlDelRange", "customXmlMoveFromRange",
               "customXmlMoveToRange", "perm")
RANGE_START = {w(name + "Start"): name for name in RANGE_TYPES}
RANGE_END = {w(name + "End"): name for name in RANGE_TYPES}
BOUNDARIES = {w(name) for name in ("fldChar", "instrText", "proofErr", "drawing", "pict",
              "object", "footnoteReference", "endnoteReference", "commentReference", "br",
              "tab", "cr", "sym", "lastRenderedPageBreak")}
VISIBLE = {w("t"): None, w("tab"): "\t", w("br"): "\n", w("cr"): "\n",
           w("noBreakHyphen"): "\u2011", w("softHyphen"): "\u00ad"}
PLAIN_RUN_CHILDREN = {w("rPr"), w("t")}
HIDDEN = {w("del"), w("moveFrom")}


def is_plain(run, paragraph):
    return (run is not None and run.tag == w("r") and run.getparent() is paragraph
            and len(run.findall(w("t"))) == 1 and all(c.tag in PLAIN_RUN_CHILDREN for c in run))


def run_format(run):
    props = run.find(w("rPr"))
    return ET.tostring(props, method="c14n", exclusive=True) if props is not None else b""


def heading_level(p, styles):
    props = p.find(w("pPr"))
    outline = child_value(props, "outlineLvl")
    style_id = child_value(props, "pStyle")
    seen = set()
    while style_id and style_id not in seen:
        seen.add(style_id)
        style = styles.get(style_id)
        label = child_value(style, "name")
        match = re.search(r"(?:Heading|标题)\s*([1-9])$", label or style_id, re.I)
        if match and not outline:
            return int(match.group(1))
        outline = outline or child_value(style, "pPr", "outlineLvl")
        style_id = child_value(style, "basedOn")
    return int(outline) + 1 if outline in tuple(str(n) for n in range(9)) else 0


def paragraph_lock(p, level, style_label):
    for a in p.iterancestors():
        if a.tag in CONTAINERS or a.tag in (w("drawing"), w("pict"), w("object"),
                "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent"):
            return "位于受保护容器"
        if a.tag in (w("tr"), w("tc")):
            props = a.find(w("trPr") if a.tag == w("tr") else w("tcPr"))
            if props is not None and any(c.tag in {w(n) for n in ("ins", "del", "trPrChange",
                    "tcPrChange", "cellIns", "cellDel", "cellMerge")} for c in props):
                return "位于受保护容器"
    if level or re.fullmatch(r"(?:Title|Subtitle|Caption|TOC\s*\d+|题注|标题|副标题|目录\s*\d*)",
                             style_label, re.I):
        return "标题、目录或题注"
    props = p.find(w("pPr"))
    if props is not None and any(e.tag in (w("ins"), w("del"), w("pPrChange"), w("moveFrom"), w("moveTo"))
                                 for e in props.iter()):
        return "段落属性含既有修订"
    return ""


def collect_paragraphs(doc, styles):
    records, stack, ranges, heading_path = [], [], set(), {}
    depth = epoch = hidden = 0
    body = doc.find(w("body"))
    if body is None:
        raise ValueError("不支持此文档的 Word XML 格式。")
    for event, node in ET.iterwalk(body, events=("start", "end")):
        tag = node.tag
        if event == "end":
            if tag == w("p"):
                rec, level = stack.pop()
                rec.ends_safe = not (depth or ranges)
                rec.finish()
                if level and rec.text.strip() and not rec.in_table:
                    for key in list(heading_path):
                        if key >= level:
                            del heading_path[key]
                    heading_path[level] = rec.text.strip()
                    rec.section = " / ".join(heading_path.values())
            if tag in HIDDEN:
                hidden -= 1
            continue
        if tag in HIDDEN:
            hidden += 1
        if tag == w("p"):
            style_id = child_value(node, "pPr", "pStyle")
            label = child_value(styles.get(style_id), "name") or style_id
            level = heading_level(node, styles)
            rec = Paragraph(f"P{len(records) + 1:04d}", node, label, " / ".join(heading_path.values()),
                            any(a.tag == w("tc") for a in node.iterancestors()),
                            paragraph_lock(node, level, label), not (ranges or depth))
            records.append(rec)
            stack.append((rec, level))
        if hidden:
            continue
        if tag in RANGE_START:
            ranges.add((RANGE_START[tag], node.get(w("id"))))
            epoch += 1
        elif tag in RANGE_END:
            ranges.discard((RANGE_END[tag], node.get(w("id"))))
            epoch += 1
        elif tag in BOUNDARIES:
            epoch += 1
            if tag == w("fldChar"):
                field_type = node.get(w("fldCharType"))
                depth = max(0, depth + (field_type == "begin") - (field_type == "end"))
        if tag not in VISIBLE or not stack:
            continue
        rec = stack[-1][0]
        reason = ""
        if depth:
            reason = "域或其缓存文字"
        elif ranges:
            reason = "书签、批注或受保护范围"
        elif tag != w("t") or not is_plain(node.getparent(), rec.elem):
            reason = "复杂控件、对象或非普通文字"
        elif node.getparent().find(".//" + w("rPrChange")) is not None:
            reason = "文字格式含既有修订"
        elif any(node.getparent().find(w("rPr") + "/" + w(flag)) is not None
                 for flag in ("vanish", "webHidden", "specVanish")):
            reason = "文字含隐藏属性"
        rec.append((node.text or "") if tag == w("t") else VISIBLE[tag], node, reason, epoch)
    return records


def locate(text, quote):
    if not quote:
        return None, "原句为空"
    first = text.find(quote)
    if first < 0:
        return None, "找不到原句"
    if text.find(quote, first + 1) >= 0:
        return None, "同段原句出现多次"
    return (first, first + len(quote)), ""


# Compare a proposed replacement in its original paragraph. A plan can select
# just "30" or "电话", leaving its unit or service action outside the selected span.
CN_NUMBER = r"[零〇一二两三四五六七八九十百千万亿半]+"
PERIOD = r"工作小时|日历天|工作日|日历日|小时|分钟|秒钟|毫秒|季度|年|月|周|天|日|秒"
UNIT = (PERIOD + r"|人天|人日|万元|元|平方米|公里|米|套|台|个|名|家|份|项|级|次|人|%|‰"
        + r"|(?:[KMGT]i?B(?:/s)?|[KMGT]bps|Gb/s|Mb/s|ms|min|h|s)(?![A-Za-z])")
NUMBER = r"[0-9]+(?:[.,][0-9]+)*"
MODEL = r"(?<![A-Za-z0-9])(?:[A-Za-z]+[-_]?[0-9][A-Za-z0-9_-]*|RPO|RTO|SLA|MTTR|MTBF)(?![A-Za-z0-9])"
TECH_MARKERS = re.compile(
    rf"(?P<model>{MODEL})|(?P<amount>{NUMBER}\s*(?:{UNIT})?|{CN_NUMBER}\s*(?:{UNIT}))"
    r"(?P<limit>以内|之内|内|以上|以下|之前|之后)?", re.I)
PERIOD_MARKER = re.compile(rf"(?:{NUMBER}|{CN_NUMBER})\s*(?:{PERIOD}|(?:min|ms|h|s)(?![A-Za-z]))", re.I)
HARD_BREAK = re.compile(r"[。；;！？!?\n]")
GOVERNANCE = {"提交", "报送", "发出", "出具", "审核", "审批", "签批", "复核", "确认", "执行", "实施", "登记"}
PREDICATE = re.compile(r"(?:(电话|书面|现场|远程|在线)\s*(?:进行|作出|予以)?\s*)?("
                       + "|".join(sorted(GOVERNANCE)) + r"|到达现场|抵达现场|到场|响应|修复|恢复)")
ACTOR = r"投标人|投标方|我方|采购人|采购方|委托方|委方|甲方|乙方|项目经理|运维人员|工程师|供应商|厂商|业主|业务主管|服务台"
SUBJECT = re.compile(ACTOR)
LOGIC = re.compile(r"(?:仅当|仅在|只有)([^。；;！？!?\n]{1,100}?)(?:时|方可|才能|才可|才允许)")
LIMIT_GROUPS = {
    "lower_inclusive": ("不少于", "不低于", "至少", "大于等于", ">=", "≥"),
    "upper_inclusive": ("不超过", "不高于", "不大于", "至多", "小于等于", "<=", "≤"),
    "not_before": ("不早于",),
    "not_after": ("不晚于",),
    "lower_exclusive": ("超过", "高于", "大于", ">"),
    "upper_exclusive": ("少于", "低于", "小于", "<"),
}
LIMIT_MAP = {term: family for family, terms in LIMIT_GROUPS.items() for term in terms}
LIMIT_MARKER = re.compile("|".join(map(re.escape, sorted(LIMIT_MAP, key=len, reverse=True))))
# Match contractual negatives, leaving rhetorical phrases like “不仅” alone.
NEGATION = re.compile(r"不得|严禁|禁止|未经|无须|无需|无权|不能|不应|不予|不再|不计入|不纳入|不允许|不包含|不包括|不含"
                      r"|不(?=提供|承担|负责|配合|协助|支持|适用|覆盖|影响|中断|停机|响应|到场|修复|恢复|接入|启用)"
                      r"|未(?=经|获|取得|授权|确认|批准|完成|验证|就绪)|仅限|除外")
ROLE = r"负责(?!人(?!员)|制)|配合(?!度)|协助"
ROLE_WORD = re.compile(ROLE)
STATED_ROLE = re.compile(rf"({ACTOR})(?:将|应当|应|须|需)?({ROLE})")


def predicates(clause):
    for match in PREDICATE.finditer(clause):
        mode, action = match.groups()
        tail = clause[match.end():]
        if action in {"审核", "审批", "签批", "复核", "确认", "执行", "实施"}:
            if re.match(r"结果|意见|记录|状态|人员|进度|情况|依据|标准|环节", tail):
                continue
            if re.search(r"(?:提交|报送|编制|收到|依据|按|将|把|对)(?:的|该|上述)?$", clause[:match.start()]) and tail.startswith("方案"):
                continue
        if action in {"发出", "出具"} and tail.startswith("书面"):
            mode = "书面"
        yield match.start(), match.end(), (mode, "到场" if action in ("到达现场", "抵达现场") else action)


def bindings(text):
    deadlines, duties = [], Counter()
    # “在2小时内，提交” and “在2小时内提交” share a scope.
    text = re.sub(r"(?<=内)[,，]", "", text)
    for sentence in HARD_BREAK.split(text):
        clauses = re.split(r"[,，]|随后|另行", sentence)
        last_subject = None
        for clause in clauses:
            acts = list(predicates(clause))
            subjects = [m for m in SUBJECT.finditer(clause)
                        if not re.search(r"(?:向|给|按照|依据|根据)$", clause[:m.start()])]
            for start, end, (_, action) in acts:
                if action not in GOVERNANCE:
                    continue
                left = [s for s in subjects if s.end() <= start]
                if len(left) > 1:
                    last = left[-1]
                    if (re.search(r"(?:在|经|待|当|收到)$", clause[:last.start()])
                            and re.search(r"(?:后|前|时)", clause[last.end():start])):
                        # “我方将在采购人审批后实施” returns to the main subject after 后.
                        left.pop()
                owner = left[-1].group() if left else None
                inverse = re.match(rf"(?:工作|事项|任务)?由({ACTOR})(?:负责|承担)", clause[end:])
                if inverse:
                    owner = inverse.group(1)
                elif not owner and re.match(r"(?:再|并|同时|且)", clause):
                    owner = last_subject
                if owner:
                    duties[owner, action] += 1
            if subjects:
                last_subject = subjects[-1].group()
            periods = list(PERIOD_MARKER.finditer(clause))
            for i, period in enumerate(periods):
                stop = periods[i + 1].start() if i + 1 < len(periods) else len(clause)
                condition = re.search(r"(?<!未)(?:经|待)[^,，。；]{1,16}?后", clause[period.end():stop])
                if condition:
                    stop = period.end() + condition.start()
                scoped = [label for start, end, label in acts if period.end() <= start < stop]
                if not scoped:
                    # Also read postposed limits such as “电话响应时限为30分钟”.
                    for start, end, label in reversed(acts):
                        if end <= period.start() and re.fullmatch(
                                r"(?:不得超过|不超过|至多|至少|[≤≥<>]=?|[^,，。；]{0,12}?"
                                r"(?:时限|时间|用时|限时|期限|[应须需]在)(?:为|是|不超过|不得超过)?)",
                                clause[end:period.start()]):
                            scoped.append(label)
                            break
                deadlines.append((re.sub(r"\s+", "", period.group()), tuple(scoped)))
    return deadlines, duties


def role_terms(text):
    roles = set()
    for match in STATED_ROLE.finditer(text):
        # “由采购人负责审核” and “采购人审核” already have the same explicit predicate.
        if match.group(2) == "负责" and any(text.startswith(v, match.end()) for v in GOVERNANCE):
            continue
        roles.add(match.groups())
    return roles


def logical_operators(text):
    return [("或" if op.startswith("或") else "且") for match in LOGIC.finditer(text)
            for op in re.findall(r"并且|且|或者|或", match.group(1))]


def commitment_change(original, start, end, replacement):
    changed = original[:start] + replacement + original[end:]
    before, after = (unicodedata.normalize("NFKC", s) for s in (original, changed))

    def markers(text):
        return [re.sub(r"\s+", "", match.group()) for match in TECH_MARKERS.finditer(text)]

    def limits(text):
        return [LIMIT_MAP[m.group()] for m in LIMIT_MARKER.finditer(text)]

    if markers(before) != markers(after):
        return "数值、单位、时限边界或技术指标标识有变，待确认"
    if limits(before) != limits(after) or NEGATION.findall(before) != NEGATION.findall(after):
        return "数值上下限、否定或适用范围的措辞有变，待确认"
    if logical_operators(before) != logical_operators(after):
        return "明确条件中的并列、择一关系有变，待确认"
    old_deadlines, old_duties = bindings(before)
    new_deadlines, new_duties = bindings(after)
    if old_deadlines != new_deadlines:
        return "时限对应的动作、响应方式或适用范围有变，待确认"
    if old_duties != new_duties:
        return "明确的责任主体与办理动作对应关系有变，待确认"
    old_roles, new_roles = set(ROLE_WORD.findall(before)), set(ROLE_WORD.findall(after))
    if (role_terms(before) != role_terms(after)
            or (old_roles - new_roles and new_roles - old_roles)):
        return "责任主体与负责、配合、协助的分工有变，待确认"
    return ""


def combined_risks(text, todo, conflicts):
    """Check partial acceptance exhaustively for small groups and selectively for long ones."""
    eligible = [i for i in range(len(todo)) if i not in conflicts]
    groups = []
    for i in eligible:
        gap = text[todo[groups[-1][-1]][0][1]:todo[i][0][0]] if groups else ""
        # A hard stop that no candidate touches remains a stop in every subset.
        if not groups or HARD_BREAK.search(gap):
            groups.append([])
        groups[-1].append(i)
    risks = {}
    for group in groups:
        if len(group) < 2:
            continue
        first, last = todo[group[0]][0][0], todo[group[-1]][0][1]
        left = max((m.end() for m in HARD_BREAK.finditer(text, 0, first)), default=0)
        next_break = HARD_BREAK.search(text, last)
        right = next_break.start() if next_break else len(text)
        original = text[left:right]
        def difference(selection):
            pieces, pos = [], left
            for i in selection:
                (start, end), row = todo[i]
                pieces.extend((text[pos:start], row["after"]))
                pos = end
            pieces.append(text[pos:right])
            return commitment_change(original, 0, len(original), "".join(pieces))

        def check(selection):
            if not any(i in risks for i in selection):
                reason = difference(selection)
                if reason:
                    where = "本句" if len(selection) == 1 else "同句多处修改一起采纳时，"
                    risks.update((i, where + reason) for i in selection)

        if len(group) <= 8:
            for size in range(1, len(group) + 1):
                for selection in combinations(group, size):
                    check(selection)
            continue
        # Check local interactions and the full result without enumerating 2**n
        # states. Nothing is held back unless a concrete difference is observed.
        for i in group:
            check((i,))
        for width in (2, 3):
            for j in range(len(group) - width + 1):
                check(group[j:j + width])
        keep = [i for i in group if i not in risks]
        while len(keep) > 1 and difference(keep):
            # Find an actual safe -> unsafe prefix transition. Disqualify only
            # the edit making that transition, then check the remaining full set.
            low, high = 0, len(keep)
            while high - low > 1:
                mid = (high + low) // 2
                if difference(keep[:mid]):
                    high = mid
                else:
                    low = mid
            reason = difference(keep[:high])
            index = keep.pop(high - 1)
            risks[index] = "与前面的修订一并采纳时，" + reason
    return risks


def safety(rec, start, end):
    if rec.lock:
        return rec.lock
    spans = [s for s in rec.spans if s.start < end and start < s.end]
    if not spans:
        return "缺少可定位文字"
    for span in spans:
        if span.reason:
            return span.reason
    if len({s.epoch for s in spans}) != 1:
        return "原句跨越 Word 锚点、域或对象"
    runs = [s.node.getparent() for s in spans]
    if len({run_format(run) for run in runs}) != 1:
        return "原句包含多种字符格式"
    if any(r1.getnext() is not r2 for r1, r2 in zip(runs, runs[1:])):
        return "原句跨越非连续文字节点"
    return ""


def live_segments(rec):
    """Offsets in the current accepted text, scoped to this paragraph only."""
    offset = 0
    for n in rec.elem.iter():
        if n.tag not in VISIBLE:
            continue
        ancestors = list(n.iterancestors())
        if next((a for a in ancestors if a.tag == w("p")), None) is not rec.elem:
            continue
        if any(a.tag in HIDDEN for a in ancestors):
            continue
        txt = (n.text or "") if n.tag == w("t") else VISIBLE[n.tag]
        if txt:
            yield offset, offset + len(txt), n
            offset += len(txt)


def set_text(node, value):
    node.text = value
    if value[:1].isspace() or value[-1:].isspace():
        node.set(SPACE, "preserve")


def split_at(rec, offset):
    for start, end, n in live_segments(rec):
        if start < offset < end:
            r = n.getparent()
            if not is_plain(r, rec.elem) or n.tag != w("t"):
                raise ValueError("不能拆分复杂文字节点")
            right = copy.deepcopy(r)
            value, at = n.text, offset - start
            set_text(n, value[:at])
            set_text(right.find(w("t")), value[at:])
            r.addnext(right)
            break


def target_runs(rec, start, end):
    current = [(s, e, n) for s, e, n in live_segments(rec) if s < end and start < e]
    if not current or any(n.tag != w("t") or not is_plain(n.getparent(), rec.elem) for _, _, n in current):
        return []
    runs = [n.getparent() for _, _, n in current]
    if len({run_format(r) for r in runs}) != 1 or any(a.getnext() is not b for a, b in zip(runs, runs[1:])):
        return []
    split_at(rec, end)
    split_at(rec, start)
    return [n.getparent() for s, e, n in live_segments(rec) if s < end and start < e]


def revision(kind, number, author, date):
    el = ET.Element(w(kind))
    for key, value in (("id", number), ("author", author), ("date", date)):
        el.set(w(key), str(value))
    return el


def track(runs, after, author, date, ids):
    parent = runs[0].getparent()
    position = parent.index(runs[0])
    deletion = revision("del", next(ids), author, date)
    replacement = copy.deepcopy(runs[0]) if after else None
    parent.insert(position, deletion)
    for run in runs:
        run.find(w("t")).tag = w("delText")
        deletion.append(run)
    used = [int(deletion.get(w("id")))]
    if replacement is not None:
        set_text(replacement.find(w("t")), after)
        insertion = revision("ins", next(ids), author, date)
        insertion.append(replacement)
        deletion.addnext(insertion)
        used.append(int(insertion.get(w("id"))))
    return used


def next_number(roots):
    largest = -1
    for root in roots:
        for node in root.iter():
            value = node.get(w("id"), "")
            if value.isascii() and value.isdecimal():
                largest = max(largest, int(value))
    while True:
        largest += 1
        yield largest


class CommentStore:
    def __init__(self, pkg):
        self.rels = parse(pkg.read(RELS)) if RELS in pkg.namelist() else ET.Element(
            f"{{{REL}}}Relationships", nsmap={None: REL})
        self.types = parse(pkg.read(TYPES))
        self.path, self.new_rel, self.new_type = None, False, False
        self.changed = False
        for rel in self.rels:
            if rel.get("Type") == OFF + "/comments":
                target = rel.get("Target", "")
                if rel.get("TargetMode") == "External" or not target:
                    raise ValueError("批注关系未指向文档内的部件。")
                self.path = posixpath.normpath(target.lstrip("/") if target.startswith("/")
                                               else posixpath.join("word", target))
                if self.path.startswith("../") or self.path not in pkg.namelist():
                    raise ValueError("原始批注部件缺失。")
                break
        if self.path is None:
            self.path = "word/comments.xml"
            if self.path in pkg.namelist():
                raise ValueError("文档有未连接的批注部件，无法安全修改。")
            used = {rel.get("Id") for rel in self.rels}
            i = 1
            while f"rId{i}" in used:
                i += 1
            ET.SubElement(self.rels, f"{{{REL}}}Relationship", Id=f"rId{i}",
                          Type=OFF + "/comments", Target="comments.xml")
            self.new_rel = True
            self.root = ET.Element(w("comments"), nsmap={"w": W})
        else:
            self.root = parse(pkg.read(self.path))
        if not any(o.get("PartName") == "/" + self.path for o in self.types):
            ET.SubElement(self.types, f"{{{CT}}}Override", PartName="/" + self.path,
                          ContentType=COMMENTS_TYPE)
            self.new_type = True

    def add(self, number, author, date, text):
        self.changed = True
        comment = revision("comment", number, author, date)
        self.root.append(comment)
        for line in text.splitlines() or [""]:
            paragraph = ET.SubElement(comment, w("p"))
            ET.SubElement(paragraph, w("r")).append(text_node("t", line))

    def updates(self):
        if not self.changed:
            return {}
        parts = {self.path: xml_bytes(self.root)}
        if self.new_rel:
            parts[RELS] = xml_bytes(self.rels)
        if self.new_type:
            parts[TYPES] = xml_bytes(self.types)
        return parts


def markers(number):
    begin = ET.Element(w("commentRangeStart"))
    end = ET.Element(w("commentRangeEnd"))
    run = ET.Element(w("r"))
    ref = ET.SubElement(run, w("commentReference"))
    for elem in (begin, end, ref):
        elem.set(w("id"), str(number))
    return begin, end, run


def add_anchor(rec, span, number, records):
    begin, end, ref = markers(number)
    if span:
        runs = target_runs(rec, *span)
        if runs:
            runs[0].addprevious(begin)
            runs[-1].addnext(end)
            end.addnext(ref)
            return rec.id, "quote"
    # A nearby insertion point can safely refer readers to protected material.
    rec_index = int(rec.id[1:]) - 1
    for distance in range(len(records)):
        indices = (rec_index,) if not distance else (rec_index - distance, rec_index + distance)
        for index in indices:
            if index < 0 or index >= len(records):
                continue
            candidate = records[index]
            if candidate.lock == "位于受保护容器" or not (candidate.starts_safe or candidate.ends_safe):
                continue
            if candidate.ends_safe:
                candidate.elem.extend([begin, end, ref])
            else:
                at = 1 if candidate.elem.find(w("pPr")) is not None else 0
                for node in (begin, end, ref):
                    candidate.elem.insert(at, node)
                    at += 1
            return candidate.id, "paragraph" if candidate is rec else "nearby"
    return None, None


class Document:
    def __init__(self, path):
        raw = Path(path).read_bytes()
        self.sha = hashlib.sha256(raw).hexdigest()
        self.pkg = zipfile.ZipFile(io.BytesIO(raw))
        names = self.pkg.namelist()
        if len(set(names)) != len(names) or DOC not in names or TYPES not in names:
            raise ValueError("输入不是支持的 DOCX 包，或包含重名部件。")
        self.doc = parse(self.pkg.read(DOC))
        styles = {}
        if "word/styles.xml" in names:
            styles = {el.get(w("styleId")): el for el in parse(self.pkg.read("word/styles.xml")).findall(w("style"))}
        self.records = collect_paragraphs(self.doc, styles)

    def require_writable(self):
        if any(n.lower().startswith("_xmlsignatures/") for n in self.pkg.namelist()):
            raise ValueError("文档含数字签名，不能写入修订。")
        if "word/settings.xml" in self.pkg.namelist():
            settings = parse(self.pkg.read("word/settings.xml"))
            protection = settings.find(w("documentProtection"))
            if protection is not None and protection.get(w("enforcement"), "false").lower() in ("1", "true", "on"):
                raise ValueError("文档启用了编辑保护。")

    def write(self, path, updates):
        def write_to(temp):
            if not updates:
                temp.write_bytes(self.pkg.fp.getvalue())
                return
            with zipfile.ZipFile(temp, "w") as out:
                out.comment = self.pkg.comment
                for info in self.pkg.infolist():
                    data = updates[info.filename] if info.filename in updates else self.pkg.read(info)
                    out.writestr(info, data)
                for name in updates.keys() - set(self.pkg.namelist()):
                    out.writestr(name, updates[name], compress_type=zipfile.ZIP_DEFLATED)
        atomic_create(path, write_to)


def check_plan(plan, document):
    if not isinstance(plan, dict) or plan.get("source_sha256") != document.sha:
        raise ValueError("原稿 SHA-256 与计划不符，请重新 inspect。")
    for group, keys in (("edits", ("p", "before", "after")), ("comments", ("p", "quote", "text"))):
        rows = plan.get(group, [])
        if not isinstance(rows, list):
            raise ValueError(f"{group} 必须是数组。")
        for row in rows:
            if not isinstance(row, dict) or any(key not in row for key in keys):
                raise ValueError(f"{group} 项需要字段 {', '.join(keys)}。")
            for key in keys:
                xml_string(row[key])
            if group == "comments" and not row["text"].strip():
                raise ValueError("批注正文不能为空。")


def apply_plan(document, plan, output, author):
    check_plan(plan, document)
    document.require_writable()
    xml_string(author)
    if not author.strip():
        raise ValueError("修订作者不能为空。")
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    book = {p.id: p for p in document.records}
    store = CommentStore(document.pkg)
    roots = [document.doc, store.root]
    for name in document.pkg.namelist():
        if re.fullmatch(r"word/(?:header\d+|footer\d+|footnotes|endnotes)\.xml", name):
            roots.append(parse(document.pkg.read(name)))
    if any(node.get(w("author")) == author for root in roots for node in root.iter()):
        raise ValueError("该修订作者已在文件中出现，请使用新的批次作者。")
    ids = next_number(roots)
    report = {"source_sha256": document.sha, "output": str(Path(output).resolve()),
              "author": author, "edits": [], "comments": []}
    schedule, applied, notices = {}, {}, []
    dirty = False
    for item in plan.get("edits", []):
        row = {k: item[k] for k in ("p", "before", "after")}
        report["edits"].append(row)
        rec = book.get(row["p"])
        reason = "找不到原稿段落" if rec is None else ""
        span = None
        if rec:
            span, reason = locate(rec.text, row["before"])
        if span and row["before"] == row["after"]:
            row.update(status="unchanged", revision_ids=[])
            continue
        if span:
            reason = safety(rec, *span)
            if not reason:
                reason = commitment_change(rec.text, *span, row["after"])
        if any(c in row["after"] for c in "\r\n\t"):
            reason = "新文字含制表符或换行，须人工确认结构"
        if reason:
            row.update(status="skipped", reason=reason, revision_ids=[])
            if rec:
                notices.append((row, rec, f"未自动改动（{reason}）。原句：{row['before']}\n拟改为：{row['after']}"))
        else:
            schedule.setdefault(rec.id, []).append((span, row))
    for pid, todo in schedule.items():
        rec = book[pid]
        todo.sort(key=lambda pair: pair[0])
        conflicts = set()
        latest_end, owner = -1, -1
        for i, ((start, end), _) in enumerate(todo):
            if start < latest_end:
                conflicts.update((i, owner))
            if end > latest_end:
                latest_end, owner = end, i
        related = combined_risks(rec.text, todo, conflicts)
        for i in reversed(range(len(todo))):
            (start, end), row = todo[i]
            runs = [] if i in conflicts or i in related else target_runs(rec, start, end)
            if not runs:
                reason = "改动范围相互重叠" if i in conflicts else related.get(i, "不能安全拆分当前文字节点")
                row.update(status="skipped", reason=reason, revision_ids=[])
                notices.append((row, rec, f"未自动改动（{reason}）。原句：{row['before']}\n拟改为：{row['after']}"))
                continue
            row.update(status="applied", revision_ids=track(runs, row["after"], author, now, ids))
            applied.setdefault(pid, []).append((start, end, len(row["after"])))
            dirty = True
    for item in plan.get("comments", []):
        row = {k: item[k] for k in ("p", "quote", "text")}
        report["comments"].append(row)
        rec = book.get(row["p"])
        if rec is None:
            row.update(status="skipped", reason="找不到原稿段落")
            continue
        notices.append((row, rec, row["text"]))
    for row, rec, message in notices:
        quote = row.get("quote", row.get("before", ""))
        span, reason = locate(rec.text, quote)
        if span and safety(rec, *span):
            span = None
        if span:
            begin, end = span
            delta = 0
            for start, stop, size in applied.get(rec.id, []):
                if start < end and begin < stop:
                    span = None
                    break
                if stop <= begin:
                    delta += size - (stop - start)
            if span:
                span = (begin + delta, end + delta)
        number = next(ids)
        anchor_p, kind = add_anchor(rec, span, number, document.records)
        if not anchor_p:
            row["comment_status"] = "skipped"
            row["comment_reason"] = "找不到安全批注锚点"
            if "quote" in row:
                row.update(status="skipped", reason=row["comment_reason"])
            continue
        if kind == "nearby":
            message = f"对应原稿 {rec.id}。\n" + message
        if "quote" in row and quote and kind != "quote":
            message = f"原稿引文：{quote}\n" + message
        store.add(number, author, now, message)
        row.update(comment_id=number, anchor_p=anchor_p, anchor=kind)
        if "quote" in row:
            row["status"] = "added"
            if reason and quote:
                row["reason"] = reason
        dirty = True
    updates = store.updates()
    if dirty:
        updates[DOC] = xml_bytes(document.doc)
    document.write(output, updates)
    return report


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="action", required=True)
    inspect = sub.add_parser("inspect")
    inspect.add_argument("source", type=Path)
    inspect.add_argument("--out", type=Path)
    apply = sub.add_parser("apply")
    apply.add_argument("source", type=Path)
    apply.add_argument("plan", type=Path)
    apply.add_argument("--out", required=True, type=Path)
    apply.add_argument("--author", default=None)
    args = parser.parse_args()
    try:
        document = Document(args.source)
        if args.action == "inspect":
            report = {"source_sha256": document.sha, "paragraphs": [p.public() for p in document.records]}
            if args.out:
                atomic_create(args.out, lambda f: f.write_text(json.dumps(report, ensure_ascii=False, indent=2),
                                                               encoding="utf-8"))
                report = {"source_sha256": document.sha, "output": str(args.out.resolve()),
                          "paragraph_count": len(document.records)}
        else:
            with args.plan.open(encoding="utf-8-sig") as f:
                plan = json.load(f)
            author = args.author if args.author is not None else "标书语言审校-" + uuid.uuid4().hex[:8]
            report = apply_plan(document, plan, args.out, author)
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0
    except (OSError, ValueError, KeyError, zipfile.BadZipFile, ET.LxmlError) as exc:
        print(json.dumps({"error": str(exc)}, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

"""Check a tracked DOCX against its source, approved text edits and batch receipt."""
import argparse
import hashlib
import json
import posixpath
from collections import defaultdict
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as ET


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
DOC = "word/document.xml"
RELS = "word/_rels/document.xml.rels"
TYPES = "[Content_Types].xml"
COMMENTS = "word/comments.xml"


def require(condition, message):
    if not condition:
        raise ValueError(message)


def read_json(path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def xml(data):
    return ET.fromstring(data, ET.XMLParser(resolve_entities=False, no_network=True))


def props(el):
    """Compare expanded XML names, including run attributes and nested properties."""
    return (str(el.tag), tuple(sorted(el.attrib.items())), el.text if not len(el) else None,
            tuple(props(child) for child in el))


def simple_run(el):
    return el.tag == W + "r" and all(k.tag in (W + "t", W + "rPr") for k in el)


def run_key(el):
    pr = el.find(W + "rPr")
    return (tuple(sorted(el.attrib.items())), props(pr) if pr is not None else None)


def signature(el):
    """Ignore only segmentation of identical-style plain text runs."""
    children = []
    for child in el:
        if simple_run(child):
            token = ("text-run", run_key(child))
            text = "".join(t.text or "" for t in child.findall(W + "t"))
            if children and children[-1][0] == token:
                children[-1] = (token, children[-1][1] + text)
            else:
                children.append((token, text))
        else:
            children.append(signature(child))
    attrs = tuple(sorted(el.attrib.items()))
    return (str(el.tag), attrs, el.text if not len(el) else None, tuple(children))


def project(tree, author, ids, accept):
    for node in list(tree.iter(W + "ins", W + "del")):
        if node.get(W + "author") != author:
            continue
        require(node.get(W + "id") in ids, "Receipt omits a batch revision")
        parent = node.getparent()
        if (node.tag == W + "ins") == accept:
            pos = parent.index(node)
            for child in list(node):
                for t in child.iter(W + "delText"):
                    t.tag = W + "t"
                parent.insert(pos, child)
                pos += 1
        parent.remove(node)


def remove_batch_anchors(tree, ids):
    for el in list(tree.iter(W + "commentReference", W + "commentRangeStart", W + "commentRangeEnd")):
        if el.get(W + "id") not in ids:
            continue
        parent = el.getparent()
        parent.remove(el)
        if parent.tag == W + "r" and all(c.tag == W + "rPr" for c in parent):
            parent.getparent().remove(parent)


def para_text(p, subtree=None):
    # Exclude nested stories (e.g. the text of a drawing) and pre-existing deletions.
    parts = []
    special = {W + "tab": "\t", W + "br": "\n", W + "cr": "\n",
               W + "noBreakHyphen": "\u2011", W + "softHyphen": "\u00ad"}
    for el in (p if subtree is None else subtree).iter(W + "t", *special):
        ancestors = list(el.iterancestors())
        if next(a for a in ancestors if a.tag == W + "p") is not p:
            continue
        if any(a.tag in (W + "del", W + "moveFrom") for a in ancestors):
            continue
        parts.append((el.text or "") if el.tag == W + "t" else special[el.tag])
    return "".join(parts)


def batch_positions(tracked, revision_ids, comment_ids):
    """Read both original-view and final-view offsets at each new Word marker."""
    revisions, anchors, lengths = {}, defaultdict(dict), {}
    for i, p in enumerate(tracked.find(W + "body").iter(W + "p"), 1):
        pid = f"P{i:04d}"
        original = accepted = 0
        for order, child in enumerate(p):
            rid = child.get(W + "id")
            if child.tag in (W + "del", W + "ins") and rid in revision_ids:
                deletion = child.tag == W + "del"
                tag = W + ("delText" if deletion else "t")
                value = "".join(t.text or "" for t in child.iter(tag))
                require(value and all(r.tag == W + "r" and all(k.tag in (W + "rPr", tag) for k in r)
                                      for r in child), "Unexpected content in text revision " + rid)
                stop = original + (len(value) if deletion else 0)
                revisions[rid] = (pid, order, original, stop, "" if deletion else value,
                                  tuple(run_key(r) for r in child))
                original = stop
                accepted += 0 if deletion else len(value)
                continue
            marker = child.find(W + "commentReference") if child.tag == W + "r" else child
            if marker is not None and marker.tag in (
                    W + "commentRangeStart", W + "commentRangeEnd", W + "commentReference"):
                cid = marker.get(W + "id")
                if cid in comment_ids:
                    require(marker.tag not in anchors[cid], "Duplicate Word comment anchor: " + cid)
                    if marker.tag == W + "commentReference":
                        require(all(n is marker or n.tag == W + "rPr" for n in child),
                                "Word comment reference shares a run with unrelated content: " + cid)
                    anchors[cid][marker.tag] = (pid, order, original, accepted)
            size = len(para_text(p, child))
            original += size
            accepted += size
        lengths[pid] = (original, accepted)
    require(set(revisions) == revision_ids, "Batch revision is outside a supported text paragraph")
    return revisions, anchors, lengths


def check_edit(edit, revisions, text):
    pid = edit["p"]
    before = edit["before"]
    start = text.find(before)
    require(before and start >= 0 and text.find(before, start + 1) < 0, "Ambiguous source edit at " + pid)
    stop = start + len(before)
    cursor, last_order = start, -1
    result, deleted, inserted = [], [], []
    for rid in edit["revision_ids"]:
        actual_p, order, left, right, replacement, styles = revisions[str(rid)]
        require(actual_p == pid and order > last_order and cursor <= left <= right <= stop,
                "Revision belongs to a different proposal or source range at " + pid)
        result.extend((text[cursor:left], replacement))
        (inserted if replacement else deleted).extend(styles)
        cursor, last_order = right, order
    result.append(text[cursor:stop])
    require("".join(result) == edit["after"], "Independent acceptance does not match the proposal at " + pid)
    require(deleted and all(style == deleted[0] for style in inserted),
            "Replacement run lost original character formatting at " + pid)


def check_notes(rows, bodies, author, anchors, lengths, source, final):
    for row in rows:
        cid, pid, target = str(row["comment_id"]), row["p"], row["anchor_p"]
        kind = row.get("anchor")
        quote = row["quote"] if "quote" in row else row["before"]
        marks = anchors[cid]
        tags = [W + tag for tag in ("commentRangeStart", "commentRangeEnd", "commentReference")]
        require(all(tag in marks for tag in tags), "Missing or misplaced Word comment anchor: " + cid)
        begin, end, ref = [marks[tag] for tag in tags]
        require(all(point[0] == target for point in (begin, end, ref)) and
                begin[1] < end[1] < ref[1] and end[2:] == ref[2:] and pid in source,
                "Word comment anchor paragraph, order or reference differs from receipt: " + cid)
        require(bodies[cid].get(W + "author") == author, "Word comment author differs from batch: " + cid)
        body_text = "\n".join("".join(t.text or "" for t in p.iter(W + "t")) for p in bodies[cid].iter(W + "p"))
        if "quote" in row:
            note = row["text"].replace("\r\n", "\n").replace("\r", "\n")
            require(note in body_text, "Word comment does not contain the planned note: " + cid)
        else:
            require(quote in body_text and row["after"] in body_text, "Missing skipped proposal in Word comment: " + cid)
        if kind == "quote":
            original = source[pid]
            require(pid == target and quote and original[begin[2]:end[2]] == quote and
                    original.find(quote) == begin[2] and original.find(quote, begin[2] + 1) < 0 and
                    final[pid][begin[3]:end[3]] == quote, "Word comment does not enclose its original quotation: " + cid)
        else:
            require(kind == ("paragraph" if pid == target else "nearby") and begin[2:] == end[2:] and
                    begin[2:] in ((0, 0), lengths[target]), "Invalid fallback Word comment location: " + cid)
            require(quote in body_text and (kind != "nearby" or pid in body_text),
                    "Fallback Word comment does not identify the source text: " + cid)


def expected_text(text, edits):
    spans = []
    for change in edits:
        before = change["before"]
        require(before and text.count(before) == 1, "Acceptance source phrase is ambiguous")
        start = text.index(before)
        spans.append((start, start + len(before), change["after"]))
    end = len(text)
    for start, stop, after in sorted(spans, reverse=True):
        require(stop <= end, "Overlapping accepted edits")
        text = text[:start] + after + text[stop:]
        end = start
    return text


def comment_path(package):
    if RELS in package.namelist():
        for rel in xml(package.read(RELS)):
            if rel.get("Type", "").endswith("/comments") and rel.get("TargetMode") != "External":
                target = rel.get("Target")
                return posixpath.normpath(target.lstrip("/") if target.startswith("/") else "word/" + target)
    return COMMENTS


def check_comment_parts(source, result, part_name, comment_ids):
    for part in (RELS, TYPES):
        if part not in result.namelist():
            require(part not in source.namelist() and not comment_ids, "Missing " + part)
            continue
        new = xml(result.read(part))
        old = xml(source.read(part)) if part in source.namelist() else ET.Element(new.tag)
        old_rows = {props(o) for o in old}
        extras = [c for c in new if props(c) not in old_rows]
        for el in extras:
            valid = (el.get("Type", "").endswith("/comments") and el.get("Target") in ("comments.xml", "/word/comments.xml")
                     if part == RELS else el.get("PartName") == "/" + part_name)
            require(comment_ids and valid, "Unexpected change in " + part)
            new.remove(el)
        require(props(old) == props(new), "Existing entries changed in " + part)
    if part_name not in result.namelist():
        require(not comment_ids, "Missing Word comments part")
        return {}
    new = xml(result.read(part_name))
    added = {}
    for comment in list(new):
        cid = comment.get(W + "id")
        if cid in comment_ids:
            require(cid not in added, "Duplicate Word comment id")
            added[cid] = comment
            new.remove(comment)
    require(set(added) == comment_ids, "Incomplete Word comment bodies")
    if part_name in source.namelist():
        require(props(xml(source.read(part_name))) == props(new), "An existing comment changed")
    else:
        require(len(new) == 0, "Unexpected new Word comments")
    return added


def verify(source_path, plan_path, tracked_path, receipt_path, accepted_path=None):
    plan, receipt = read_json(plan_path), read_json(receipt_path)
    digest = hashlib.sha256(source_path.read_bytes()).hexdigest()
    require(plan["source_sha256"] == receipt["source_sha256"] == digest, "Source fingerprint changed")
    author = receipt["author"]
    edits, comments = receipt["edits"], receipt["comments"]
    for planned, logged, fields in ((plan.get("edits", []), edits, ("p", "before", "after")),
                                    (plan.get("comments", []), comments, ("p", "quote", "text"))):
        require([[e[k] for k in fields] for e in planned] == [[e[k] for k in fields] for e in logged],
                "Receipt differs from requested operations")
    changes = defaultdict(list)
    revisions = set()
    for edit in edits:
        require(edit["status"] in ("applied", "skipped", "unchanged"), "Unknown edit outcome in receipt")
        if edit["status"] == "applied":
            changes[edit["p"]].append(edit)
            claimed = [str(v) for v in edit["revision_ids"]]
            require(claimed and len(claimed) == len(set(claimed)) and revisions.isdisjoint(claimed),
                    "Revision IDs do not belong to exactly one proposal")
            revisions.update(claimed)
        else:
            require(not edit["revision_ids"], "Unapplied edit claims Word revisions")
    for note in comments:
        require(note["status"] in ("added", "skipped") and
                (note["status"] == "added") == (note.get("comment_id") is not None),
                "Word comment outcome and ID disagree")
    notes = [row for row in edits + comments if row.get("comment_id") is not None]
    cids = {str(e["comment_id"]) for e in notes}
    require(len(cids) == len(notes), "Duplicate comment ID in receipt")
    with ZipFile(source_path) as src, ZipFile(tracked_path) as out:
        names = set(src.namelist())
        part_name = comment_path(out)
        require(set(out.namelist()) == (names | ({part_name, RELS} if cids else set())), "Package inventory changed")
        for part in names - {DOC, RELS, TYPES, part_name}:
            require(src.read(part) == out.read(part), "Unexpected part modification: " + part)
        bodies = check_comment_parts(src, out, part_name, cids)
        orig, tracked = xml(src.read(DOC)), xml(out.read(DOC))
        for name in ("commentReference", "commentRangeStart", "commentRangeEnd"):
            found = [el.get(W + "id") for el in tracked.iter(W + name) if el.get(W + "id") in cids]
            require(len(found) == len(cids) and set(found) == cids, "Invalid Word comment anchors: " + name)
        require(all(e.get(W + "author") != author for e in orig.iter(W + "ins", W + "del")),
                "Batch author collides with original revisions")
        actual = [e.get(W + "id") for e in tracked.iter(W + "ins", W + "del")
                  if e.get(W + "author") == author]
        require(len(actual) == len(set(actual)) and set(actual) == revisions, "Revision ID mismatch")
        rejected = deepcopy(tracked)
        project(rejected, author, revisions, False)
        remove_batch_anchors(rejected, cids)
        require(signature(orig) == signature(rejected), "Rejecting the batch changed source content, formatting or objects")
        positions, anchors, lengths = batch_positions(tracked, revisions, cids)
        project(tracked, author, revisions, True)
        before = list(orig.find(W + "body").iter(W + "p"))
        after = list(tracked.find(W + "body").iter(W + "p"))
        require(len(before) == len(after), "Paragraph count changed")
        source_text, final_text = {}, {}
        for idx, (old, new) in enumerate(zip(before, after), 1):
            pid = f"P{idx:04d}"
            source_text[pid], final_text[pid] = para_text(old), para_text(new)
            for edit in changes.get(pid, []):
                check_edit(edit, positions, source_text[pid])
            wanted = expected_text(source_text[pid], changes.get(pid, []))
            require(final_text[pid] == wanted, "Accepted text differs from the plan at " + pid)
        require(not (set(changes) - source_text.keys()), "Invalid receipt paragraph ID")
        check_notes(notes, bodies, author, anchors, lengths, source_text, final_text)
        if accepted_path:
            data = ET.tostring(tracked, encoding="UTF-8", xml_declaration=True, standalone=True)
            with ZipFile(accepted_path, "x") as target:
                for part in out.infolist():
                    target.writestr(part, data if part.filename == DOC else out.read(part.filename))
    return {"verified": True, "applied_edits": sum(map(len, changes.values())),
            "revision_nodes": len(revisions), "new_comments": len(cids), "checked_paragraphs": len(before)}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    for arg in ("source", "plan", "tracked", "receipt"):
        parser.add_argument(arg, type=Path)
    parser.add_argument("--accepted", type=Path, help="New temporary DOCX for Word text verification")
    args = parser.parse_args()
    print(json.dumps(verify(args.source, args.plan, args.tracked, args.receipt, args.accepted), ensure_ascii=False))

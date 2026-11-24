"""Run with the same Python used for the skill (python-docx 1.2+, lxml)."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree


SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
MAIN = "word/document.xml"
AUTHOR = "本轮审校"
PREVIOUS = "原审阅人"


def cli(script, *args, check=True):
    return subprocess.run([sys.executable, str(SCRIPTS / script), *map(str, args)],
                          capture_output=True, encoding="utf-8", errors="replace", check=check)


def fixture(path):
    doc = Document()
    doc.add_paragraph("接到通知后，我方于30分钟内电话响应。")
    p = doc.add_paragraph("现场检查")
    p.add_run("X-42").bold = True
    p.add_run("控制器。")
    p = doc.add_paragraph()
    doc.add_comment(p.add_run("既有记录"), text="由采购人确认。", author=PREVIOUS)
    insertion = OxmlElement("w:ins")
    insertion.set(qn("w:id"), "17")
    insertion.set(qn("w:author"), PREVIOUS)
    run = p.add_run("原修订")
    insertion.append(run._r)
    p._p.append(insertion)
    doc.save(path)


def part_xml(path, name=MAIN):
    with ZipFile(path) as package:
        return etree.fromstring(package.read(name))


class RoundTrip(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        temp = TemporaryDirectory()
        cls.addClassCleanup(temp.cleanup)
        cls.root = Path(temp.name)
        cls.source = cls.root / "source.docx"
        cls.plan = cls.root / "plan.json"
        cls.tracked = cls.root / "review.docx"
        cls.receipt = cls.root / "receipt.json"
        fixture(cls.source)
        info = json.loads(cli("bid_docx.py", "inspect", cls.source).stdout)
        cls.digest = info["source_sha256"]
        p, mixed, _ = [row["id"] for row in info["paragraphs"]]
        plan = {"source_sha256": cls.digest,
                "edits": [{"p": p, "before": "我方于", "after": "我方在"},
                          {"p": mixed, "before": "现场检查X-42", "after": "检查设备"},
                          {"p": p, "before": "接到通知后", "after": "收到通知后"}],
                "comments": [{"p": p, "quote": "30分钟", "text": "请核查时限。"}]}
        cls.plan.write_text(json.dumps(plan, ensure_ascii=False), encoding="utf-8")
        out = cli("bid_docx.py", "apply", cls.source, cls.plan, "--out", cls.tracked, "--author", AUTHOR)
        cls.report = json.loads(out.stdout)
        cls.receipt.write_text(out.stdout, encoding="utf-8")

    def verify(self, tracked, *opts, check=True):
        return cli("verify_docx.py", self.source, self.plan, tracked, self.receipt, *opts, check=check)

    def tamper(self, filename, mutate):
        dest = self.root / filename
        with ZipFile(self.tracked) as old, ZipFile(dest, "x") as new:
            for item in old.infolist():
                data = old.read(item)
                if item.filename == MAIN:
                    root = etree.fromstring(data)
                    mutate(root)
                    data = etree.tostring(root, encoding="utf-8", xml_declaration=True)
                new.writestr(item, data)
        return dest

    def test_roundtrip_and_existing_review(self):
        accepted = self.root / "accepted.docx"
        result = json.loads(self.verify(self.tracked, "--accepted", accepted).stdout)
        self.assertTrue(result["verified"])
        self.assertEqual((result["applied_edits"], result["new_comments"]), (2, 2))
        self.assertEqual(Document(accepted).paragraphs[0].text, "收到通知后，我方在30分钟内电话响应。")
        self.assertTrue(any(n.get(qn("w:author")) == PREVIOUS and n.get(qn("w:id")) == "17"
                            for n in part_xml(accepted).iter(qn("w:ins"))))
        original_note = next(c for c in Document(accepted).comments if c.author == PREVIOUS)
        self.assertEqual(original_note.text, "由采购人确认。")

    def test_mixed_format_degrades_to_comment(self):
        protected = self.report["edits"][1]
        self.assertEqual(protected["status"], "skipped")
        doc = Document(self.tracked)
        notes = {c.comment_id: c.text for c in doc.comments}
        self.assertIn("现场检查X-42", notes[protected["comment_id"]])
        self.assertEqual(doc.paragraphs[1].text, "现场检查X-42控制器。")
        self.assertTrue(next(r for r in doc.paragraphs[1].runs if r.text == "X-42").bold)

    def test_verifier_detects_broken_rejection(self):
        def alter_original(root):
            deleted = next(e for e in root.iter(qn("w:del")) if e.get(qn("w:author")) == AUTHOR)
            next(deleted.iter(qn("w:delText"))).text = "不可恢复"

        out = self.verify(self.tamper("bad-original.docx", alter_original), check=False)
        self.assertNotEqual(out.returncode, 0)
        self.assertIn("Rejecting the batch", out.stderr + out.stdout)

    def test_verifier_detects_inserted_style_change(self):
        def add_bold(root):
            inserted = next(e for e in root.iter(qn("w:ins")) if e.get(qn("w:author")) == AUTHOR)
            run = next(inserted.iter(qn("w:r")))
            props = run.find(qn("w:rPr"))
            if props is None:
                props = OxmlElement("w:rPr")
                run.insert(0, props)
            props.append(OxmlElement("w:b"))

        out = self.verify(self.tamper("bad-style.docx", add_bold), check=False)
        self.assertNotEqual(out.returncode, 0)
        self.assertIn("formatting", out.stderr + out.stdout)

    def test_verifier_detects_reassigned_revision_ids(self):
        receipt = json.loads(self.receipt.read_text(encoding="utf-8"))
        first, second = [edit for edit in receipt["edits"] if edit["status"] == "applied"]
        self.assertEqual(first["p"], second["p"])
        first["revision_ids"], second["revision_ids"] = second["revision_ids"], first["revision_ids"]
        switched = self.root / "switched-ids.json"
        switched.write_text(json.dumps(receipt), encoding="utf-8")
        out = cli("verify_docx.py", self.source, self.plan, self.tracked, switched, check=False)
        self.assertNotEqual(out.returncode, 0)

    def test_verifier_detects_relocated_comment(self):
        cid = str(self.report["comments"][0]["comment_id"])

        def relocate(root):
            paragraphs = list(root.iter(qn("w:p")))
            target = paragraphs[2]
            def anchor(tag):
                return next(n for n in paragraphs[0].iter(qn("w:" + tag)) if n.get(qn("w:id")) == cid)
            target.insert(0, anchor("commentRangeStart"))
            target.append(anchor("commentRangeEnd"))
            target.append(anchor("commentReference").getparent())

        out = self.verify(self.tamper("wrong-anchor.docx", relocate), check=False)
        self.assertNotEqual(out.returncode, 0)

    def test_prevents_overwriting_existing_files(self):
        existing = self.root / "preexisting.docx"
        existing.write_bytes(b"pre-existing file")
        for dest in (existing, self.source):
            with self.subTest(destination=dest.name):
                before = dest.read_bytes()
                out = cli("bid_docx.py", "apply", self.source, self.plan, "--out", dest, check=False)
                self.assertNotEqual(out.returncode, 0)
                self.assertEqual(dest.read_bytes(), before)
        self.assertEqual(hashlib.sha256(self.source.read_bytes()).hexdigest(), self.digest)

    def test_service_level_changes_become_word_comments(self):
        phrase = "30分钟内电话响应"
        for idx, substitute in enumerate(("15分钟内电话响应", "30分钟内到场", "30工作日内电话响应")):
            with self.subTest(substitute=substitute):
                plan = self.root / f"service-{idx}.json"
                output = self.root / f"service-{idx}.docx"
                request = {"source_sha256": self.digest, "comments": [],
                           "edits": [{"p": "P0001", "before": phrase, "after": substitute}]}
                plan.write_text(json.dumps(request), encoding="utf-8")
                report = json.loads(cli("bid_docx.py", "apply", self.source, plan, "--out", output).stdout)
                edit = report["edits"][0]
                self.assertEqual(edit["status"], "skipped")
                self.assertEqual(edit["anchor"], "quote")
                self.assertEqual(edit["revision_ids"], [])
                doc = Document(output)
                self.assertEqual(doc.paragraphs[0].text, "接到通知后，我方于30分钟内电话响应。")
                self.assertIn(substitute, next(c.text for c in doc.comments if c.comment_id == edit["comment_id"]))


class CommitmentReview(unittest.TestCase):
    # One proposed rewrite per paragraph, with no structural hazards that could mask a semantic miss.
    EXAMPLES = (
        ("deadline", "我方在2小时内提交方案。采购人确认后，我方实施。",
         "我方提交方案。采购人确认后，我方在2小时内实施。", "skipped"),
        ("roles", "变更前，采购人审核方案，投标人执行。",
         "变更前，投标人审核方案，采购人执行。", "skipped"),
        ("postposed_deadline", "我方电话响应不超过30分钟。",
         "我方到场不超过30分钟。", "skipped"),
        ("both_conditions", "仅当账本平衡且抽样复核通过时，允许封存本批数据。",
         "仅当账本平衡或抽样复核通过时，允许封存本批数据。", "skipped"),
        ("prohibition", "采购人确认前，投标人不得执行变更。",
         "采购人确认前，投标人可以执行变更。", "skipped"),
        ("condition", "未经采购人书面确认，投标人不得执行变更。",
         "经采购人书面确认，投标人不得执行变更。", "skipped"),
        ("preposition", "系统故障时，工程师应于3小时内提交报告。",
         "系统故障时，工程师应在3小时内提交报告。", "applied"),
        ("active_voice", "演练前，由采购人审核清单，再由投标人执行。",
         "演练前，采购人审核清单，再由投标人执行。", "applied"),
        ("approval_then_execution", "我方将在采购人审批后开展实施工作。",
         "采购人审批后，我方实施。", "applied"),
    )

    @classmethod
    def setUpClass(cls):
        temp = TemporaryDirectory()
        cls.addClassCleanup(temp.cleanup)
        root = Path(temp.name)
        source, plan, reviewed, receipt, accepted = (root / f for f in (
            "source.docx", "plan.json", "review.docx", "receipt.json", "accepted.docx"))
        fixture(source)
        doc = Document(source)
        for _, before, _, _ in cls.EXAMPLES:
            doc.add_paragraph().add_run(before).italic = True
        doc.save(source)
        index = json.loads(cli("bid_docx.py", "inspect", source).stdout)
        ids = {r["text"]: r["id"] for r in index["paragraphs"]}
        edits = [{"p": ids[before], "before": before, "after": after} for _, before, after, _ in cls.EXAMPLES]
        plan.write_text(json.dumps({"source_sha256": index["source_sha256"], "edits": edits}, ensure_ascii=False),
                        encoding="utf-8")
        result = cli("bid_docx.py", "apply", source, plan, "--out", reviewed, "--author", AUTHOR)
        receipt.write_text(result.stdout, encoding="utf-8")
        cls.edits = json.loads(result.stdout)["edits"]
        cls.summary = json.loads(cli("verify_docx.py", source, plan, reviewed, receipt, "--accepted", accepted).stdout)
        cls.review = Document(reviewed)
        cls.accepted = Document(accepted)
        cls.accepted_xml = part_xml(accepted)
        cls.first = len(cls.accepted.paragraphs) - len(cls.EXAMPLES)

    def test_changes_of_meaning_annotate_the_original_sentence(self):
        for idx, (name, before, after, status) in enumerate(self.EXAMPLES):
            if status != "skipped":
                continue
            with self.subTest(example=name):
                result = self.edits[idx]
                self.assertEqual(result["status"], "skipped")
                self.assertEqual(result["revision_ids"], [])
                self.assertEqual(result["anchor"], "quote")
                cid = result["comment_id"]
                body = next(c.text for c in self.review.comments if c.comment_id == cid)
                self.assertIn(before, body)
                self.assertIn(after, body)
                paragraph = self.review.paragraphs[self.first + idx]
                in_anchor, quoted = False, []
                for node in paragraph._p.iter():
                    if node.get(qn("w:id")) == str(cid) and node.tag in (qn("w:commentRangeStart"), qn("w:commentRangeEnd")):
                        in_anchor = node.tag == qn("w:commentRangeStart")
                    elif in_anchor and node.tag == qn("w:t"):
                        quoted.append(node.text or "")
                self.assertEqual("".join(quoted), before)
                self.assertEqual(self.accepted.paragraphs[self.first + idx].text, before)

    def test_equivalent_proposals_are_tracked_and_keep_format(self):
        for idx, (name, _, after, status) in enumerate(self.EXAMPLES):
            if status != "applied":
                continue
            with self.subTest(example=name):
                self.assertEqual(self.edits[idx]["status"], "applied")
                self.assertTrue(self.edits[idx]["revision_ids"])
                paragraph = self.accepted.paragraphs[self.first + idx]
                self.assertEqual(paragraph.text, after)
                self.assertTrue(all(r.italic for r in paragraph.runs if r.text))

    def test_semantic_review_preserves_existing_records_and_format(self):
        self.assertTrue(self.summary["verified"])
        self.assertTrue(any(node.get(qn("w:id")) == "17" and node.get(qn("w:author")) == PREVIOUS
                            for node in self.accepted_xml.iter(qn("w:ins"))))
        self.assertEqual(next(c.text for c in self.review.comments if c.author == PREVIOUS), "由采购人确认。")
        self.assertTrue(next(r for r in self.review.paragraphs[1].runs if r.text == "X-42").bold)


if __name__ == "__main__":
    unittest.main()

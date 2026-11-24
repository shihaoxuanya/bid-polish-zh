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
        ("phone_word_order", "值班工程师应在30分钟内以电话方式向采购人作出响应。",
         "值班工程师应在30分钟内向采购人作出电话响应。", "applied"),
        ("phone_recipient", "接报后，我方在30分钟内以电话方式向采购人作出响应。",
         "接报后，我方在30分钟内向供应商作出电话响应。", "skipped"),
        ("phone_mode", "试运行期间，我方应在30分钟内以电话方式向采购人作出响应。",
         "试运行期间，我方应在30分钟内向采购人作出响应。", "skipped"),
        ("phone_action", "发生故障时，工程师应在30分钟内以电话方式向采购人作出响应。",
         "发生故障时，工程师应在30分钟内到场修复。", "skipped"),
        ("both_conditions", "仅当账本平衡且抽样复核通过时，允许封存本批数据。",
         "仅当账本平衡或抽样复核通过时，允许封存本批数据。", "skipped"),
        ("prohibition", "采购人确认前，投标人不得执行变更。",
         "采购人确认前，投标人可以执行变更。", "skipped"),
        ("condition", "未经采购人书面确认，投标人不得执行变更。",
         "经采购人书面确认，投标人不得执行变更。", "skipped"),
        ("written_approval", "经采购人书面批准后，实施单位开始切换。",
         "经采购人批准后，实施单位开始切换。", "skipped"),
        ("informal_assent", "经采购人书面批准后，我方启动迁移。",
         "经采购人同意后，我方启动迁移。", "skipped"),
        ("written_approval_owner", "经采购人书面批准后，我方开始切换。",
         "经供应商书面批准后，我方开始切换。", "skipped"),
        ("keep_written_approval", "经采购人书面批准后，实施单位开展切换工作。",
         "经采购人书面批准后，实施单位开始切换。", "applied"),
        ("optional_delivery", "值班工程师必须提交故障报告。",
         "值班工程师可提交故障报告。", "skipped"),
        ("different_clock_start", "接到采购人报障后，值班工程师应在30分钟内电话响应。",
         "故障发生后，值班工程师应在30分钟内电话响应。", "skipped"),
        ("different_notice_recipient", "故障通知送达服务台后，值班工程师须在30分钟内电话响应。",
         "故障通知送达采购人后，值班工程师须在30分钟内电话响应。", "skipped"),
        ("equivalent_clock_start", "接到采购人报障后，值班工程师须在30分钟内电话响应。",
         "收到采购人报障后，值班工程师须在30分钟内电话响应。", "applied"),
        ("best_effort_deadline", "我方力争在2小时内恢复服务。",
         "我方必须在2小时内恢复服务。", "skipped"),
        ("response_priority", "我方应优先恢复核心业务。",
         "我方应恢复核心业务。", "skipped"),
        ("qualified_training_priority", "我方原则上优先安排同一名讲师授课。",
         "我方安排同一名讲师授课。", "skipped"),
        ("drop_default", "培训期间，我方原则上优先安排同一名讲师授课。",
         "培训期间，我方优先安排同一名讲师授课。", "skipped"),
        ("drop_priority", "连续培训时，我方原则上优先安排同一名讲师授课。",
         "连续培训时，我方原则上安排同一名讲师授课。", "skipped"),
        ("keep_default_and_priority", "我方原则上应当优先安排同一名讲师进行授课。",
         "原则上，我方应优先安排同一名讲师授课。", "applied"),
        ("quoted_default", "培训材料列举了“原则上”一词的使用示例。",
         "培训材料举例说明“原则上”的用法。", "applied"),
        ("permitted_delivery", "采购人可以提交补充材料。",
         "采购人可提交补充材料。", "applied"),
        ("required_delivery", "我方应当及时提交工单。",
         "我方应及时提交工单。", "applied"),
        ("ordinary_word_parts", "从巡检记录可以看出，应用服务响应时间波动的可能性较低。",
         "巡检记录显示，应用服务响应时间波动的可能性较低。", "applied"),
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


class ProjectTerms(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        temp = TemporaryDirectory()
        cls.addClassCleanup(temp.cleanup)
        cls.root = Path(temp.name)
        cls.source = cls.root / "source.docx"
        doc = Document()
        doc.add_paragraph("为提升服务水平，值班工程师提供电话响应，并将本次电话响应的时间填写至故障工单中。")
        doc.add_paragraph("我方对配置基线开展核对工作。")
        doc.add_paragraph("培训示例中的电话响应使用测试号码。")
        doc.add_paragraph("迁移前须核对配置基线。验收时核查记录。")
        doc.save(cls.source)
        info = json.loads(cli("bid_docx.py", "inspect", cls.source).stdout)
        cls.ids = [r["id"] for r in info["paragraphs"]]
        cls.digest = info["source_sha256"]

    def test_scoped_terms_preserve_wording_and_allow_surrounding_edits(self):
        first, second, third, fourth = self.ids
        scoring = "评分表第3项"
        appendix = "采购需求附件A第2段"
        plan = {
            "source_sha256": self.digest,
            "protected_terms": [
                {"p": first, "term": "电话响应", "source": scoring},
                {"p": second, "term": "配置基线", "source": appendix},
                {"p": fourth, "term": "配置基线", "source": appendix},
            ],
            "edits": [
                {"p": first, "before": "为提升服务水平，", "after": ""},
                {"p": first, "before": "响应，并", "after": "回复，并"},
                {"p": second, "before": "对配置基线开展核对工作", "after": "核对配置基线"},
                {"p": third, "before": "电话响应", "after": "电话回复"},
                {"p": fourth, "before": "配置基线", "after": "当前配置"},
                {"p": first, "before": "将本次电话响应的时间填写至故障工单中", "after": "在故障工单中记录电话响应时间"},
            ],
        }
        plan_path = self.root / "plan.json"
        tracked = self.root / "tracked.docx"
        receipt = self.root / "receipt.json"
        accepted = self.root / "accepted.docx"
        plan_path.write_text(json.dumps(plan, ensure_ascii=False), encoding="utf-8")
        out = cli("bid_docx.py", "apply", self.source, plan_path, "--out", tracked, "--author", AUTHOR)
        receipt.write_text(out.stdout, encoding="utf-8")
        edits = json.loads(out.stdout)["edits"]
        self.assertEqual([e["status"] for e in edits], ["applied", "skipped", "applied", "applied", "skipped", "applied"])
        notes = {c.comment_id: c.text for c in Document(tracked).comments}
        for edit, term, source in ((edits[1], "电话响应", scoring), (edits[4], "配置基线", appendix)):
            with self.subTest(term=term):
                self.assertEqual(edit["anchor"], "quote")
                self.assertEqual(edit["revision_ids"], [])
                self.assertIn(term, notes[edit["comment_id"]])
                self.assertIn(source, notes[edit["comment_id"]])
        result = cli("verify_docx.py", self.source, plan_path, tracked, receipt, "--accepted", accepted)
        self.assertTrue(json.loads(result.stdout)["verified"])
        self.assertEqual([p.text for p in Document(accepted).paragraphs], [
            "值班工程师提供电话响应，并在故障工单中记录电话响应时间。",
            "我方核对配置基线。",
            "培训示例中的电话回复使用测试号码。",
            "迁移前须核对配置基线。验收时核查记录。",
        ])

    def test_moving_a_protected_term_to_another_sentence_is_flagged(self):
        original = "迁移前须核对配置基线。验收时核查记录。"
        plan = {"source_sha256": self.digest, "protected_terms": [
            {"p": self.ids[3], "term": "配置基线", "source": "迁移评分项第2项"}], "edits": [
            {"p": self.ids[3], "before": original, "after": "迁移前须核对配置。验收时核查配置基线。"}]}
        plan_path, tracked = self.root / "move-term.json", self.root / "move-term.docx"
        plan_path.write_text(json.dumps(plan), encoding="utf-8")
        report = cli("bid_docx.py", "apply", self.source, plan_path, "--out", tracked)
        edit = json.loads(report.stdout)["edits"][0]
        self.assertEqual(edit["status"], "skipped")
        self.assertEqual(edit["anchor"], "quote")
        self.assertIn("配置基线", edit["reason"])
        self.assertEqual(Document(tracked).paragraphs[3].text, original)

    def test_invalid_term_scope_fails_before_writing(self):
        plan_path = self.root / "invalid.json"
        output = self.root / "invalid.docx"
        for pid, term in (("P9999", "电话响应"), (self.ids[0], "配置基线")):
            with self.subTest(pid=pid, term=term):
                plan = {"source_sha256": self.digest,
                        "protected_terms": [{"p": pid, "term": term, "source": "评分表第3项"}]}
                plan_path.write_text(json.dumps(plan), encoding="utf-8")
                out = cli("bid_docx.py", "apply", self.source, plan_path, "--out", output, check=False)
                self.assertNotEqual(out.returncode, 0)
                self.assertIn(pid, json.loads(out.stdout)["error"])
                self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()

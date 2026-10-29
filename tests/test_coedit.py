"""共同編集（deck_merge）のテスト。PowerPoint なしで python-pptx だけで「人の編集」を作って 3-way マージを確かめる。
usage: python3 -m unittest discover -s tests
"""
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import deck_merge as dm  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.dml.color import RGBColor  # noqa: E402
from pptx.util import Emu  # noqa: E402
from slide_template import (new_presentation, add_content_slide, add_conclusion, add_source, add_section_slide,  # noqa: E402
                            add_flow, set_notes, save)


def build(msg_a="Message A", msg_b="Message B", note_b="note v1", line_rgb=None, with_c=True):
    prs = new_presentation()
    add_section_slide(prs, "Intro")
    s, _ = add_content_slide(prs, "Slide A", 1)
    add_conclusion(s, msg_a)
    add_source(s, "[1] A. Author, Journal 1 (2020) 1.")
    s, _ = add_content_slide(prs, "Slide B", 2)
    add_conclusion(s, msg_b)
    boxes = add_flow(s, ["Step 1", "Step 2"], 500000, 1500000, 8000000, 1200000, direction="right")
    set_notes(s, note_b)
    if line_rgb is not None:
        sh = next(x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("Step 1"))
        sh.line.color.rgb = line_rgb
    if with_c:
        s, _ = add_content_slide(prs, "Slide C", 3)
        add_conclusion(s, "Message C")
    return prs


def by_key(prs, key):
    return next(s for s in prs.slides if dm.slide_key(s) == key)


def shape(slide, name):
    return next((s for s in slide.shapes if s.name == name), None)


class CoeditTest(unittest.TestCase):
    def setUp(self):
        self.d = Path(tempfile.mkdtemp())
        self.out = self.d / "deck.pptx"

    def tearDown(self):
        shutil.rmtree(self.d)

    def test_merge(self):
        r = save(build(), str(self.out))
        self.assertEqual(r["mode"], "new")
        # 人の編集: A の結論を書き換え・B に枠を追加・A の出典を削除・B の Step 1 の線の色を変える
        th = Presentation(str(self.out))
        shape(by_key(th, "Slide A#1"), "MSG").text_frame.paragraphs[0].runs[0].text = "MY EDIT"
        tb = by_key(th, "Slide B#1").shapes.add_textbox(Emu(300000), Emu(300000), Emu(3000000), Emu(400000))
        tb.text_frame.text = "USER NOTE"
        ref = shape(by_key(th, "Slide A#1"), "REF")
        ref._element.getparent().remove(ref._element)
        th.save(str(self.out))
        # 生成側の変更: A・B の結論・B のノート・Step 1 の線の色・C を削除
        r2 = save(build(msg_a="GEN A v2", msg_b="GEN B v2", note_b="note v2", line_rgb=RGBColor(0, 0x80, 0), with_c=False),
                  str(self.out))
        self.assertEqual(r2["mode"], "merge")
        self.assertTrue(Path(r2["backup"]).exists())
        m = Presentation(str(self.out))
        keys = [dm.slide_key(s) for s in m.slides]
        self.assertEqual(shape(by_key(m, "Slide A#1"), "MSG").text_frame.text, "MY EDIT")          # 衝突: あなたの版
        self.assertTrue(any(l.startswith("[衝突]") and "Slide A#1" in l for l in r2["report"]))
        self.assertEqual(shape(by_key(m, "Slide B#1"), "MSG").text_frame.text, "GEN B v2")         # 未編集: 更新
        self.assertEqual(by_key(m, "Slide B#1").notes_slide.notes_text_frame.text, "note v2")
        self.assertTrue(any(s.has_text_frame and s.text_frame.text == "USER NOTE" for s in by_key(m, "Slide B#1").shapes))
        self.assertIsNone(shape(by_key(m, "Slide A#1"), "REF"))                                  # 消した出典は戻らない
        step1 = next(x for x in by_key(m, "Slide B#1").shapes if x.has_text_frame and x.text_frame.text.startswith("Step 1"))
        self.assertEqual(step1.line.color.rgb, RGBColor(0, 0x80, 0))                             # 書式も更新
        self.assertNotIn("Slide C#1", keys)                                                      # 生成側で消したスライド
        # 同じ生成で再保存 → 衝突を繰り返さない
        r3 = save(build(msg_a="GEN A v2", msg_b="GEN B v2", note_b="note v2", line_rgb=RGBColor(0, 0x80, 0), with_c=False),
                  str(self.out))
        self.assertFalse(any(l.startswith("[衝突]") for l in r3["report"]))
        self.assertEqual(shape(by_key(Presentation(str(self.out)), "Slide A#1"), "MSG").text_frame.text, "MY EDIT")

    def test_citation_footers(self):
        from slide_template import keep_citation_footers
        prs = new_presentation()
        for i, t in enumerate(["Data: my_table.csv, Re = 1e5.", "Cropped from Fig. 3 of [2] B. Author, J. 2 (2021) 5 (CC BY 4.0).",
                               "Re = 1e5. [3] C. Author, J. 3 (2022). This study: run 1.", "Exp.: NACA TN 1234. model xyz"]):
            s, _ = add_content_slide(prs, f"S{i}", i + 1)
            add_source(s, t)
        keep_citation_footers(prs)
        refs = [shape(s, "REF").text_frame.text if shape(s, "REF") else None for s in prs.slides]
        self.assertEqual(refs, [None, "[2] B. Author, J. 2 (2021) 5.", "[3] C. Author, J. 3 (2022).", "Experiment: NACA TN 1234."])


if __name__ == "__main__":
    unittest.main()

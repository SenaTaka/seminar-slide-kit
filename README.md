# seminar-slide-kit

研究室ゼミのスライド（進捗報告・論文紹介）を **Python（python-pptx）で同じデザインに生成する**ためのキット。
人が PowerPoint で直した部分を消さずに再生成できる **共同編集つき保存**（3-way マージ）と、
提出前の自動チェック（`audit`）を備える。

## できること

- **型関数でスライドを作る**（`slide_template.py`）: タイトル・扉・目的・流れ図＋This time・図＋結論・解析条件・表・
  二者比較・数式＋記号表・Conclusion・Next plan・スケジュール・論文紹介（Abstract・図＋引用・要点）など 30 種類前後。
  一覧と見本は `SLIDE_RULES.md` §4 と `examples/make_slide_type_catalog.py`。
- **ルールの自動チェック**（`audit(prs)`）: 20 pt 未満の文字・許可色（黒・赤・青）以外・結論行が 2 行以上・
  最下段の文字の枠囲み・色の説明文・出典の欠け・タイトル重複・文書プロパティの個人情報・草稿の残り。
- **共同編集つき保存**（`save()` → `deck_merge.py`）: 再生成しても、PowerPoint で人が変えた文字・位置・表・画像・
  書式（枠線・塗り）・ノート・足したスライドは残る。生成側の変更は人が触っていない部分にだけ入り、両方が変えた所は
  「衝突」として報告して人の版を残す。保存のたびに `output/bg/` へ時刻付きで退避し、PowerPoint で開いていれば止まる。
- **出典の作法**: 下端のフッタは論文・資料の引用 `[n]` だけにそろえる（データの注記はノートへ）。
- **変数のイタリック**: `Re`・`Cl/Cd`・`x/c` などを保存時に自動で斜体・下付きにする。

## 必要なもの

- Python 3.10 以上、`pip install -r requirements.txt`
- PDF 化（任意）: WSL + Windows の PowerPoint（`scripts/export_pdf.sh`）。他の環境では PowerPoint などで開いて書き出す。

## 使い方

```python
from slide_template import (new_presentation, save, audit, add_title_slide, add_section_slide,
                            add_content_slide, add_picture_fit, add_conclusion, set_notes,
                            CT_BODY_L, CT_BODY_W, FIG_AREA_T, FIG_AREA_H)

prs = new_presentation()
add_title_slide(prs, title="My Study", seminar_label="Seminar (1), 2026/01/01",
                lab_name="Example Lab.", member_grade_name="M1  Name", supervisors=[])
add_section_slide(prs, "Background and Objective")
s, _ = add_content_slide(prs, "Result", 1)
add_picture_fit(s, "fig.png", CT_BODY_L, FIG_AREA_T, CT_BODY_W, FIG_AREA_H)
add_conclusion(s, "One message per slide, in red at the bottom.")
set_notes(s, "発表の原稿", qa=[("想定質問", "回答")])
print(audit(prs))            # [] なら合格
save(prs, "output/my_deck.pptx")
```

型の見本を作る: `python3 examples/make_slide_type_catalog.py` → `output/slide_type_catalog.pptx`

PDF 化（WSL）: `bash scripts/export_pdf.sh output/my_deck.pptx`

テスト: `python3 -m unittest discover -s tests`

## デザインとルール

`SLIDE_RULES.md` が正典（文字は 20 pt 以上、1 スライド 1 メッセージ、色は黒・赤・青、最下段の結論は枠なし赤太字、
流入は青・流出は赤、変数はイタリック など）。デザイン（ネイビーのバー・フォント・レイアウト）は `template/base.pptx`。
自分の研究室のテンプレートを使うときは、スライドを全部消した pptx を `template/base.pptx` に置き換える
（文書プロパティ・`docProps/app.xml` のスライド題目一覧に個人情報が残らないよう注意）。

## Claude Code で使う

`skill/make-slide/SKILL.md` を `~/.claude/skills/make-slide/` にコピーすると、「スライド作って」で
このキットの手順（型関数で生成 → `audit` → PDF 目視）に沿って作る。リポジトリのフォルダで使う。

## ライセンス

MIT（`LICENSE`）。

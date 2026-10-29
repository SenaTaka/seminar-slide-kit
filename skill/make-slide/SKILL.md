---
name: make-slide
description: ゼミスライド（進捗・論文紹介）を slide_template.py の型関数で生成し、audit() と PDF 目視で 20pt 未満・禁止色・1 スライド複数メッセージを潰す。「スライド作って」「ゼミスライド」「論文紹介のパワポ」等で起動。
---

# make-slide

`slide_template.py`（このリポジトリ）でゼミスライドの PPTX を生成する。

## 最初に読む
- `SLIDE_RULES.md` — 共通ルール R1〜R13（20 pt 以上・1 スライド 1 メッセージ・色は黒/赤/青・最下段の結論は枠なし赤太字・
  流入は青/流出は赤・下端は引用だけ）、毎回のゼミの骨格、論文紹介の型、スライド型カタログ（関数名つき）、提出前チェックリスト。
- 型の見本: `python3 examples/make_slide_type_catalog.py` → `output/slide_type_catalog.pptx`。

## 手順
1. 指示から構成（型・タイトル・メッセージ 1 文・図）を決める。型は手書きで組まず `add_*_slide` を使う。
2. 図は貼り付け寸法の `figsize`・文字 20 pt 以上で作る（matplotlib）。
3. 生成スクリプトを書いて実行し、`audit(prs)` が 0 件であることを確かめてから `save(prs, path)`。
   - `save()` は既定で共同編集つき保存（`deck_merge.py`）: PowerPoint で開いていれば止まり、既存ファイルは
     `output/bg/` に時刻付きで退避し、人が編集した部分（文字・位置・表・画像・書式・ノート）は残して 3-way でマージする。
   - `save()` は既定でフッタを引用だけにする（`keep_citation_footers`）。自分のデータの注記はノートへ（`set_notes`）。
4. PDF 化（Windows + PowerPoint の WSL なら `bash scripts/export_pdf.sh output/<name>.pptx`）して全ページを目視し、
   余白・はみ出し・重なり・禁止色を直して再生成する。
5. `deck_merge.py` を変えたら `python3 -m unittest discover -s tests` を通す。

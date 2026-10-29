"""
slide_template.py — 修士ゼミスライドテンプレートライブラリ

テンプレート: template/base.pptx（ネイビーのバーとテーマ・フォント・レイアウト）。

カラールール:
  - アクセントは赤 (#FF0000) と青 (#0000FF 原色) のみ（規約全体は SLIDE_RULES.md）
  - 赤: 最重要の発見・結論
  - 青: ラベル・矢印・二次的強調
  - 本文・タイトルはすべて黒
  - 水色・黄・緑・オレンジは使わない
"""

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
import copy
import re
from lxml import etree

# ─── テンプレートパス ─────────────────────────────────────────────────────────
TEMPLATE_PATH = str(__import__("pathlib").Path(__file__).resolve().parent / "template" / "base.pptx")   # 掃除済み（10/29）

# ─── スライドサイズ (EMU) ─────────────────────────────────────────────────────
SLIDE_W = 12192000
SLIDE_H = 6858000

# ─── カラーパレット ────────────────────────────────────────────────────────────
NAVY       = RGBColor(0x00, 0x20, 0x60)  # #002060  ネイビーバー・区切り線
RED        = RGBColor(0xFF, 0x00, 0x00)  # #FF0000  強調1 (結論・最重要)
BLUE       = RGBColor(0x00, 0x00, 0xFF)  # #0000FF  強調2 (ラベル・矢印) 原色青
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
BLACK      = RGBColor(0x00, 0x00, 0x00)
TABLE_HDR  = RGBColor(0x00, 0x20, 0x60)  # ヘッダー背景 = ネイビー
TABLE_ROW1 = RGBColor(0xF2, 0xF2, 0xF2)  # 交互行 (薄灰 — 水色禁止)
TABLE_ROW2 = RGBColor(0xFF, 0xFF, 0xFF)  # 交互行 (白)

# ─── レイアウト index ────────────────────────────────────────────────────────
LAYOUT_TITLE   = 0   # タイトル スライド
LAYOUT_CONTENT = 1   # タイトルとコンテンツ (未使用: 手動レイアウトに切り替え)
LAYOUT_BLANK   = 6   # 白紙 (コンテンツスライドもこれを使う)

# ─── タイトルスライド 寸法 (EMU) ─────────────────────────────────────────────
TOP_BAR_H  = 1396079
BOT_BAR_T  = 5482722
BOT_BAR_H  = SLIDE_H - BOT_BAR_T   # 1375279

TITLE_L, TITLE_T = 1524002, 1961307
TITLE_W, TITLE_H = 9143999, 2387600

SEMINAR_L, SEMINAR_T = 5800000, 1600000
SEMINAR_W, SEMINAR_H = 5400000,  450000

AUTHOR_L, AUTHOR_T = 1524000, 4700000
AUTHOR_W, AUTHOR_H = 9144000,  600000

SUP_LABEL_L, SUP_LABEL_T = 5324764, BOT_BAR_T + 80000
SUP_LABEL_W, SUP_LABEL_H = 1800000, BOT_BAR_H - 80000

SUP_NAME_L, SUP_NAME_T = 7100000, BOT_BAR_T + 80000
SUP_NAME_W, SUP_NAME_H = 5092000, BOT_BAR_H - 80000

# ─── コンテンツスライド 寸法 (EMU) ───────────────────────────────────────────
# タイトルは区切り線の「上」に配置 (LAYOUT_BLANK で手動制御)
_MARGIN = 457200          # 左右マージン (~1.2 cm)

CT_TITLE_L = _MARGIN + 900000
CT_TITLE_T = 40000        # 参照デッキはタイトル 40–44pt 中央・線の直上
CT_TITLE_W = SLIDE_W - 2 * 457200 - 2 * 900000   # 中央寄せ・右にスライド番号
CT_TITLE_H = 700000       # 40pt 1行分

_CT_LINE_T = CT_TITLE_T + CT_TITLE_H + 20000   # = 760000（従来と同じ位置）
_CT_LINE_H = 25000         # 区切り線の太さ (navy)

SLIDENUM_L = 10700000
SLIDENUM_T = CT_TITLE_T
SLIDENUM_W = 1200000
SLIDENUM_H = 450000

CT_BODY_L = _MARGIN
CT_BODY_T = _CT_LINE_T + _CT_LINE_H + 80000    # = 865000
CT_BODY_W = SLIDE_W - 2 * _MARGIN              # = 11277600
CT_BODY_H = SLIDE_H - CT_BODY_T - 200000       # = 5793000
BODY_RATIO = CT_BODY_W / CT_BODY_H             # ≈ 1.95  (図の figsize 目安)

# 本文を「図エリア(上) + 結論ストリップ(下)」に分割する標準寸法。
# 図エリアを目一杯使い、余白を抑えるため caller はこれらを使う。
CONCL_H    = 720000                            # 下部 結論ストリップ高さ
CONCL_T    = CT_BODY_T + CT_BODY_H - CONCL_H    # 結論ストリップ 上端
FIG_AREA_T = CT_BODY_T                          # 図エリア 上端
FIG_AREA_H = CT_BODY_H - CONCL_H - 80000        # 図エリア 高さ (結論との隙間込み)
FIG_AREA_RATIO = CT_BODY_W / FIG_AREA_H         # ≈ 2.2  (全幅図の figsize 目安)

REF_L = _MARGIN
REF_T = SLIDE_H - 270000
REF_W = CT_BODY_W
REF_H = 220000


TITLE_PT = 40   # コンテンツスライドのタイトル（参照デッキのマスターは 44pt。日本語の長さを考え 40pt）
BODY_PT = 28    # 本文の既定（参照デッキの最頻 28–32pt）

# ─── プレゼンテーション生成 ───────────────────────────────────────────────────

def new_presentation() -> Presentation:
    """テンプレート PPTX を読み込んでスライドを全削除。テーマ・フォント・レイアウトを継承。"""
    prs = Presentation(TEMPLATE_PATH)
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        rId = sldId.get(qn("r:id"))
        prs.part.drop_rel(rId)
    sldIdLst.clear()
    _scrub_metadata(prs)
    return prs


def _scrub_metadata(prs):
    """テンプレート（参照デッキ pptx）から継承した個人情報を消す: 文書プロパティと ppt/authors.xml・commentAuthors。"""
    cp = prs.core_properties
    for k in ("author", "last_modified_by", "title", "subject", "keywords", "comments", "category"):
        setattr(cp, k, "")
    cp.revision = 1
    for rel in list(prs.part.rels.values()):
        if rel.reltype.endswith(("/authors", "/commentAuthors")):
            prs.part.drop_rel(rel.rId)
    _scrub_app_props(prs)


APP_XML_MIN = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/'
               'officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/'
               'docPropsVTypes"><Application>Microsoft Office PowerPoint</Application></Properties>').encode()


def _app_part(prs):
    for part in prs.part.package.iter_parts():
        if str(part.partname) == "/docProps/app.xml":
            return part
    return None


def _scrub_app_props(prs):
    """docProps/app.xml を最小にする。テンプレート元デッキのスライド題目一覧（TitlesOfParts）・会社名などが残るため（10/29 実測）。
    python-pptx は app.xml を更新しないので、生成デッキに元デッキの全題目が埋め込まれたままになる。"""
    part = _app_part(prs)
    if part is not None and part.blob != APP_XML_MIN:
        part._blob = APP_XML_MIN


# ─── 内部ヘルパー ──────────────────────────────────────────────────────────────

def _txbox(slide, text, l, t, w, h, size_pt=20, bold=False,
           color=BLACK, align=PP_ALIGN.LEFT, wrap=True) -> object:
    """テキストボックスを追加して shape を返す。"""
    box = slide.shapes.add_textbox(Emu(l), Emu(t), Emu(w), Emu(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def _rect(slide, l, t, w, h, fill_color=None, line_color=None, line_width_pt=1.5) -> object:
    """矩形を追加して shape を返す。fill/line を設定。"""
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Emu(l), Emu(t), Emu(w), Emu(h)
    )
    if fill_color is not None:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color is not None:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_width_pt)
    else:
        shape.line.fill.background()
    return shape


# ─── 公開 API ─────────────────────────────────────────────────────────────────

def styled_run(para, text: str, color=BLACK, bold: bool = False, size_pt=None):
    """段落に styled run を追加して返す。"""
    run = para.add_run()
    run.text = text
    run.font.bold = bold
    run.font.color.rgb = color
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    return run


def add_title_slide(prs, title: str, seminar_label: str,
                    lab_name: str, member_grade_name: str,
                    supervisors: list, subtitle: str = None) -> object:
    """
    タイトルスライドを追加。

    Parameters
    ----------
    title            : 発表タイトル
    seminar_label    : 右上ラベル e.g. "Master Seminar M2(3), 2026/05/21"
    lab_name         : 研究室名 e.g. "Example Lab."
    member_grade_name: e.g. "M2  Your Name"
    supervisors      : 指導教員リスト e.g. ["Prof. A", "Prof. B"]（指導教員のみ。論文著者は入れない）
    subtitle         : タイトル下の1行（論文紹介の書誌など・青 20pt）。None なら出さない
    """
    slide = prs.slides.add_slide(prs.slide_layouts[LAYOUT_BLANK])

    _rect(slide, 0, 0, SLIDE_W, TOP_BAR_H, fill_color=NAVY)
    _rect(slide, 0, BOT_BAR_T, SLIDE_W, BOT_BAR_H, fill_color=NAVY)

    _txbox(slide, seminar_label,
           SEMINAR_L, SEMINAR_T, SEMINAR_W, SEMINAR_H,
           size_pt=20, align=PP_ALIGN.RIGHT, color=BLACK)

    _txbox(slide, title,
           TITLE_L, TITLE_T, TITLE_W, TITLE_H,
           size_pt=36, align=PP_ALIGN.CENTER, color=BLACK, wrap=True)

    if subtitle:
        _txbox(slide, subtitle, TITLE_L, 3900000, TITLE_W, 500000,
               size_pt=20, color=BLUE, align=PP_ALIGN.CENTER)

    author_text = f"{lab_name}\t\t{member_grade_name}".strip()
    if author_text:
        _txbox(slide, author_text,
               AUTHOR_L, AUTHOR_T, AUTHOR_W, AUTHOR_H,
               size_pt=24, color=BLACK)

    if not supervisors:          # 名前を出さない版（試作・外部共有）
        return slide

    _txbox(slide, "Supervisor",
           SUP_LABEL_L, SUP_LABEL_T, SUP_LABEL_W, SUP_LABEL_H,
           size_pt=22, color=WHITE)

    sup_box = slide.shapes.add_textbox(
        Emu(SUP_NAME_L), Emu(SUP_NAME_T),
        Emu(SUP_NAME_W), Emu(SUP_NAME_H))
    sup_tf = sup_box.text_frame
    sup_tf.word_wrap = True
    for i, name in enumerate(supervisors):
        para = sup_tf.paragraphs[0] if i == 0 else sup_tf.add_paragraph()
        run = para.add_run()
        run.text = name
        run.font.size = Pt(22)
        run.font.color.rgb = WHITE

    return slide


def add_section_slide(prs, section_name: str, color=BLACK, underline=False, items=None, current_idx=None) -> object:
    """セクション扉。黒太字 40pt 中央・下線なし・番号なし。
    color=BLUE, underline=True で旧版（青＋ネイビー線）。"""
    slide = prs.slides.add_slide(prs.slide_layouts[LAYOUT_BLANK])
    b = _txbox(slide, section_name,
               800000, 2400000, SLIDE_W - 1600000, 1400000,
               size_pt=40, bold=True, align=PP_ALIGN.CENTER, color=color)
    b.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    if underline:
        _rect(slide, 1000000, 3880000, 10192000, 30000, fill_color=NAVY)
    if items:   # 番号付き小目次: 扉名を上に寄せ、下に番号付きの手順。現在 = 赤、済み/先 = 灰
        b.top = Emu(1300000)
        box = slide.shapes.add_textbox(Emu(2600000), Emu(2900000), Emu(SLIDE_W - 5200000), Emu(3000000))
        tf = box.text_frame
        tf.word_wrap = True
        for i, it in enumerate(items):
            p_ = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p_.space_after = Pt(14)
            col = BLACK if current_idx is None else (RED if i == current_idx else GRAY)
            styled_run(p_, f"{i + 1}.  {it}", color=col, bold=True, size_pt=36)
    return slide


def add_content_slide(prs, title: str, slide_number=None):
    """
    標準コンテンツスライド (LAYOUT_BLANK 使用、タイトルは区切り線の上)。

    Returns
    -------
    (slide, body_tf)
      body_tf: コンテンツエリアの TextFrame。caller が段落/画像を追加する。
    """
    slide = prs.slides.add_slide(prs.slide_layouts[LAYOUT_BLANK])

    # タイトル (区切り線の上)
    tsz = TITLE_PT if len(title) <= 34 else (36 if len(title) <= 44 else 32)   # 参照デッキ: 44pt、長い題は 36pt
    tb = _txbox(slide, title,
                CT_TITLE_L, CT_TITLE_T, CT_TITLE_W, CT_TITLE_H,
                size_pt=tsz, color=BLACK, align=PP_ALIGN.CENTER)
    tb.name = "TITLE"
    tb.text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM

    # 区切り線 (navy, 全幅)
    _rect(slide, 0, _CT_LINE_T, SLIDE_W, _CT_LINE_H, fill_color=NAVY)

    # スライド番号 (右上)
    if slide_number is not None:
        _txbox(slide, str(slide_number),
               SLIDENUM_L, SLIDENUM_T, SLIDENUM_W, SLIDENUM_H,
               size_pt=22, align=PP_ALIGN.RIGHT, color=BLACK)

    # コンテンツエリア
    body_box = slide.shapes.add_textbox(
        Emu(CT_BODY_L), Emu(CT_BODY_T), Emu(CT_BODY_W), Emu(CT_BODY_H))
    body_tf = body_box.text_frame
    body_tf.word_wrap = True
    return slide, body_tf


def add_reference(slide, ref_text: str, font_size: float = 20.0):
    """スライド下部に参考文献テキストを追加。"""
    _txbox(slide, ref_text,
           REF_L, REF_T, REF_W, REF_H,
           size_pt=font_size, color=BLACK)


def add_conclusion(slide, text: str, color=RED, size_pt: float = 26.0):
    """
    本文下部に中央寄せの結論ストリップを置く（赤太字が標準）。
    位置は CONCL_T/CONCL_H 固定なので全スライドで結論行が揃う。
    """
    b = _txbox(slide, text,
               CT_BODY_L, CONCL_T, CT_BODY_W, CONCL_H,
               size_pt=size_pt, bold=True, color=color,
               align=PP_ALIGN.CENTER)
    b.name = "MSG"
    return b


def fig_area_box(width_frac: float = 1.0, halign: str = 'center'):
    """
    図エリア(結論ストリップの上)の配置ボックス (l, t, w, h) を EMU で返す。
    width_frac で幅割合を指定。halign='center'/'left'/'right' で水平寄せ。
    figsize は w/h ≈ FIG_AREA_RATIO×width_frac に合わせると余白が出ない。
    """
    w = int(CT_BODY_W * width_frac)
    if halign == 'center':
        l = CT_BODY_L + (CT_BODY_W - w) // 2
    elif halign == 'right':
        l = CT_BODY_L + (CT_BODY_W - w)
    else:
        l = CT_BODY_L
    return l, FIG_AREA_T, w, FIG_AREA_H


def add_red_box(slide, text: str, l: int, t: int, w: int, h: int,
                font_size: float = 20.0) -> object:
    """
    赤枠ボックス（黒テキスト）。重要な発見・結論の強調に使う。
    Parameters: l, t, w, h は EMU。
    """
    shape = slide.shapes.add_shape(1, Emu(l), Emu(t), Emu(w), Emu(h))
    shape.fill.background()
    shape.line.color.rgb = RED
    shape.line.width = Pt(2)
    tf = shape.text_frame
    tf.word_wrap = True
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.LEFT
    run = para.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.color.rgb = BLACK
    return shape


def add_picture_fit(slide, img_path: str, l: int, t: int,
                    max_w: int, max_h: int, anchor: str = 'center') -> object:
    """
    アスペクト比を保ったまま画像を挿入。max_w × max_h のボックス内に収める。

    Parameters
    ----------
    l, t, max_w, max_h : EMU
    anchor : 'center' | 'top' | 'topleft'
    """
    try:
        from PIL import Image as PILImage
        with PILImage.open(img_path) as im:
            iw, ih = im.size
        scale = min(max_w / iw, max_h / ih)
        w = int(iw * scale)
        h = int(ih * scale)
    except Exception:
        w, h = max_w, max_h

    if anchor == 'center':
        l_off = (max_w - w) // 2
        t_off = (max_h - h) // 2
    elif anchor == 'top':
        l_off = (max_w - w) // 2
        t_off = 0
    else:  # topleft
        l_off = t_off = 0

    return slide.shapes.add_picture(
        img_path, Emu(l + l_off), Emu(t + t_off), Emu(w), Emu(h))


def add_image_slide(prs, title: str, img_path: str,
                    caption: str = None, caption_color=BLACK,
                    slide_number=None) -> object:
    """
    画像1枚のコンテンツスライド。アスペクト比を保って挿入する。
    """
    slide, _ = add_content_slide(prs, title, slide_number)

    img_max_h = CT_BODY_H - (600000 if caption else 0)
    add_picture_fit(slide, img_path,
                    CT_BODY_L, CT_BODY_T, CT_BODY_W, img_max_h, anchor='top')

    if caption:
        cap_t = CT_BODY_T + img_max_h + 80000
        _txbox(slide, caption,
               CT_BODY_L, cap_t, CT_BODY_W, 520000,
               size_pt=22, bold=True, align=PP_ALIGN.CENTER, color=caption_color)

    return slide


def add_table_slide(prs, title: str, headers: list, rows: list,
                    slide_number=None) -> object:
    """
    表スライド。
    ヘッダー: ネイビー背景+白文字。行: 薄灰/白交互 (水色禁止)。
    """
    slide, _ = add_content_slide(prs, title, slide_number)

    n_cols = len(headers)
    n_rows = len(rows) + 1
    col_w  = CT_BODY_W // n_cols
    row_h  = min(Emu(600000), Emu(CT_BODY_H // n_rows))

    table = slide.shapes.add_table(
        n_rows, n_cols,
        Emu(CT_BODY_L), Emu(CT_BODY_T + 100000),
        Emu(CT_BODY_W), Emu(row_h * n_rows)
    ).table

    for j, hdr in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = hdr
        cell.fill.solid()
        cell.fill.fore_color.rgb = TABLE_HDR
        para = cell.text_frame.paragraphs[0]
        para.alignment = PP_ALIGN.CENTER
        run = para.runs[0] if para.runs else para.add_run()
        run.font.color.rgb = WHITE
        run.font.bold = True
        run.font.size = Pt(22)

    for i, row in enumerate(rows):
        bg = TABLE_ROW1 if i % 2 == 0 else TABLE_ROW2
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            para = cell.text_frame.paragraphs[0]
            para.alignment = PP_ALIGN.CENTER
            run = para.runs[0] if para.runs else para.add_run()
            run.font.size = Pt(20)
            run.font.color.rgb = BLACK

    return slide


def add_table_at(slide, headers, rows, l, t, w, row_h=560000, col_fracs=None, hdr_pt=22, row_pt=22,
                 red_cells=(), blue_cells=()):
    """
    位置指定の表（add_table_slide と同じ配色: ネイビー見出し＋薄灰/白の交互行）。表の上下に説明文や
    add_conclusion を置くとき用。red_cells / blue_cells は本文行の (行, 列)（0 始まり・見出しを除く）で、
    言えたこと = 赤太字、留保 = 青太字（R3）。col_fracs は列幅の比（合計 1）。
    """
    n_cols = len(headers)
    tb = slide.shapes.add_table(len(rows) + 1, n_cols, Emu(l), Emu(t), Emu(w), Emu(row_h * (len(rows) + 1))).table
    if col_fracs:
        for j, f in enumerate(col_fracs):
            tb.columns[j].width = Emu(int(w * f))
    for j, hd in enumerate(headers):
        c = tb.cell(0, j)
        c.fill.solid()
        c.fill.fore_color.rgb = TABLE_HDR
        p = c.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        styled_run(p, str(hd), color=WHITE, bold=True, size_pt=hdr_pt)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = tb.cell(i + 1, j)
            c.fill.solid()
            c.fill.fore_color.rgb = TABLE_ROW1 if i % 2 == 0 else TABLE_ROW2
            p = c.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            col = RED if (i, j) in red_cells else BLUE if (i, j) in blue_cells else BLACK
            styled_run(p, str(val), color=col, bold=col is not BLACK, size_pt=row_pt)
    return tb


# ─── スライド型（参照デッキから抽出・SLIDE_RULES.md §3 のカタログと 1 対 1） ─────────

GRAY = RGBColor(0x76, 0x76, 0x76)   # 文字に使える最も淡い灰（白地 4.54:1・WCAG AA）。目次の非現在項目・非現在ステップ


def _dash(shape, val='lgDash'):
    """図形の枠線を破線にする。"""
    ln = shape._element.spPr.find(qn('a:ln'))
    if ln is None:
        ln = etree.SubElement(shape._element.spPr, qn('a:ln'))
    for old in ln.findall(qn('a:prstDash')):
        ln.remove(old)
    d = etree.SubElement(ln, qn('a:prstDash'))
    d.set('val', val)


def _line(slide, x1, y1, x2, y2, color=BLACK, width_pt=1.5, dash=None, arrow_end=False):
    """直線（矢印可）。座標は EMU。"""
    from pptx.enum.shapes import MSO_CONNECTOR
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(x1), Emu(y1), Emu(x2), Emu(y2))
    c.line.color.rgb = color
    c.line.width = Pt(width_pt)
    ln = c.line._get_or_add_ln()
    if dash:
        d = etree.SubElement(ln, qn('a:prstDash'))
        d.set('val', dash)
    if arrow_end:
        t = etree.SubElement(ln, qn('a:tailEnd'))
        t.set('type', 'triangle')
        t.set('w', 'med')
        t.set('len', 'med')
    return c


def _box(slide, text, l, t, w, h, size_pt=24, color=BLACK, line=BLACK, bold=False,
         align=PP_ALIGN.CENTER, line_pt=1.5):
    """枠付きテキスト（塗りなし）。text の改行で段落を分ける。"""
    from pptx.enum.text import MSO_ANCHOR
    shp = slide.shapes.add_shape(1, Emu(l), Emu(t), Emu(w), Emu(h))
    shp.fill.background()
    shp.line.color.rgb = line
    shp.line.width = Pt(line_pt)
    tf = shp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, ln in enumerate(text.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        styled_run(p, ln, color=color, bold=bold, size_pt=size_pt)
    return shp


def add_highlight_box(slide, l, t, w, h, label="This time", label_pos="bottom-right"):
    """型 H「This time / Now working on」: 赤破線枠で今回の範囲を囲む。"""
    shp = slide.shapes.add_shape(1, Emu(l), Emu(t), Emu(w), Emu(h))
    shp.fill.background()
    shp.line.color.rgb = RED
    shp.line.width = Pt(3)
    _dash(shp)
    if label:
        lw, lh = 2600000, 450000
        lx = l + w - lw - 80000 if label_pos.endswith('right') else l + 80000
        ly = t + h - lh - 110000 if label_pos.startswith('bottom') else t + 90000
        _txbox(slide, label, lx, ly, lw, lh, size_pt=24, bold=True, color=RED,
               align=PP_ALIGN.RIGHT if label_pos.endswith('right') else PP_ALIGN.LEFT)
    return shp


def add_flow(slide, steps, l, t, w, h, direction='down', size_pt=24, highlight=None):
    """
    処理の流れ（枠 → 矢印 → 枠）。direction='down' は縦、'right' は横。
    steps: ["ステップ1", "ステップ2\n補足", ...]（各ステップの1行目は青太字）
    highlight: (i0, i1) を渡すとその範囲を add_highlight_box で囲む（label は "This time"）。
    戻り値: 各ステップの (l, t, w, h) リスト
    """
    n = len(steps)
    boxes = []
    if direction == 'down':
        gap = 360000
        bh = (h - gap * (n - 1)) // n
        for i, st in enumerate(steps):
            y = t + i * (bh + gap)
            boxes.append((l, y, w, bh))
    else:
        gap = 520000
        bw = (w - gap * (n - 1)) // n
        for i, st in enumerate(steps):
            x = l + i * (bw + gap)
            boxes.append((x, t, bw, h))
    for i, (st, (bl, bt, bw_, bh_)) in enumerate(zip(steps, boxes)):
        shp = _box(slide, st, bl, bt, bw_, bh_, size_pt=size_pt)
        tf = shp.text_frame
        for r in tf.paragraphs[0].runs:
            r.font.color.rgb = BLUE
            r.font.bold = True
        if i < n - 1:
            nl, nt, nw, nh = boxes[i + 1]
            if direction == 'down':
                _line(slide, bl + bw_ // 2, bt + bh_, bl + bw_ // 2, nt, color=BLUE, width_pt=2.5, arrow_end=True)
            else:
                _line(slide, bl + bw_, bt + bh_ // 2, nl, bt + bh_ // 2, color=BLUE, width_pt=2.5, arrow_end=True)
    if highlight:
        i0, i1 = highlight
        a, b = boxes[i0], boxes[i1]
        pad = 120000
        x0, y0 = min(a[0], b[0]) - pad, min(a[1], b[1]) - pad
        x1 = max(a[0] + a[2], b[0] + b[2]) + pad
        y1 = max(a[1] + a[3], b[1] + b[3]) + pad
        add_highlight_box(slide, x0, y0, x1 - x0, y1 - y0, label=None)
    return boxes


def add_agenda_slide(prs, sections, current_idx):
    """
    型 A「目次」: 現在の節を赤太字 36pt、他を灰 28pt。スライド番号なし。
    
    """
    from pptx.enum.text import MSO_ANCHOR
    slide = prs.slides.add_slide(prs.slide_layouts[LAYOUT_BLANK])
    box = slide.shapes.add_textbox(Emu(1500000), Emu(900000), Emu(9200000), Emu(5000000))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, name in enumerate(sections):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(18)
        cur = (i == current_idx)
        styled_run(p, f"{i + 1}.  {name}", color=RED if cur else GRAY, bold=cur,
                   size_pt=36 if cur else 28)
    return slide


def add_objective_slide(prs, objective, steps=None, title="目的", slide_number=None,
                        conclusion=None, highlight=None, direction='down'):
    """
    型 O「目的」: 上部に黒枠の目的 1 文（➤）、下に処理の流れ（任意）、最下段に結論（任意）。
    
    """
    slide, _ = add_content_slide(prs, title, slide_number)
    obj_h = 1000000
    _box(slide, "➤ " + objective, CT_BODY_L, FIG_AREA_T + 60000, CT_BODY_W, obj_h,
         size_pt=26, align=PP_ALIGN.LEFT)
    if steps:
        top = FIG_AREA_T + obj_h + 300000
        bottom = (CONCL_T - 100000) if conclusion else (CT_BODY_T + CT_BODY_H - 100000)
        if direction == 'down':
            fw = CT_BODY_W * 60 // 100
            add_flow(slide, steps, CT_BODY_L + (CT_BODY_W - fw) // 2, top, fw, bottom - top,
                     direction='down', highlight=highlight)
        else:
            fh = min(1500000, bottom - top)
            add_flow(slide, steps, CT_BODY_L, top + (bottom - top - fh) // 2, CT_BODY_W, fh,
                     direction='right', highlight=highlight)
    if conclusion:
        add_conclusion(slide, conclusion)
    return slide


def add_conclusion_slide(prs, lead=None, items=(), title="Conclusion", slide_number=None, message=None,
                         numbered=False, size_pt=32):
    """
    型 C「Conclusion」。既定は 参照デッキどおり「•」の箇条 32pt を縦中央、リード枠・番号なし。
    色は 参照デッキの使い分け: 赤 = 言えたこと（肯定）、青 = 留保・課題、黒 = 説明（items の各要素に [(text, color, bold)] で指定）。
    lead を渡すと黒枠のリード文、numbered=True で番号付き。message で下部の赤 1 行。
    """
    from pptx.enum.text import MSO_ANCHOR
    slide, _ = add_content_slide(prs, title, slide_number)
    top = FIG_AREA_T
    if lead:
        lead_h = 900000
        _box(slide, lead, CT_BODY_L, FIG_AREA_T + 60000, CT_BODY_W, lead_h, size_pt=28, align=PP_ALIGN.LEFT)
        top = FIG_AREA_T + lead_h + 250000
    bottom = CONCL_T - 80000 if message else CT_BODY_T + CT_BODY_H
    box = slide.shapes.add_textbox(Emu(CT_BODY_L + 300000), Emu(top), Emu(CT_BODY_W - 600000), Emu(bottom - top))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(18)
        if numbered:
            styled_run(p, f"{i + 1}.  ", color=BLUE, bold=True, size_pt=size_pt)
        else:
            styled_run(p, "•  ", color=BLACK, size_pt=size_pt)
        runs = [(it, BLACK, False)] if isinstance(it, str) else it
        for text, color, bold in runs:
            styled_run(p, text, color=color, bold=bold, size_pt=size_pt)
    if message:
        add_conclusion(slide, message)
    return slide


def add_plan_options_slide(prs, lead, plans, selected=None, title="今後の予定",
                           slide_number=None, message=None):
    """
    型 P「今後の予定（案の比較）」: 1 行のリード + 案を縦の点線で区切った列。
    選んだ案は赤破線で囲む。plans: [("Plan A", ["行1", "↓", "行2"]), ...]。
    """
    from pptx.enum.text import MSO_ANCHOR
    slide, _ = add_content_slide(prs, title, slide_number)
    _txbox(slide, lead, CT_BODY_L, FIG_AREA_T + 40000, CT_BODY_W, 550000, size_pt=24)
    top = FIG_AREA_T + 750000
    bottom = CONCL_T - 100000 if message else CT_BODY_T + CT_BODY_H - 100000
    n = len(plans)
    cw = CT_BODY_W // n
    for i, (name, lines) in enumerate(plans):
        x = CT_BODY_L + i * cw
        box = slide.shapes.add_textbox(Emu(x + 150000), Emu(top + 100000), Emu(cw - 300000), Emu(bottom - top - 200000))
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        styled_run(tf.paragraphs[0], name, color=RED if i == selected else BLUE, bold=True, size_pt=28)
        for ln in lines:
            p = tf.add_paragraph()
            styled_run(p, ln, color=BLACK, bold=False, size_pt=24)
        if i > 0:
            _line(slide, x, top, x, bottom, color=BLACK, width_pt=1.25, dash='sysDot')
        if i == selected:
            add_highlight_box(slide, x + 60000, top, cw - 120000, bottom - top, label=None)
    if message:
        add_conclusion(slide, message)
    return slide


def add_schedule_slide(prs, months, tasks, title="今後の予定", slide_number=None,
                       message=None, now=None, done_until=None):
    """
    型 S「スケジュール（ガント）」: ネイビー見出し行に月、各行に矢印でタスク。
    months: ["10月", "11月", ...]
    tasks : [(label, start, end, kind), ...]  start/end は月 index（小数可・0 = 先頭月の頭）
            kind = 'task'（青矢印）| 'deadline'（赤・発表/提出などの締切。start=end で点）| 'done'（黒）
    now   : 現在位置（月 index）。赤破線の縦線と「Now」。
    done_until: 月 index。そこまでを白の半透明で淡くする。
    
    """
    slide, _ = add_content_slide(prs, title, slide_number)
    top = FIG_AREA_T + 40000
    bottom = CONCL_T - 100000 if message else CT_BODY_T + CT_BODY_H - 100000
    hdr_h = 520000
    nm = len(months)
    cw = CT_BODY_W // nm
    for i, m in enumerate(months):
        x = CT_BODY_L + i * cw
        cell = _rect(slide, x, top, cw - 20000, hdr_h, fill_color=NAVY)
        tf = cell.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        styled_run(p, m, color=WHITE, bold=True, size_pt=22)
        if i > 0:
            _line(slide, x, top + hdr_h, x, bottom, color=GRAY, width_pt=1.0, dash='sysDot')
    body_t = top + hdr_h + 100000
    rh = min(760000, (bottom - body_t) // max(len(tasks), 1))
    for k, (label, s0, s1, kind) in enumerate(tasks):
        y = body_t + k * rh
        x0 = CT_BODY_L + int(s0 * cw)
        x1 = CT_BODY_L + int(s1 * cw)
        color = {'task': BLUE, 'deadline': RED, 'done': BLACK}[kind]
        ay = y + rh - 150000
        if x1 - x0 > 60000:
            _line(slide, x0, ay, x1, ay, color=color, width_pt=3, arrow_end=True)
        else:   # 点の締切 = ▼
            tri = slide.shapes.add_shape(7, Emu(x0 - 110000), Emu(ay - 200000), Emu(220000), Emu(200000))
            tri.rotation = 180
            tri.fill.solid()
            tri.fill.fore_color.rgb = color
            tri.line.fill.background()
        lx = max(CT_BODY_L, min(x0, CT_BODY_L + CT_BODY_W - 4200000))
        _txbox(slide, label, lx, ay - 520000, 4200000, 420000, size_pt=20,
               color=color, bold=(kind == 'deadline'))
    if done_until is not None:
        set_done_shade(slide, CT_BODY_L, top + hdr_h, int(done_until * cw), bottom - top - hdr_h)
    if now is not None:
        xn = CT_BODY_L + int(now * cw)
        _line(slide, xn, top + hdr_h, xn, bottom, color=RED, width_pt=2, dash='dash')
        _txbox(slide, "Now", xn - 500000, bottom - 420000, 1000000, 420000, size_pt=20,
               color=RED, bold=True, align=PP_ALIGN.CENTER)
    if message:
        add_conclusion(slide, message)
    return slide


def add_closing_slide(prs, text="Thank you for your attention"):
    """型 E「終わり」: 中央 1 行。"""
    slide = prs.slides.add_slide(prs.slide_layouts[LAYOUT_BLANK])
    _txbox(slide, text, 0, 2900000, SLIDE_W, 800000, size_pt=36, align=PP_ALIGN.CENTER)
    return slide


# ─── 追加の型（2026/10/13）。淡色塗りは黒/赤/青の規約に置換済み ──

def _paras(tf, items, size_pt=24, space_pt=12):
    """items: [str | [(text, color, bold), ...]] を段落として流し込む。"""
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 and not tf.paragraphs[0].runs else tf.add_paragraph()
        p.space_after = Pt(space_pt)
        runs = [(it, BLACK, False)] if isinstance(it, str) else it
        for text, color, bold in runs:
            styled_run(p, text, color=color, bold=bold, size_pt=size_pt)


def add_buildup_slides(prs, title, steps, problem, slide_number=None, build=True):
    """
    型 B「背景の段階的ビルドアップ」: 文を ▼ でつなぎ、最後に赤の課題 1 行。
    steps: ["現象", "原因", ...]（2 段目以降は青太字）。build=True で 1 段ずつ追記した連続スライドを返す
    （発表ではページ送りで積み上がって見える）。
    """
    from pptx.enum.text import MSO_ANCHOR
    stages = range(1, len(steps) + 2) if build else [len(steps) + 1]
    slides = []
    lines = list(steps) + [problem]
    n = len(lines)
    row_h, tri_h, gap = 800000, 260000, 120000
    block = n * row_h + (n - 1) * (tri_h + 2 * gap)
    y0 = FIG_AREA_T + max(0, (FIG_AREA_H + CONCL_H - block) // 2)   # 完成形で縦中央。途中段も同じ位置
    for k in stages:
        slide, _ = add_content_slide(prs, title, slide_number)
        y = y0
        for i, text in enumerate(lines[:k]):
            last = (i == len(steps))
            color = RED if last else (BLACK if i == 0 else BLUE)
            b = _txbox(slide, text, CT_BODY_L, y, CT_BODY_W, row_h, size_pt=36 if last else 28,
                       bold=(i > 0), color=color, align=PP_ALIGN.CENTER)
            b.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            y += row_h
            if i < k - 1:
                tri = slide.shapes.add_shape(7, Emu(SLIDE_W // 2 - 220000), Emu(y + gap), Emu(440000), Emu(tri_h))
                tri.rotation = 180
                tri.fill.solid()
                tri.fill.fore_color.rgb = NAVY
                tri.line.fill.background()
            y += tri_h + 2 * gap
        slides.append(slide)
    return slides


def add_question_slide(prs, title, question, img_path, naming, slide_number=None):
    """
    型 Q「問い → 図 → 呼称」: 上に赤の問い 28pt、中央に図、下に赤の呼称/結論。
    "Is it possible to …?" → "This problem is called …"）。
    """
    slide, _ = add_content_slide(prs, title, slide_number)
    _txbox(slide, question, CT_BODY_L, FIG_AREA_T, CT_BODY_W, 600000, size_pt=28, color=RED)
    add_picture_fit(slide, img_path, CT_BODY_L, FIG_AREA_T + 650000, CT_BODY_W, FIG_AREA_H - 650000)
    add_conclusion(slide, naming)
    return slide


def add_compare_slide(prs, title, left, right, slide_number=None, message=None, bridge=None,
                      body_pt=24, body_anchor="top"):
    """
    型 V「二者比較」: 左 = 赤枠、右 = 青枠（角丸 2pt）。left/right = (見出し, [本文...], 画像パス or None)。
    bridge を渡すと中央下に両矢印＋その語（例「データ同化」「統合」）。
    本文が短く枠の下が空くときは body_pt=28, body_anchor="middle"（画像なしの箇条 3 行程度の目安）。
    
    """
    from pptx.enum.text import MSO_ANCHOR
    slide, _ = add_content_slide(prs, title, slide_number)
    gap = 500000 if bridge else 250000
    bw = (CT_BODY_W - gap) // 2
    bottom = (CONCL_T - 80000) if message else CT_BODY_T + CT_BODY_H - 80000
    bh = bottom - FIG_AREA_T
    for i, (side, col) in enumerate(((left, RED), (right, BLUE))):
        head, lines, img = (list(side) + [None])[:3]
        x = CT_BODY_L + i * (bw + gap)
        fr = slide.shapes.add_shape(5, Emu(x), Emu(FIG_AREA_T), Emu(bw), Emu(bh))  # ROUNDED_RECTANGLE
        fr.fill.background()
        fr.line.color.rgb = col
        fr.line.width = Pt(2)
        fr.adjustments[0] = 0.06
        _txbox(slide, head, x + 120000, FIG_AREA_T + 60000, bw - 240000, 520000, size_pt=26,
               bold=True, color=col, align=PP_ALIGN.CENTER)
        top = FIG_AREA_T + 620000
        if img:
            ih = (bh - 700000) // 2
            add_picture_fit(slide, img, x + 120000, top, bw - 240000, ih)
            top += ih + 60000
        box = slide.shapes.add_textbox(Emu(x + 150000), Emu(top), Emu(bw - 300000), Emu(FIG_AREA_T + bh - top - 60000))
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE if body_anchor == "middle" else MSO_ANCHOR.TOP
        _paras(tf, lines, size_pt=body_pt, space_pt=body_pt if body_anchor == "middle" else 12)
    if bridge:
        aw, ah = 2600000, 560000
        from pptx.enum.shapes import MSO_SHAPE
        ar = slide.shapes.add_shape(MSO_SHAPE.LEFT_RIGHT_ARROW, Emu(SLIDE_W // 2 - aw // 2), Emu(FIG_AREA_T + bh // 2 - ah // 2), Emu(aw), Emu(ah))  # LEFT_RIGHT_ARROW
        ar.fill.solid()
        ar.fill.fore_color.rgb = WHITE
        ar.line.color.rgb = BLACK
        ar.line.width = Pt(1.5)
        tf = ar.text_frame
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        styled_run(tf.paragraphs[0], bridge, color=BLACK, bold=True, size_pt=20)
    if message:
        add_conclusion(slide, message)
    return slide


def add_method_nav_slide(prs, section_title, steps, current, img_path=None, caption=None, slide_number=None):
    """
    型 N「手法フロー＋現在地ナビ」: 左に縦の手順（現在 = 赤枠・赤太字、他 = 灰）、右に赤破線枠の代表図。
    各手順の説明に入る直前に毎回差し込む（章扉の代わり）。steps: ["(1) 流れ場解析\n(OpenFOAM)", ...]
    
    """
    slide = prs.slides.add_slide(prs.slide_layouts[LAYOUT_BLANK])
    _txbox(slide, section_title, CT_TITLE_L, 150000, CT_BODY_W, 700000, size_pt=36)
    if slide_number is not None:
        _txbox(slide, str(slide_number), SLIDENUM_L, SLIDENUM_T, SLIDENUM_W, SLIDENUM_H,
               size_pt=22, align=PP_ALIGN.RIGHT)
    lw = CT_BODY_W * 46 // 100
    top, bottom = 1200000, SLIDE_H - 450000
    n = len(steps)
    ah = 380000
    bh = (bottom - top - (n - 1) * (ah + 120000)) // n
    y = top
    cur_box = None
    for i, st in enumerate(steps):
        cur = (i == current)
        shp = _box(slide, st, CT_BODY_L, y, lw, bh, size_pt=26 if cur else 24,
                   color=RED if cur else GRAY, line=RED if cur else GRAY, bold=cur,
                   line_pt=2.5 if cur else 1.0)
        for p in shp.text_frame.paragraphs:
            for r in p.runs:
                r.font.color.rgb = RED if cur else GRAY
                r.font.bold = cur
        if cur:
            cur_box = (CT_BODY_L, y, lw, bh)
        y += bh
        if i < n - 1:
            _line(slide, CT_BODY_L + lw // 2, y + 40000, CT_BODY_L + lw // 2, y + ah + 80000,
                  color=BLACK, width_pt=3, arrow_end=True)
            y += ah + 120000
    if img_path:
        rx, rw = CT_BODY_L + lw + 400000, CT_BODY_W - lw - 400000
        rt, rh = top, bottom - top
        add_highlight_box(slide, rx, rt, rw, rh, label=None)
        cap_h = 500000 if caption else 0
        add_picture_fit(slide, img_path, rx + 100000, rt + 100000, rw - 200000, rh - 200000 - cap_h)
        if caption:
            _txbox(slide, caption, rx, rt + rh - cap_h - 60000, rw, cap_h, size_pt=22, align=PP_ALIGN.CENTER)
        if cur_box:
            cl, ct, cw, chh = cur_box
            _line(slide, cl + cw, ct, rx, rt, color=RED, width_pt=1.5, dash='dash')
            _line(slide, cl + cw, ct + chh, rx, rt + rh, color=RED, width_pt=1.5, dash='dash')
    return slide


def add_step_tag(slide, text, color=BLUE):
    """型 N の詳細スライド用「ステップタグ」: タイトル直下の中央に枠付きで手順名を再掲（青枠・青太字 22pt）。"""
    w = 5200000
    return _box(slide, text, (SLIDE_W - w) // 2, CT_BODY_T, w, 480000, size_pt=22, color=color,
                line=color, bold=True)


def add_conditions_slide(prs, items, title="解析条件", slide_number=None, message=None, value_x_frac=0.36,
                         size_pt=24, label_frac=None):
    """
    型 K「解析条件」— 参照デッキの作り: テキストボックス 1 つ（上寄せ）に 24pt、
    ラベル = 黒・下線、値 = タブ位置（左から value_x_frac）にそろえる。グループの間は空行。赤は値の中の 1 語だけ。
    参照デッキの条件スライドには下部の結論行がない（message は任意）。
    items の要素:
      None                                  … 空行（グループ区切り）
      ("流れ場", "二次元，定常，非圧縮性")       … 1 行
      ("流れ場", [("二次元，", BLACK), ("定常", RED), ("，非圧縮性", BLACK)])  … 値の一部だけ色
      ("支配方程式", ["1 行目", "2 行目"])        … 値が複数行（2 行目以降も値の列にそろう）
      ("評価器", [[("NF + XFOIL", RED)], "2 行目"])  … 複数行で一部だけ色
      ("離散化手法", None, [("圧力項", "二次精度中心差分"), ...])  … 子項目（全角 1 字下げ）
    """
    from pptx.oxml.ns import qn as _qn
    slide, _ = add_content_slide(prs, title, slide_number)
    if label_frac is not None:            # 旧引数の互換
        value_x_frac = label_frac
    l = CT_BODY_L + 100000
    w = CT_BODY_W - 100000
    bottom = CONCL_T - 60000 if message else CT_BODY_T + CT_BODY_H
    box = slide.shapes.add_textbox(Emu(l), Emu(CT_BODY_T + 150000), Emu(w), Emu(bottom - CT_BODY_T - 150000))
    tf = box.text_frame
    tf.word_wrap = True
    tab_pos = int(w * value_x_frac)
    first = [True]

    def para(gap=0):
        p_ = tf.paragraphs[0] if first[0] else tf.add_paragraph()
        first[0] = False
        p_.space_before = Pt(gap)          # 項目の頭だけ空ける（参照デッキ は項目間に約半行、子項目・2 行目は詰める）
        pPr = p_._p.get_or_add_pPr()
        tl = etree.SubElement(pPr, _qn('a:tabLst'))
        t_ = etree.SubElement(tl, _qn('a:tab'))
        t_.set('pos', str(tab_pos))
        t_.set('algn', 'l')
        return p_

    def put_value(p_, val):
        runs = [(val, BLACK)] if isinstance(val, str) else val
        for r_ in runs:
            text, color = r_[0], r_[1]
            styled_run(p_, text, color=color, bold=False, size_pt=size_pt)

    for it in items:
        if it is None:
            p_ = para()
            styled_run(p_, " ", size_pt=size_pt * 0.4)
            continue
        label, val = it[0], it[1]
        extra = it[2] if len(it) > 2 else None
        p_ = para(gap=size_pt * 0.6)
        r_ = styled_run(p_, label, size_pt=size_pt)
        r_.font.underline = True
        if val is not None:
            # 複数行 = 要素が str か run のリスト（例 [[("NF + XFOIL", RED)], "2 行目"]）。単一行の run 列は tuple のリスト
            lines = val if (isinstance(val, list) and val and isinstance(val[0], (str, list))) else [val]
            if extra is True:                         # 旧形式: 値全体を赤太字
                lines = [[(lines[0], RED)]] + lines[1:]
            for k, v in enumerate(lines):
                if k > 0:
                    p_ = para()
                styled_run(p_, "\t", size_pt=size_pt)
                put_value(p_, v)
        if isinstance(extra, list):
            for sub, sv in extra:
                q = para()
                styled_run(q, "　" + sub, size_pt=size_pt)
                styled_run(q, "\t", size_pt=size_pt)
                put_value(q, sv)
    if message:
        add_conclusion(slide, message)
    return slide


def add_issue_table_slide(prs, rows, this_time=(0,), title="Current Issues", slide_number=None,
                          message=None, headers=("No.", "Issue", "Detail")):
    """
    型 I「課題一覧（This time 行）」: 表（No / 課題 / 詳細）。今回扱う行を赤破線で囲み、左上に赤「This time」。
    前回の指摘・残課題を棚卸しして、以降のスライドで 1 行ずつ回答する。
    rows: [("課題", "詳細"), ...]、this_time: 今回扱う行 index（連続を想定）。
    """
    slide, _ = add_content_slide(prs, title, slide_number)
    bottom = CONCL_T - 80000 if message else CT_BODY_T + CT_BODY_H
    t0 = FIG_AREA_T + 480000
    n = len(rows) + 1
    rh = min(700000, (bottom - t0) // n)
    tb = slide.shapes.add_table(n, 3, Emu(CT_BODY_L), Emu(t0), Emu(CT_BODY_W), Emu(rh * n)).table
    for j, cw in enumerate((0.08, 0.37, 0.55)):
        tb.columns[j].width = Emu(int(CT_BODY_W * cw))
    for j, hd in enumerate(headers):
        c = tb.cell(0, j)
        c.fill.solid()
        c.fill.fore_color.rgb = TABLE_HDR
        r = c.text_frame.paragraphs[0].add_run()
        r.text = hd
        r.font.size, r.font.bold, r.font.color.rgb = Pt(22), True, WHITE
        c.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    for i, (issue, detail) in enumerate(rows):
        for j, v in enumerate((str(i + 1), issue, detail)):
            c = tb.cell(i + 1, j)
            c.fill.solid()
            c.fill.fore_color.rgb = TABLE_ROW1 if i % 2 == 0 else TABLE_ROW2
            r = c.text_frame.paragraphs[0].add_run()
            r.text = v
            r.font.size = Pt(22 if j == 1 else 20)
            r.font.bold = (j == 1)
            r.font.color.rgb = BLACK
    if this_time:
        i0, i1 = min(this_time), max(this_time)
        add_highlight_box(slide, CT_BODY_L - 80000, t0 + rh * (i0 + 1) - 50000, CT_BODY_W + 160000,
                          rh * (i1 - i0 + 1) + 100000, label=None)
        _txbox(slide, "This time", CT_BODY_L, FIG_AREA_T, 3000000, 450000, size_pt=24, bold=True, color=RED)
    if message:
        add_conclusion(slide, message)
    return slide


def add_before_after_slides(prs, title, before, after, labels=("Previous method", "New method"),
                            slide_number=None, label_colors=(BLUE, RED), message_colors=(BLACK, RED)):
    """
    型 BA「前回 → 今回」2 枚組: 同じレイアウトで左上ラベル（前 = 青、後 = 赤、下線 28pt）・図・下部の赤メッセージだけ変える。
    before/after = (画像パス, メッセージ)。
    """
    out = []
    for (img, msg), lab, col, mcol in ((before, labels[0], label_colors[0], message_colors[0]),
                                       (after, labels[1], label_colors[1], message_colors[1])):
        slide, _ = add_content_slide(prs, title, slide_number)
        b = _txbox(slide, lab, CT_BODY_L, FIG_AREA_T, 5000000, 500000, size_pt=28, bold=True, color=col)
        b.text_frame.paragraphs[0].runs[0].font.underline = True
        add_picture_fit(slide, img, CT_BODY_L, FIG_AREA_T + 520000, CT_BODY_W, FIG_AREA_H - 520000)
        add_conclusion(slide, msg, color=mcol)   # 手法 A/B の比較では message_colors=(BLUE, RED)
        out.append(slide)
    return out


def add_grid_figure_slide(prs, title, images, col_labels=None, row_labels=None, highlight_row=None,
                          slide_number=None, message=None):
    """
    型 G「結果の格子比較」: images[row][col]。列 = 物理量、行 = 今回/先行研究 など。列の間に縦点線。
    highlight_row の行ラベルは赤太字（"this time"）。
    """
    slide, _ = add_content_slide(prs, title, slide_number)
    nr, nc = len(images), len(images[0])
    lab_w = 1500000 if row_labels else 0
    top = FIG_AREA_T + (480000 if col_labels else 0)
    bottom = (CONCL_T - 60000) if message else CT_BODY_T + CT_BODY_H
    cw = (CT_BODY_W - lab_w) // nc
    ch = (bottom - top) // nr
    for j in range(nc):
        x = CT_BODY_L + lab_w + j * cw
        if col_labels:
            _txbox(slide, col_labels[j], x, FIG_AREA_T, cw, 450000, size_pt=22, bold=True, align=PP_ALIGN.CENTER)
        if j > 0:
            _line(slide, x, FIG_AREA_T, x, bottom, color=BLACK, width_pt=1.25, dash='sysDot')
    for i in range(nr):
        y = top + i * ch
        if row_labels:
            hl = (i == highlight_row)
            _txbox(slide, row_labels[i], CT_BODY_L, y + ch // 2 - 300000, lab_w, 600000, size_pt=22,
                   bold=hl, color=RED if hl else BLACK)
        for j in range(nc):
            add_picture_fit(slide, images[i][j], CT_BODY_L + lab_w + j * cw + 60000, y + 40000, cw - 120000, ch - 80000)
    if message:
        add_conclusion(slide, message)
    return slide


def add_results_next_slide(prs, lead, results, next_items, problems=None, title="Conclusion", slide_number=None):
    """
    型 C'「結論（ゼミ版 Results / Next）」: リード 1 行 → 赤枠内に Results（箇条）→ 枠外に Problems・Next（番号）。
    
    """
    from pptx.enum.text import MSO_ANCHOR
    slide, _ = add_content_slide(prs, title, slide_number)
    _txbox(slide, lead, CT_BODY_L, FIG_AREA_T, CT_BODY_W, 600000, size_pt=26, align=PP_ALIGN.CENTER)
    rt = FIG_AREA_T + 750000
    rh = 700000 + 560000 * len(results)
    fr = _rect(slide, CT_BODY_L, rt, CT_BODY_W, rh, line_color=RED, line_width_pt=2)
    tf = fr.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = styled_run(p, "Results", color=RED, bold=True, size_pt=26)
    r.font.underline = True
    _paras(tf, ["・" + x if isinstance(x, str) else x for x in results], size_pt=26, space_pt=6)
    box = slide.shapes.add_textbox(Emu(CT_BODY_L), Emu(rt + rh + 120000), Emu(CT_BODY_W),
                                   Emu(CT_BODY_T + CT_BODY_H - rt - rh - 120000))
    tf = box.text_frame
    tf.word_wrap = True
    for head, items in (("Problems", problems), ("Next", next_items)):
        if not items:
            continue
        p = tf.paragraphs[0] if not tf.paragraphs[0].runs else tf.add_paragraph()
        r = styled_run(p, head, color=BLUE, bold=True, size_pt=26)
        r.font.underline = True
        _paras(tf, [f"{k + 1}. {x}" for k, x in enumerate(items)], size_pt=24, space_pt=4)
    return slide


def add_chapter_plan_slide(prs, chapters, title="修士論文の構成（予定）", slide_number=None):
    """
    型 TH「修論の章立て」: 2 段組。章 = 24pt 下線、節 = 20pt 字下げ。chapters: [("第1章 緒論", ["1-1 研究背景", ...]), ...]
    """
    from pptx.enum.text import MSO_ANCHOR
    slide, _ = add_content_slide(prs, title, slide_number)
    half = (len(chapters) + 1) // 2
    cw = (CT_BODY_W - 200000) // 2
    for c, chunk in enumerate((chapters[:half], chapters[half:])):
        box = slide.shapes.add_textbox(Emu(CT_BODY_L + c * (cw + 200000)), Emu(FIG_AREA_T),
                                       Emu(cw), Emu(CT_BODY_H))
        tf = box.text_frame
        tf.word_wrap = True
        first = True
        for ch, secs in chunk:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.space_before = Pt(14)
            r = styled_run(p, ch, color=BLACK, bold=True, size_pt=26)
            r.font.underline = True
            for sname in secs:
                q = tf.add_paragraph()
                styled_run(q, "　" + sname, color=BLACK, size_pt=22)
    return slide


def add_schedule_grid_slide(prs, tasks, title="Schedule and Next plan", slide_number=None,
                            months=("April", "May", "June", "July", "August", "September",
                                    "October", "November", "December", "January", "February", "March"),
                            message=None):
    """
    型 S（参照デッキ 版・既定）: 1 年を上下 2 段（4–9 月 / 10–3 月）に分けた点線グリッド。見出しは月名テキスト（24pt）、
    各段の下に横線。tasks: [(label, start, end, color)]、start/end は 0–12 の月 index（0 = 4 月頭、小数可）。
    段をまたぐタスクは自動で 2 本に分ける。color = BLUE（研究タスク）/ RED（修論執筆・発表など）/ BLACK。
    
    """
    slide, _ = add_content_slide(prs, title, slide_number)
    bottom = CONCL_T - 60000 if message else CT_BODY_T + CT_BODY_H
    half = (bottom - FIG_AREA_T) // 2
    cw = CT_BODY_W // 6
    rows = [[] for _ in range(2)]
    for lab, a, b, col in tasks:
        first = True
        for r in range(2):
            lo, hi = max(a, 6 * r), min(b, 6 * (r + 1))
            if hi > lo or (a == b and 6 * r <= a < 6 * (r + 1)):
                rows[r].append((lab if first else "", lo - 6 * r, hi - 6 * r, col))   # 段をまたぐときラベルは 1 回
                first = False
    for r in range(2):
        top = FIG_AREA_T + r * half
        for j in range(6):
            _txbox(slide, months[6 * r + j], CT_BODY_L + j * cw, top, cw, 450000, size_pt=22, align=PP_ALIGN.CENTER)
            if j > 0:
                _line(slide, CT_BODY_L + j * cw, top, CT_BODY_L + j * cw, top + half - 60000, color=BLACK,
                      width_pt=1, dash='sysDot')
        _line(slide, CT_BODY_L, top + 470000, CT_BODY_L + CT_BODY_W, top + 470000, color=BLACK, width_pt=1, dash='dash')
        n = max(len(rows[r]), 1)
        rh = (half - 560000) // n          # 1 段 3 本までが読みやすい（参照デッキ も 2〜3 本）
        for k, (lab, a, b, col) in enumerate(rows[r]):
            y = top + 560000 + k * rh + rh // 2
            x0, x1 = CT_BODY_L + int(a * cw), CT_BODY_L + int(b * cw)
            if x1 - x0 < 80000:
                x0, x1 = x0 - 60000, x0 + 60000
            c = _line(slide, x0, y, x1, y, color=col, width_pt=2.5, arrow_end=True)
            ln = c.line._get_or_add_ln()
            h = etree.SubElement(ln, qn('a:headEnd'))
            h.set('type', 'triangle')
            if not lab:
                continue
            lw = 5600000
            if x1 + 80000 + lw <= CT_BODY_L + CT_BODY_W:      # 矢印の右に置けるなら右
                _txbox(slide, lab, x1 + 80000, y - 230000, lw, 460000, size_pt=20, color=col)
            else:                                            # 入らなければ左に右揃え
                _txbox(slide, lab, x0 - 80000 - lw, y - 230000, lw, 460000, size_pt=20, color=col,
                       align=PP_ALIGN.RIGHT)
    if message:
        add_conclusion(slide, message)
    return slide


def add_cite(paragraph, n, size_pt=None):
    """本文中の出典番号を上付きで付ける。出典フッタも必ず置く（R6）。"""
    r = paragraph.add_run()
    r.text = f"[{n}]"
    if size_pt:
        r.font.size = Pt(size_pt)
    r.font._rPr.set('baseline', '30000')
    return r


def add_availability_table_slide(prs, title, rows, cols, cells, highlight_col=None, img_path=None,
                                 legend=(("✓✓", "data exist and have been verified"),
                                         ("✓", "data exist but have not been verified"),
                                         ("Not available", "data do not exist")),
                                 slide_number=None, source=None):
    """
    型 DA「データ有無表」: 上に対象の図（任意）、下に表（行 = 計測法・計算法、列 = 条件・測定面）。
    セルは "✓✓" / "✓" / "Not available"（赤太字）。今回使う列を赤破線で囲む。表の下に記号の凡例。
    
    """
    slide, _ = add_content_slide(prs, title, slide_number)
    top = FIG_AREA_T
    if img_path:
        ih = FIG_AREA_H * 36 // 100
        add_picture_fit(slide, img_path, CT_BODY_L, top, CT_BODY_W, ih)
        top += ih + 80000
    nr, nc = len(rows) + 1, len(cols) + 1
    leg_h = 420000 * len(legend)
    rh = min(560000, (CT_BODY_T + CT_BODY_H - top - leg_h - (300000 if source else 80000)) // nr)
    lw = CT_BODY_W * 26 // 100
    cw = (CT_BODY_W - lw) // (nc - 1)
    tb = slide.shapes.add_table(nr, nc, Emu(CT_BODY_L), Emu(top), Emu(CT_BODY_W), Emu(rh * nr)).table
    tb.columns[0].width = Emu(lw)
    for j in range(1, nc):
        tb.columns[j].width = Emu(cw)
    for j, hd in enumerate([""] + list(cols)):
        c = tb.cell(0, j)
        c.fill.solid()
        c.fill.fore_color.rgb = TABLE_HDR
        r = c.text_frame.paragraphs[0].add_run()
        r.text = hd
        r.font.size, r.font.bold, r.font.color.rgb = Pt(22), True, WHITE
        c.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    for i, rl in enumerate(rows):
        for j, v in enumerate([rl] + list(cells[i])):
            c = tb.cell(i + 1, j)
            c.fill.solid()
            c.fill.fore_color.rgb = TABLE_ROW1 if i % 2 == 0 else TABLE_ROW2
            p_ = c.text_frame.paragraphs[0]
            p_.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
            r = p_.add_run()
            r.text = v
            na = (v == "Not available")
            r.font.size, r.font.bold = Pt(22), na
            r.font.color.rgb = RED if na else BLACK
    if highlight_col is not None:
        x = CT_BODY_L + lw + highlight_col * cw
        add_highlight_box(slide, x - 40000, top - 80000, cw + 80000, rh * nr + 160000, label=None)
    y = top + rh * nr + 140000
    for sym, desc in legend:
        b_ = slide.shapes.add_textbox(Emu(CT_BODY_L + 300000), Emu(y), Emu(CT_BODY_W - 600000), Emu(400000))
        p_ = b_.text_frame.paragraphs[0]
        styled_run(p_, sym, color=RED if sym == "Not available" else BLACK, bold=True, size_pt=20)
        styled_run(p_, " : " + desc, size_pt=20)
        y += 400000
    if source:
        add_source(slide, source)
    return slide


def add_equation_slide(prs, title, latex, symbols, img_path=None, slide_number=None, message=None,
                       eq_pt=36, source=None):
    """
    型 EQ「数式＋記号表」: 左 = 図（任意）、右上 = 数式（matplotlib mathtext で画像化）、右下 = 「記号 : 意味 [単位]」。
    latex: "$Z(t) = \\frac{V(t) - \\bar{V}}{\\sigma}$" のような mathtext。symbols: [("V(t)", "output voltage [V]"), ...]
    
    """
    import tempfile
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    slide, _ = add_content_slide(prs, title, slide_number)
    bottom = CONCL_T - 60000 if message else CT_BODY_T + CT_BODY_H - (250000 if source else 0)
    rx = CT_BODY_L + (CT_BODY_W * 48 // 100 if img_path else 0)
    rw = CT_BODY_L + CT_BODY_W - rx
    if img_path:
        add_picture_fit(slide, img_path, CT_BODY_L, FIG_AREA_T, rx - CT_BODY_L - 150000, bottom - FIG_AREA_T)
    fig = plt.figure(figsize=(0.01, 0.01))
    fig.text(0, 0, latex, fontsize=eq_pt)
    f = tempfile.NamedTemporaryFile(suffix=".png", delete=False).name
    fig.savefig(f, dpi=300, bbox_inches="tight", pad_inches=0.05, transparent=True)
    plt.close(fig)
    eh = (bottom - FIG_AREA_T) * 38 // 100
    add_picture_fit(slide, f, rx, FIG_AREA_T, rw, eh)
    box = slide.shapes.add_textbox(Emu(rx + 200000), Emu(FIG_AREA_T + eh + 100000), Emu(rw - 200000),
                                   Emu(bottom - FIG_AREA_T - eh - 100000))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, (sym, desc) in enumerate(symbols):
        p_ = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p_.space_after = Pt(6)
        styled_run(p_, sym, bold=True, size_pt=24)
        styled_run(p_, " : " + desc, size_pt=24)
    if message:
        add_conclusion(slide, message)
    if source:
        add_source(slide, source)
    return slide


def set_done_shade(slide, l, t, w, h, alpha_pct=70):
    """型 S の変種: 済んだ期間に白の半透明矩形を重ねて淡くする。"""
    shp = _rect(slide, l, t, w, h, fill_color=WHITE)
    sf = shp._element.spPr.find(qn('a:solidFill'))
    clr = sf.find(qn('a:srgbClr'))
    a = etree.SubElement(clr, qn('a:alpha'))
    a.set('val', str(int(alpha_pct * 1000)))
    return shp


# ─── 英語・論文紹介（引き継ぎ型 / SLIDE_RULES.md §2b） ──────────────────────────

def set_notes(slide, script, qa=None):
    """発表者ノート: 日本語原稿＋「--- 想定質問と回答 ---」＋ Q/A。qa: [(質問, 回答), ...]"""
    text = script.strip()
    if qa:
        text += "\n\n--- 想定質問と回答 ---\n" + "\n".join(f"Q{i + 1}. {q}\nA{i + 1}. {a}" for i, (q, a) in enumerate(qa))
    slide.notes_slide.notes_text_frame.text = text
    return slide


def add_abstract_slide(prs, journal, title, authors, abstract, slide_number=None,
                       slide_title="Article title and abstract", abstract_pt=20):
    """
    型 AB「題目と Abstract」: 上に誌名（中央 20pt）・題目（2 行 28pt）・著者（右寄せ 20pt）、
    下 約 60% に Abstract 全文（両端揃え）。abstract_pt < 20 は R1 違反 → ユーザー確認のうえで。
    
    """
    from pptx.enum.text import MSO_ANCHOR
    slide, _ = add_content_slide(prs, slide_title, slide_number)
    _txbox(slide, journal, CT_BODY_L, CT_BODY_T - 20000, CT_BODY_W, 420000, size_pt=20, align=PP_ALIGN.CENTER)
    _txbox(slide, title, CT_BODY_L, CT_BODY_T + 380000, CT_BODY_W, 1000000, size_pt=28, align=PP_ALIGN.CENTER)
    _txbox(slide, authors, CT_BODY_L, CT_BODY_T + 1380000, CT_BODY_W, 420000, size_pt=20, align=PP_ALIGN.RIGHT)
    y = CT_BODY_T + 1850000
    _txbox(slide, "Abstract", SLIDE_W // 2 - 1000000, y - 230000, 2000000, 420000, size_pt=20, align=PP_ALIGN.CENTER)
    _line(slide, CT_BODY_L, y, SLIDE_W // 2 - 1000000, y, color=BLACK, width_pt=1)
    _line(slide, SLIDE_W // 2 + 1000000, y, CT_BODY_L + CT_BODY_W, y, color=BLACK, width_pt=1)
    box = slide.shapes.add_textbox(Emu(CT_BODY_L), Emu(y + 120000), Emu(CT_BODY_W), Emu(SLIDE_H - y - 300000))
    box.name = "QUOTE"                     # 逐語引用（audit の caution 対象外）
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.JUSTIFY
    styled_run(p, abstract, size_pt=abstract_pt)
    return slide


def add_quote_figure_slide(prs, title, images, quotes, slide_number=None, labels=None, source=None,
                           text_h=None, quote_pt=20):
    """
    型 QF「図＋本文引用」: 図（複数なら横並び・(a)(b) ラベル）を上に大きく、下に「➢」付きの本文引用 1〜2 文。
    
    """
    from pptx.enum.text import MSO_ANCHOR
    slide, _ = add_content_slide(prs, title, slide_number)
    n_lines = sum(max(1, len(q) // 95 + 1) for q in quotes)
    th = text_h or (380000 * n_lines + 150000)
    bottom = SLIDE_H - (330000 if source else 120000)
    fig_h = bottom - th - CT_BODY_T - 60000
    n = len(images)
    gap = 150000
    cw = (CT_BODY_W - gap * (n - 1)) // n
    lab_h = 380000 if labels else 0
    for i, img in enumerate(images):
        x = CT_BODY_L + i * (cw + gap)
        add_picture_fit(slide, img, x, CT_BODY_T, cw, fig_h - lab_h)
        if labels:
            _txbox(slide, labels[i], x, CT_BODY_T + fig_h - lab_h, cw, lab_h, size_pt=20, align=PP_ALIGN.CENTER)
    box = slide.shapes.add_textbox(Emu(CT_BODY_L), Emu(bottom - th), Emu(CT_BODY_W), Emu(th))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, q in enumerate(quotes):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.JUSTIFY
        styled_run(p, "➢ ", color=BLUE, bold=True, size_pt=quote_pt)
        styled_run(p, q, size_pt=quote_pt)
    if source:
        add_source(slide, source)
    return slide


def add_keypoints_slide(prs, points, title="Conclusion", slide_number=None, message=None):
    """
    型 KP「Conclusion（➢ 見出し＋説明）」: points = [(太字の見出し, 1 行説明), ...]（3〜4 個）。ブロックを縦中央。
    
    """
    from pptx.enum.text import MSO_ANCHOR
    slide, _ = add_content_slide(prs, title, slide_number)
    bottom = CONCL_T - 60000 if message else CT_BODY_T + CT_BODY_H
    box = slide.shapes.add_textbox(Emu(CT_BODY_L + 500000), Emu(FIG_AREA_T), Emu(CT_BODY_W - 1000000), Emu(bottom - FIG_AREA_T))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    first = True
    for head, desc in points:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_before = Pt(10)
        styled_run(p, "➢ ", color=BLUE, bold=True, size_pt=26)
        styled_run(p, head, bold=True, size_pt=26)
        q = tf.add_paragraph()
        styled_run(q, "     " + desc, size_pt=22)
    if message:
        add_conclusion(slide, message)
    return slide


def add_principle_slide(prs, title, img_path, cause, effect, message, slide_number=None, source=None, img_frac=0.44):
    """
    型 PR「原理・仕組み」: 上に模式図、中段に「原因 ▼ 結果」の 2 文（28pt）、下に赤太字のメッセージ（枠なし・add_conclusion）。
    計測原理・物理現象・解析手法の前提など、結果を読むのに要る仕組みを 1 枚で説明する。
    
    """
    slide, _ = add_content_slide(prs, title, slide_number)
    img_h = int(FIG_AREA_H * img_frac)      # 既定 0.44。原因・結果が 1 行ずつなら 0.56 まで広げられる
    add_picture_fit(slide, img_path, CT_BODY_L, FIG_AREA_T, CT_BODY_W, img_h)
    txt_h, tri_h = (900000 if img_frac <= 0.44 else 520000), 240000   # 28pt で 2 行 / 1 行の高さ
    y = FIG_AREA_T + img_h + 40000
    for k, text in enumerate((cause, effect)):
        b_ = _txbox(slide, text, CT_BODY_L, y, CT_BODY_W, txt_h, size_pt=28, align=PP_ALIGN.CENTER)
        b_.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        y += txt_h
        if k == 0:
            tri = slide.shapes.add_shape(7, Emu(SLIDE_W // 2 - 220000), Emu(y + 20000), Emu(440000), Emu(tri_h))
            tri.rotation = 180
            tri.fill.solid()
            tri.fill.fore_color.rgb = NAVY
            tri.line.fill.background()
            y += tri_h + 40000
    # 最下段のメッセージは枠で囲まず赤太字（全体ルール R9・10/29 ユーザー指示）
    add_conclusion(slide, message)
    if source:
        add_source(slide, source)
    return slide


def add_next_plan_slide(prs, lead, items=None, img_path=None, steps=None, working=None,
                        working_text="Now I'm working on", title="Next plan", slide_number=None, source=None,
                        loop=False):
    """
    型 NP「Next plan（次回までにやること）」— 毎回のゼミの既定。上に枠なしの 1 文 32pt（参照デッキ は 36pt）、下は次のどれか:
      items   : 番号付きの作業 1〜3 個（26pt）
      img_path: 次に使う手法の説明図 1 枚
      steps + working=(i0, i1): 左に全体の流れ、今取り組んでいる段を赤破線で囲み、右に working_text（赤太字）と items
    「Now I'm working on」の流れ図版あり。案の比較（型 P）は方針が分岐した回だけ使う。
    """
    from pptx.enum.text import MSO_ANCHOR
    slide, _ = add_content_slide(prs, title, slide_number)
    lb = _txbox(slide, lead, CT_BODY_L, FIG_AREA_T + 40000, CT_BODY_W, 800000, size_pt=32, align=PP_ALIGN.CENTER)
    lb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    top = FIG_AREA_T + 1050000
    bottom = CT_BODY_T + CT_BODY_H - (250000 if source else 0)
    if steps:
        lw = CT_BODY_W * 45 // 100
        boxes = add_flow(slide, steps, CT_BODY_L, top, lw, bottom - top, direction='down', size_pt=24)
        if loop and len(boxes) > 1:   # 最適化の反復などの戻り線
            fx = boxes[0][0] + boxes[0][2] + 160000
            fy0, fy1 = boxes[0][1] + boxes[0][3] // 2, boxes[-1][1] + boxes[-1][3] // 2
            _line(slide, boxes[-1][0] + boxes[-1][2], fy1, fx, fy1, color=BLACK, width_pt=1.25)
            _line(slide, fx, fy1, fx, fy0, color=BLACK, width_pt=1.25)
            _line(slide, fx, fy0, boxes[0][0] + boxes[0][2], fy0, color=BLACK, width_pt=1.25, arrow_end=True)
        if working:
            a, b = boxes[working[0]], boxes[working[1]]
            add_highlight_box(slide, a[0] - 100000, a[1] - 90000, a[2] + 200000, b[1] + b[3] - a[1] + 180000, label=None)
        rx = CT_BODY_L + lw + 500000
        box = slide.shapes.add_textbox(Emu(rx), Emu(top), Emu(CT_BODY_W - lw - 400000), Emu(bottom - top))
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        styled_run(tf.paragraphs[0], working_text, color=RED, bold=True, size_pt=28)
        if items:
            _paras(tf, [""] + [f"{k + 1}. {x}" if isinstance(x, str) else x for k, x in enumerate(items)], size_pt=24)
    elif img_path:
        add_picture_fit(slide, img_path, CT_BODY_L, top, CT_BODY_W, bottom - top)
    elif items:
        box = slide.shapes.add_textbox(Emu(CT_BODY_L + 300000), Emu(top), Emu(CT_BODY_W - 600000), Emu(bottom - top))
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        for k, it in enumerate(items):
            p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
            p.space_after = Pt(18)
            styled_run(p, f"{k + 1}.  ", color=BLUE, bold=True, size_pt=28)
            runs = [(it, BLACK, False)] if isinstance(it, str) else it
            for text, color, bold in runs:
                styled_run(p, text, color=color, bold=bold, size_pt=28)
    if source:
        add_source(slide, source)
    return slide


ALLOWED_COLORS = {RGBColor(*c) for c in [
    (0, 0, 0), (0xFF, 0, 0), (0, 0, 0xFF), (0xFF, 0xFF, 0xFF), (0x00, 0x20, 0x60), (0x76, 0x76, 0x76)]}


# 禁句（本文・表・ノートのすべてで違反）。研究上の誤記につながる語。追加は Pptx/RESEARCH_DOMAIN.md §2 と同時に
FORBIDDEN_TERMS = [
    (r"fully[ -]turbulent", "SA は「no transition model / 遷移モデルなし」と書く"),
    (r"完全乱流", "SA は「遷移モデルなし」と書く"),
]
# 誤解を招きやすい表現（注意 = warnings）。検定をしていない「有意」、根拠の書かれない最上級
CAUTION_TERMS = [
    (r"\bsignifican", "統計的に検定していないなら significant を使わない（differs by … と数値で）"),
    (r"有意", "検定していないなら「有意」を使わない"),
    (r"\b(always|never|proves?|proven)\b", "断定語。条件（n・範囲）を添えるか言い換える"),
]


def audit(prs, ref_name="REF", min_pt=20, repeat_ok=("Conclusion",), warnings=None):
    """
    SLIDE_RULES §5 の機械点検。違反を [(slide_no, kind, value, text)] で返す（0 件が合格）。
    warnings にリストを渡すと、目視で判断すべき注意（草稿の疑い・ノートの仮置き）をそこに追記する。
    kind: size（20pt 未満・REF 除く）/ color（許可色以外）/ meta（文書プロパティ・authors に個人情報が残る）/
          colorword（本文に色・装飾の説明）/ dup_title（タイトル重複）/ msg_len（結論 1 行が長い・改行あり）/
          cite_no参照デッキ（[n] があるのに出典フッタなし）/ draft（TODO・仮・空の箇条）/
          term（FORBIDDEN_TERMS の禁句。ノートも対象）
    warnings: caution（CAUTION_TERMS の表現。「➢」の逐語引用は除く）ほか
    """
    import re
    bad = []
    cp = prs.core_properties
    for k in ("author", "last_modified_by", "title"):
        if getattr(cp, k):
            bad.append((0, 'meta', k, getattr(cp, k)))
    if any(r.reltype.endswith(("/authors", "/commentAuthors")) for r in prs.part.rels.values()):
        bad.append((0, 'meta', 'authors.xml', 'present'))
    _ap = _app_part(prs)
    if _ap is not None and (b"TitlesOfParts" in _ap.blob or b"<Company>" in _ap.blob):
        bad.append((0, 'meta', 'app.xml', 'slide titles / company of the template deck'))
    color_word = re.compile(r"(（|\()(赤|青|黒|灰|破線|赤破線|太字)[^）)]*(）|\))|(赤|青)(字|枠|線|破線)|\b(red|blue|dashed)\s+(box|frame|line|text)\b", re.I)
    draft = re.compile(r"^\s*(TODO|XXX|\?\?\?|ここに|仮置き|未記入)")
    titles = []
    for si, s_ in enumerate(prs.slides, 1):
        has_ref, has_cite = False, False
        for shp in s_.shapes:
            frames = []
            if shp.has_text_frame:
                frames.append(shp.text_frame)
            if getattr(shp, 'has_table', False) and shp.has_table:
                frames += [c.text_frame for row in shp.table.rows for c in row.cells]
            if shp.name == ref_name:
                has_参照デッキ = True
            if shp.name == "TITLE":
                titles.append((si, shp.text_frame.text.strip()))
            if (shp.has_text_frame and shp.text_frame.text.strip() and shp.top is not None
                    and shp.top >= CONCL_T - 200000 and shp.name not in (ref_name, "MSG")):
                try:                                            # R12: 最下段の文字を枠で囲まない
                    boxed = shp.line.fill.type == 1
                except Exception:
                    boxed = False
                if boxed:
                    bad.append((si, 'boxed_bottom', shp.text_frame.text[:30]))
            if shp.name == "MSG":
                t = shp.text_frame.text
                if "\n" in t or len(t.encode("utf-8")) > 150:
                    bad.append((si, 'msg_len', len(t), t[:30]))
            for tf in frames:
                for p in tf.paragraphs:
                    ptxt = "".join(r.text for r in p.runs)
                    for pat, why in FORBIDDEN_TERMS:
                        if re.search(pat, ptxt, re.I):
                            bad.append((si, 'term', why, ptxt[:30]))
                    if (warnings is not None and shp.name not in (ref_name, "QUOTE")
                            and not ptxt.lstrip().startswith("➢")):         # 逐語引用（➢・Abstract）は除く
                        for pat, why in CAUTION_TERMS:
                            if re.search(pat, ptxt, re.I):
                                warnings.append((si, 'caution', why, ptxt[:30]))
                    if shp.name != ref_name:
                        if color_word.search(ptxt):
                            bad.append((si, 'colorword', '', ptxt[:30]))
                        if draft.search(ptxt):
                            bad.append((si, 'draft', '', ptxt[:30]))
                        if re.search(r"\[\d+\]", ptxt):
                            has_cite = True
                        if ptxt.strip() in ("•", "・", "➢", "1.", "2.", "3.") and warnings is not None:
                            warnings.append((si, 'empty_bullet', '', ptxt))
                    for r in p.runs:
                        if not r.text.strip():
                            continue
                        sz = r.font.size.pt if r.font.size else None
                        if shp.name != ref_name and (sz is None or sz < min_pt):
                            bad.append((si, 'size', sz, r.text[:30]))
                        try:
                            col = r.font.color.rgb
                        except AttributeError:
                            col = None
                        if col is not None and col not in ALLOWED_COLORS:
                            bad.append((si, 'color', str(col), r.text[:30]))
        if has_cite and not has_ref:
            bad.append((si, 'cite_noref', '', ''))
        if s_.has_notes_slide:
            for pat, why in FORBIDDEN_TERMS:
                if re.search(pat, s_.notes_slide.notes_text_frame.text, re.I):
                    bad.append((si, 'term', why + "（ノート）", ''))
        if warnings is not None and s_.has_notes_slide:
            nt = s_.notes_slide.notes_text_frame.text
            if re.search(r"(仮置き|試作|TODO|あとで直す|後で直す|要修正)", nt):   # 「測り直す」は誤検出なので除く
                warnings.append((si, 'note_draft', '', nt[:30]))
    from collections import Counter
    cnt = Counter(t for _, t in titles)
    for k, (si, t) in enumerate(titles):
        if cnt[t] < 2 or t in repeat_ok or t.endswith("structure"):
            continue
        prev_same = k > 0 and titles[k - 1][1] == t and titles[k - 1][0] == si - 1
        next_same = k + 1 < len(titles) and titles[k + 1][1] == t and titles[k + 1][0] == si + 1
        if not (prev_same or next_same):          # 連続スライド（ビルドアップ・前回→今回）は許可
            bad.append((si, 'dup_title', cnt[t], t[:30]))
    return bad


def add_source(slide, text, size_pt=12):
    """出典フッタ（R1 の唯一の例外）。audit() は名前 'REF' で除外する。"""
    box = _txbox(slide, text, REF_L, REF_T - 60000, REF_W, REF_H + 60000, size_pt=size_pt)
    box.name = "REF"
    return box


# ─── 変数のイタリック（R11）──────────────────────────────────────────────────
# 物理量・記号はイタリック、添字が語（crit・max）なら立体。数字・単位・関数名は立体。
# (トークン, [(文字, 斜体?, 下付き?)])。長いものを先に照合する。
VAR_TOKENS = [
    ("Ncrit", [("N", True, False), ("crit", False, True)]),
    ("t_max", [("t", True, False), ("max", False, True)]),
    # 下付き l と / が重なるので / の前に細いスペース（U+2009）を入れる（10/29 ユーザー指摘）
    ("Cl/Cd", [("C", True, False), ("l", True, True), ("\u2009/", False, False), ("C", True, False), ("d", True, True)]),
    ("Cdp", [("C", True, False), ("dp", True, True)]),
    ("Cdf", [("C", True, False), ("df", True, True)]),
    ("Cl", [("C", True, False), ("l", True, True)]),
    ("Cd", [("C", True, False), ("d", True, True)]),
    ("Cp", [("C", True, False), ("p", True, True)]),
    ("Cm", [("C", True, False), ("m", True, True)]),
    ("Re", [("Re", True, False)]),
    ("Ma", [("Ma", True, False)]),
    ("t/c", [("t", True, False), ("/", False, False), ("c", True, False)]),
    ("x/c", [("x", True, False), ("/", False, False), ("c", True, False)]),
    ("y/c", [("y", True, False), ("/", False, False), ("c", True, False)]),
    ("y+", [("y", True, False), ("+", False, False)]),
    ("α", [("α", True, False)]),
    ("ν", [("ν", True, False)]),
    ("σ", [("σ", True, False)]),
]


def _var_regex():
    import re
    alts = "|".join(re.escape(t) for t, _ in VAR_TOKENS)
    # 前後が英字・数字・ハイフンでない（Clerc・Remaining・Re-run・Max は対象外）
    return re.compile(r"(?<![A-Za-z0-9])(" + alts + r")(?![A-Za-z0-9\-])")


def italicize_variables(prs, skip_names=()):
    """
    全スライド（表を含む）の文字を走査し、VAR_TOKENS の変数をイタリック（必要なら下付き添字）に分割する。
    元の run の書式（色・太字・サイズ）は引き継ぐ。数式画像（EQ 型）と図の中は対象外（図は mathtext で書く）。
    save() が自動で呼ぶ。
    """
    import copy as _copy
    rx = _var_regex()
    spec = dict(VAR_TOKENS)

    def split_run(r):
        text = r.text
        if not rx.search(text):
            return
        pieces = []                     # (文字, italic or None, subscript)
        pos = 0
        for m in rx.finditer(text):
            if m.start() > pos:
                pieces.append((text[pos:m.start()], None, False))
            for t, it, sub in spec[m.group(1)]:
                pieces.append((t, it, sub))
            pos = m.end()
        if pos < len(text):
            pieces.append((text[pos:], None, False))
        base = r._r
        pristine = _copy.deepcopy(base)          # 書き換え前の書式から複製する
        prev = base
        for k, (t, it, sub) in enumerate(pieces):
            nr = base if k == 0 else _copy.deepcopy(pristine)
            if k > 0:
                prev.addnext(nr)
            prev = nr
            tnode = nr.find(qn('a:t'))
            tnode.text = t
            rPr = nr.get_or_add_rPr()
            if it is not None:
                rPr.set('i', '1' if it else '0')
            if sub:
                rPr.set('baseline', '-25000')

    def frames(shapes):
        for shp in shapes:
            if shp.name in skip_names:
                continue
            if shp.shape_type == 6:          # group
                yield from frames(shp.shapes)
                continue
            if shp.has_text_frame:
                yield shp.text_frame
            if getattr(shp, 'has_table', False) and shp.has_table:
                for row in shp.table.rows:
                    for c in row.cells:
                        yield c.text_frame

    from pptx.text.text import _Run
    for slide in prs.slides:
        for tf in frames(slide.shapes):
            for p in tf.paragraphs:
                for r in list(p.runs):
                    split_run(r)


# ─── 出典フッタ（R6・R6'）────────────────────────────────────────────────────
_CITE = re.compile(r"\[\d+\]")


def keep_citation_footers(prs):
    """下端のフッタ（名前 'REF'）を論文・資料の引用だけにする（R6: 自分のデータの注記はノートへ）。
    - [n] を含む: 先頭の条件文（Re = … など）・「Cropped/Adapted from Fig. n of」・「(CC BY x.x)」・後ろの
      「This study: / Data: / Values from / Est.:」以降を落とす（R6'）。
    - [n] を含まない: 「Exp.: …」/「Experiment: …」（実験値の出典）だけ「Experiment: …」で残し、他は消す。
    戻り値: (残した数, 消した数)。save() が既定で呼ぶ。"""
    kept = removed = 0
    for sl in prs.slides:
        for sh in list(sl.shapes):
            if sh.name != "REF" or not sh.has_text_frame:
                continue
            t0 = sh.text_frame.text
            if _CITE.search(t0):
                t = t0
                if not re.match(r"^(Left:\s*)?(Cropped|Adapted|Schematic|\[)", t, re.I):
                    t = t[_CITE.search(t).start():]
                t = re.sub(r"^(Left:\s*)?(Cropped|Adapted) from .*? of (?=\[\d+\])", "", t, flags=re.I)
                t = re.sub(r"\s*\(CC BY[^)]*\)", "", t)
                t = re.split(r"\s(?:This study:|Data:|Values from|Est\.:)", t)[0].strip()
            else:
                m = re.search(r"\bExp(?:eriment)?\.?:\s*([^.;]+)", t0)
                if not m:
                    sh._element.getparent().remove(sh._element)
                    removed += 1
                    continue
                t = f"Experiment: {m.group(1).strip()}."
            kept += 1
            if t != t0:
                para = sh.text_frame.paragraphs[0]
                for r in para.runs[1:]:
                    r._r.getparent().remove(r._r)
                for extra in sh.text_frame.paragraphs[1:]:
                    extra._p.getparent().remove(extra._p)
                para.runs[0].text = t
    return kept, removed


def strip_section_footers(prs, start_title, end_titles=(), quotes=False, keep_quote_titles=("Conclusion",)):
    """start_title の扉から end_titles のどれかで始まるスライドの手前までの区間で、下端フッタ（REF）を消す。
    quotes=True なら「➢」で始まる引用文の枠も消す（keep_quote_titles のスライドは除く）。論文紹介を図だけにする用。"""
    on, n = False, 0
    for sl in prs.slides:
        t = next((x.text_frame.text for x in sl.shapes if x.has_text_frame), "")
        if t == start_title:
            on = True
        elif any(t.startswith(e) for e in end_titles):
            on = False
        if not on:
            continue
        for sh in list(sl.shapes):
            quote = quotes and t not in keep_quote_titles and sh.has_text_frame and sh.text_frame.text.lstrip().startswith("➢")
            if sh.name == "REF" or quote:
                sh._element.getparent().remove(sh._element)
                n += 1
    return n


def save(prs, out_path: str, italic_vars=True, coedit=True, coedit_init=False, citations_only=True):
    """プレゼンテーションを保存。ディレクトリが無ければ作成。
    citations_only=True（既定）: keep_citation_footers でフッタを引用だけにする（R6）。
    coedit=True（既定）: deck_merge.save_coedit で共同編集つき保存（PowerPoint で開いていたら止まる・
    既存ファイルを output/bg/ に時刻付きでコピー・手編集があれば 3-way マージ）。仕組み = deck_merge.py の冒頭。"""
    import os
    _scrub_app_props(prs)
    if citations_only:
        keep_citation_footers(prs)
    if italic_vars:
        italicize_variables(prs)
    if not coedit:
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        prs.save(out_path)
        print(f"Saved: {out_path}")
        return None
    import sys as _sys
    _sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import deck_merge
    res = deck_merge.save_coedit(prs, out_path, init=coedit_init)
    print(f"Saved: {out_path}  [{res['mode']}]" + (f"  backup: {res['backup']}" if res["backup"] else ""))
    if res["report"]:
        print("  merge: " + res["report"][0] + f"  （詳細: {res['report_path']}）")
        for ln in res["report"][1:]:
            if ln.startswith(("[衝突]", "[注意]")):
                print("   ", ln)
    return res

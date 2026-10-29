"""スライド型カタログ: SLIDE_RULES.md §3 の各型を 1 枚ずつ生成し、再現性を確かめる。

実行:
        python3 scripts/make_slide_type_catalog.py
出力: output/slide_type_catalog.pptx（PDF 化は SKILL の手順）
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from slide_template import (  # noqa: E402
    new_presentation, save, audit, add_title_slide, add_section_slide, add_agenda_slide,
    add_objective_slide, add_conclusion_slide, add_plan_options_slide, add_schedule_slide,
    add_content_slide, add_flow, add_highlight_box, add_conclusion, add_source, add_closing_slide,
    add_table_slide, RED, BLUE, BLACK, CT_BODY_L, CT_BODY_W, FIG_AREA_T, FIG_AREA_H,
    add_buildup_slides, add_question_slide, add_compare_slide, add_method_nav_slide, add_step_tag,
    add_conditions_slide, add_issue_table_slide, add_before_after_slides, add_grid_figure_slide,
    add_results_next_slide, add_chapter_plan_slide, add_picture_fit, add_next_plan_slide, add_principle_slide,
    add_availability_table_slide, add_equation_slide,
)

FIG = str(Path(__file__).resolve().parent / "_fig")

OUT = str(Path(__file__).resolve().parent.parent / "output" / "slide_type_catalog.pptx")


def _make_placeholder_figs(d):
    """カタログ用の仮の図（fig5/fig7/fig8.png）。中身は意味のない曲線。"""
    import os
    import numpy as np
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    os.makedirs(d, exist_ok=True)
    x = np.linspace(0, 1, 200)
    for name, f in (("fig5", np.sin), ("fig7", np.cos), ("fig8", np.tanh)):
        fig, ax = plt.subplots(figsize=(8, 4.5))
        ax.plot(x, f(6 * x), color="#0000FF", lw=3)
        ax.plot(x, 0.8 * f(6 * x + 0.5), color="#FF0000", lw=3, ls="--")
        ax.set_xlabel("x", fontsize=20)
        ax.set_ylabel("y", fontsize=20)
        ax.tick_params(labelsize=18)
        fig.tight_layout()
        fig.savefig(os.path.join(d, name + ".png"), dpi=150)
        plt.close(fig)


def main():
    _make_placeholder_figs(FIG)
    prs = new_presentation()
    add_title_slide(prs, title="Slide Type Catalog\n(SLIDE_RULES.md §3)",
                    seminar_label="Master Seminar M2, 2026/10/13",
                    lab_name="Example Lab.", member_grade_name="M2",
                    supervisors=[],
                    subtitle="型 T: タイトル（subtitle = 論文紹介の書誌 1 行）")
    add_agenda_slide(prs, ["研究背景・目的", "解析手法", "解析結果および考察", "結論・今後の予定"], 0)
    add_section_slide(prs, "型 X: セクション区切り")
    add_objective_slide(prs, "翼型の空力性能を、不確かさを含めて評価できる最適化手法をつくる",
                        steps=["翼型の形状表現\nCST 18 変数", "空力評価\nXFOIL / OpenFOAM", "形状最適化\nPSO・勾配法"],
                        title="型 O: 目的", slide_number=1, highlight=(1, 1),
                        conclusion="今回は空力評価の信頼性を扱う")
    add_objective_slide(prs, "（横フロー版）direction='right'",
                        steps=["入力\n熱線データ", "機械学習\nモデル", "出力\n順流/逆流"],
                        title="型 O': 目的（横フロー）", slide_number=2, direction='right',
                        conclusion="This problem is called “time-series classification”")
    s, _ = add_content_slide(prs, "型 H: This time ハイライト付きフロー", slide_number=3)
    boxes = add_flow(s, ["Airfoil geometry", "Aerodynamic analysis\nOpenFOAM, XFOIL", "Shape optimization\nPSO or ML"],
                     CT_BODY_L + CT_BODY_W // 5, FIG_AREA_T + 100000, CT_BODY_W * 3 // 5, FIG_AREA_H - 200000)
    l, t, w, h = boxes[1]
    add_highlight_box(s, l - 150000, t - 120000, w + 300000, h + 240000, label="This time")
    add_conclusion(s, "スライドのメッセージは 1 行・赤で下に置く")
    add_source(s, "[1] 出典フッタ（12 pt・R1 の唯一の例外）")
    add_table_slide(prs, "型 TB: 表", ["", "E387", "Optimized"],
                    [["Cl", "0.743", "0.928"], ["Cd", "0.0208", "0.0202"], ["Cl/Cd", "35.7", "45.9"]],
                    slide_number=4)
    add_conclusion_slide(prs, "E387 を出発点に Re = 10⁵ で最適化し、以下の結果を得た",
                         ["Cl/Cd は XFOIL で +76%",
                          [("OpenFOAM（SA）では ", BLACK, False), ("+16% に縮む", RED, True)],
                          "格子化できない形状が 15/37 本"],
                         title="型 C: 結論", slide_number=5,
                         message="XFOIL の利得は CFD での確認とセットで評価する")
    add_next_plan_slide(prs, "格子化できない形状の原因を調べる",
                        items=["裏返るセルの位置を全 15 本で集計する", "前縁の曲率と格子品質の関係を調べる"],
                        title="型 NP: Next plan（既定）", slide_number=6)
    add_next_plan_slide(prs, "メッシュ生成と解析を自動化する",
                        steps=["Airfoil geometry", "Aerodynamic analysis\nOpenFOAM, XFOIL", "Shape optimization\nPSO"],
                        working=(1, 1), items=["格子生成の自動化", "収束判定の自動化"], loop=True,
                        title="型 NP': Next plan（Now I'm working on）", slide_number=6)
    add_principle_slide(prs, "型 PR: 原理・仕組み", FIG + "/fig5.png",
                        "低 Re では、層流境界層が逆圧力勾配に耐えられず剥離する",
                        "剥離した流れが乱流に遷移して再付着し、層流剥離泡ができる",
                        "XFOIL（eN 法）は剥離泡を解くが、SA（遷移モデルなし）は解かない", slide_number=6)
    add_plan_options_slide(prs, "修論までに取りうる方針は 3 つ",
                           [("Plan A", ["遷移モデル付き CFD", "↓", "最適形状の再評価"]),
                            ("Plan B", ["不確かさ定量化", "↓", "ロバスト最適化"]),
                            ("Plan C", ["風洞実験で検証"])],
                           selected=0, title="型 P: 案の比較（方針が分岐した回だけ）", slide_number=6,
                           message="Plan A から着手する")
    add_schedule_slide(prs, ["10月", "11月", "12月", "1月", "2月", "3月"],
                       [("遷移モデル付き CFD", 0.0, 1.8, 'task'),
                        ("最適化の再実行", 1.5, 3.0, 'task'),
                        ("修論執筆", 2.8, 4.6, 'task'),
                        ("修論提出", 4.6, 4.6, 'deadline'),
                        ("修論発表", 4.9, 4.9, 'deadline')],
                       title="型 S: スケジュール", slide_number=7, now=0.4,
                       message="12 月中に CFD の再評価を終える")
    # ── ref/M1/ から追加した型 ──
    add_buildup_slides(prs, "型 B: 背景のビルドアップ（3 枚連続）",
                       ["成層圏飛行船のプロペラは Re ≈ 10⁵ で動く",
                        "層流剥離泡ができ、抵抗が急増する"],
                       "層流剥離泡を抑える翼型設計が必要", slide_number=8)
    add_question_slide(prs, "型 Q: 問い → 図 → 呼称", "Is it possible to design low-Re airfoils automatically?",
                       FIG + "/fig7.png", "This problem is called “aerodynamic shape optimization”", slide_number=9)
    add_compare_slide(prs, "型 V: 二者比較", ("XFOIL", ["速い（0.1 s/評価）", "層流剥離泡を扱える", "高 α で非収束"]),
                      ("OpenFOAM（SA）", ["遅い（1 h/ケース）", "層流剥離泡を再現しない", "失速後も解ける"]),
                      bridge="使い分け", slide_number=10, message="最適化は XFOIL、検証は CFD で使い分ける")
    steps = ["(1) 形状表現\nCST 18 変数", "(2) 空力評価\nXFOIL / OpenFOAM", "(3) 形状最適化\nPSO"]
    add_method_nav_slide(prs, "2. 解析手法（型 N: 現在地ナビ）", steps, 1, FIG + "/fig8.png",
                         caption="翼面の圧力分布", slide_number=11)
    s, _ = add_content_slide(prs, "型 N の詳細スライド（ステップタグ付き）", slide_number=12)
    add_step_tag(s, "(2) 空力評価")
    add_picture_fit(s, FIG + "/fig5.png", CT_BODY_L, FIG_AREA_T + 600000, CT_BODY_W, FIG_AREA_H - 600000)
    add_conclusion(s, "詳細スライドでは上にステップタグを再掲する")
    add_conditions_slide(prs, [("流れ場", "二次元・非圧縮・定常", True), ("支配方程式", "RANS"),
                               ("乱流モデル", "Spalart–Allmaras"), ("解析ソフト", "OpenFOAM v2212"),
                               ("圧力–速度連成", "SIMPLE"), ("Re / α", "1.0 × 10⁵ / −2〜10°")],
                         title="型 K: 解析条件", slide_number=13, message="定常計算を選んだのは収束判定を揃えるため")
    add_issue_table_slide(prs, [("XFOIL の退化解", "パネル数で Cd が 3 割変わる形状がある"),
                                ("格子化の失敗", "前縁下面の凹みで格子が裏返る"),
                                ("CFD の層流剥離泡", "SA では剥離泡が出ない"),
                                ("最適化の未収束", "seed 間で Cl/Cd が 2〜4 ばらつく")],
                          this_time=(0, 1), title="型 I: 課題一覧（Current Issues）", slide_number=14,
                          message="今回は 1・2 を扱う")
    add_before_after_slides(prs, "型 BA: 前回 → 今回（2 枚組）",
                            (FIG + "/fig7.png", "前回: 形状だけを比較していた"),
                            (FIG + "/fig8.png", "今回: 圧力分布で剥離位置まで比較した"), slide_number=15)
    add_grid_figure_slide(prs, "型 G: 結果の格子比較",
                          [[FIG + "/fig7.png", FIG + "/fig8.png"], [FIG + "/fig7.png", FIG + "/fig8.png"]],
                          col_labels=["Shape", "Pressure distribution"], row_labels=["This time", "Wei 2022"],
                          highlight_row=0, slide_number=16, message="行 = 今回／先行研究、列 = 物理量")
    add_results_next_slide(prs, "This time, the CFD validation was redone and the following results were obtained.",
                           ["XFOIL gain (+76%) shrinks to +16% in OpenFOAM SA",
                            "15/37 shapes cannot be meshed"],
                           ["Transition-model CFD", "Re-run optimization with mesh constraint"],
                           problems=["SA does not capture the laminar separation bubble"],
                           title="型 C': 結論（ゼミ版 Results / Next）", slide_number=17)
    add_chapter_plan_slide(prs, [("第1章 緒論", ["1-1 研究背景", "1-2 研究目的"]),
                                 ("第2章 解析手法", ["2-1 形状表現", "2-2 空力評価", "2-3 最適化手法"]),
                                 ("第3章 結果および考察", ["3-1 最適化結果", "3-2 CFD による検証"]),
                                 ("第4章 結論", ["参考文献", "付録", "謝辞"])],
                           title="型 TH: 修論の章立て（予定）", slide_number=18)
    add_schedule_slide(prs, ["4月", "5月", "6月", "7月", "8月", "9月"],
                       [("文献調査", 0.0, 2.0, 'done'), ("XFOIL 最適化", 1.5, 4.0, 'done'),
                        ("CFD 検証", 3.5, 6.0, 'task'), ("中間発表", 4.8, 4.8, 'deadline')],
                       title="型 S': スケジュール（済んだ期間を淡く）", slide_number=19, now=4.2, done_until=4.2)
    add_section_slide(prs, "型 X': 番号付き小目次扉", items=["Data standardization", "Feature extraction"], current_idx=1)
    add_availability_table_slide(prs, "型 DA: The data currently available",
                                 ["XFOIL", "OpenFOAM (SA)", "Wind tunnel"], ["Re = 1×10⁵", "Re = 2×10⁵", "Re = 3×10⁵"],
                                 [["✓✓", "✓", "✓✓"], ["✓✓", "Not available", "✓"], ["✓✓", "✓✓", "✓✓"]],
                                 highlight_col=0, img_path=FIG + "/fig7.png", slide_number=20)
    add_equation_slide(prs, "型 EQ: Data standardization", r"$Z(t) = \dfrac{V(t) - \bar{V}}{\sigma}$",
                       [("V(t)", "output voltage [V]"), ("V̄", "time-averaged voltage [V]"), ("σ", "standard deviation [V]")],
                       img_path=FIG + "/fig8.png", slide_number=21,
                       message="Standardization removes the velocity dependence.")
    add_section_slide(prs, "Introduce an article")
    add_closing_slide(prs)
    add_section_slide(prs, "Appendix")

    bad = audit(prs)
    save(prs, OUT)
    print("slides:", len(prs.slides), "audit violations:", len(bad))
    for b in bad:
        print("  ", b)


if __name__ == "__main__":
    main()

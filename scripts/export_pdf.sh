#!/usr/bin/env bash
# PPTX → PDF（Windows の PowerPoint COM）。開いている VS Code の PDF タブが自動で更新されるよう、
# 既存 PDF を削除せず「一時ファイルに保存 → mv で置き換え」にする（削除するとビューアがファイルを見失う）。
# usage: bash scripts/export_pdf.sh output/<name>.pptx [out.pdf]   （WSL + Windows の PowerPoint が必要）
set -o pipefail
PPTX=$(realpath "$1")
LOCK="$(dirname "$PPTX")/~\$$(basename "$PPTX")"
if [ -e "$LOCK" ]; then   # PowerPoint で開いている間は書き出さない（開いているウィンドウと干渉する）
  echo "PowerPoint で開いています（$(basename "$LOCK")）。閉じてから再実行してください" >&2; exit 2
fi
PDF=${2:-${PPTX%.pptx}.pdf}
TMP="${PDF%.pdf}.__tmp__.pdf"
# PowerPoint が同じファイルを（別ウィンドウ・記憶上の古い版で）開いていると、COM の Open はその古い版を返して
# 古い内容の PDF になる（2026/10/28 実測）。毎回、一意な名前の一時コピーを開いて書き出す
SRC="$(dirname "$PPTX")/.__export_$$_$(date +%s)__.pptx"
cp -f "$PPTX" "$SRC"
trap 'rm -f "$SRC"' EXIT
rm -f "$TMP"   # 一時ファイルだけ消す（既存の上書き失敗を避ける）。本体の PDF は消さない
PS=/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe
"$PS" -Command "
\$ppt = New-Object -ComObject PowerPoint.Application
\$wb = \$ppt.Presentations.Open('$(wslpath -w "$SRC")', \$true, \$false, \$false)
\$wb.SaveAs('$(wslpath -w "$TMP")', 32)
\$wb.Close()" || { echo "export failed: $PPTX" >&2; exit 1; }
[ -s "$TMP" ] || { echo "no output: $TMP" >&2; exit 1; }
mv -f "$TMP" "$PDF"
echo "wrote $PDF"

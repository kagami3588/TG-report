"""TG訪問営業チーム 月次実績レポート(サマリー)を生成する。

入力: 個人ごとの 今月/先月 の 訪問・IF越し・対面・客宅・アポ獲得・キャンセル
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L

MONTH, PREV = "9月", "8月"
TITLE = "TG訪問営業チーム　月次実績　2026年9月（9/1〜9/29）"
PERIOD = "2026年9月1日〜30日"
TARGET = 13

# 名前: (今月[訪問,IF越し,対面,客宅,アポ,キャンセル], 先月[...])
DATA = [
    ("水取和也",   (0, 0, 0, 0, 0, 0),              (93, 45, 23, 5, 3, 0)),
    ("越智大介",   (0, 0, 0, 0, 0, 0),              (311, 87, 64, 5, 1, 0)),
    ("大川涼真",   (1049, 468, 242, 7, 3, 0),       (860, 375, 200, 13, 8, 0)),
    ("上運天寛太", (0, 0, 0, 0, 0, 0),              (474, 141, 100, 1, 1, 0)),
    ("葛西杏樹",   (1011, 571, 303, 14, 8, 0),      (1118, 602, 300, 15, 6, 0)),
    ("細川麻里菜", (1220, 518, 229, 15, 5, 0),      (1005, 439, 203, 9, 4, 0)),
    ("野夫井信哉", (1416, 520, 228, 9, 3, 0),       (1169, 383, 179, 6, 2, 0)),
    ("柴田和希",   (1131, 476, 198, 5, 2, 0),       (1389, 612, 252, 5, 2, 0)),
    ("上ノ堀聡士", (1448, 586, 255, 7, 2, 0),       (1362, 472, 218, 9, 2, 0)),
]

NAVY = "1F3864"
hdr_fill = PatternFill("solid", fgColor=NAVY)
sec_fill = PatternFill("solid", fgColor="D9E2F3")
sub_fill = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
CNT = '#,##0"件"'
PCT = '0.0%'
DCNT = '+#,##0"件";-#,##0"件";"±0件"'
DPT = '+0.0"pt";-0.0"pt";"±0pt"'

wb = Workbook()

# ---------------- 入力シート ----------------
wi = wb.active
wi.title = "入力"
wi["A1"] = "月次入力（数字はここだけ更新）"
wi["A1"].font = Font(bold=True, size=14)
wi["A2"], wi["B2"] = "月間目標(件/人)", TARGET
wi["A3"], wi["B3"] = "データ取得期間", PERIOD
heads = ["訪問", "IF越し", "対面", "客宅・会社説明", "アポ獲得", "キャンセル"]
wi["B5"], wi["H5"] = f"{MONTH}（今月）", f"{PREV}（先月）"
for c in ("B5", "H5"):
    wi[c].font = Font(bold=True)
wi["A6"] = "名前"
for i, h in enumerate(heads):
    wi.cell(6, 2 + i, h)
    wi.cell(6, 8 + i, h)
for c in range(1, 14):
    if c == 7:
        continue
    cell = wi.cell(6, c)
    cell.fill, cell.font = hdr_fill, Font(bold=True, color="FFFFFF")
    cell.alignment = Alignment(horizontal="center")
R0 = 7  # first data row
for r, (name, cur, prv) in enumerate(DATA, start=R0):
    wi.cell(r, 1, name)
    for i, v in enumerate(cur):
        wi.cell(r, 2 + i, v)
    for i, v in enumerate(prv):
        wi.cell(r, 8 + i, v)
R1 = R0 + len(DATA) - 1
wi.column_dimensions["A"].width = 16
for c in range(2, 14):
    wi.column_dimensions[L(c)].width = 12
wi.column_dimensions["G"].width = 3
wi["A17"] = "※ 今月の訪問が0の人は「未稼働」として個人表から除き、先月チーム合計も稼働者のみで比較する。"

rng = lambda col: f"入力!${col}${R0}:${col}${R1}"
# 今月 col: B訪問 C IF D対面 E客宅 F アポ G cancel / 先月 H..M
CUR = dict(v="B", i="C", f="D", k="E", a="F", x="G")
PRV = dict(v="H", i="I", f="J", k="K", a="L", x="M")

def cur_sum(key):
    return f"SUM({rng(CUR[key])})"

def prv_sum(key):  # 今月稼働者のみ
    return f"SUMIF({rng('B')},\">0\",{rng(PRV[key])})"

# ---------------- サマリー ----------------
ws = wb.create_sheet("サマリー（印刷用）", 0)
ws.sheet_view.showGridLines = False
widths = [16] + [11] * 18
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[L(i)].width = w

def merge(r, c1, c2, v=None, **kw):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    cell = ws.cell(r, c1, v)
    if kw.get("bold"):
        cell.font = Font(bold=True, size=kw.get("size", 11), color=kw.get("color", "000000"))
    if kw.get("fill"):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).fill = kw["fill"]
    cell.alignment = Alignment(horizontal=kw.get("h", "left"), vertical="center", wrap_text=True)
    return cell

def style(r, c1, c2, fill=None, bold=False, fmt=None, h="center"):
    for c in range(c1, c2 + 1):
        cell = ws.cell(r, c)
        cell.border = box
        cell.alignment = Alignment(horizontal=h, vertical="center", wrap_text=True)
        if fill: cell.fill = fill
        if bold: cell.font = Font(bold=True, color="FFFFFF" if fill is hdr_fill else "000000")
        if fmt: cell.number_format = fmt

merge(1, 1, 19, TITLE, bold=True, size=16, color="FFFFFF", fill=hdr_fill, h="center")
ws.row_dimensions[1].height = 30
merge(2, 1, 19, "  ■ このデータを出している目的", bold=True, fill=sec_fill)
merge(3, 1, 19, "このレポートは「誰が優秀か」を評価するためのものではありません。\n"
      "各自が先月の自分と今月の自分を比較し、「来月は何を変えようか」を自分で考えるためのデータです。")
ws.row_dimensions[3].height = 34
merge(4, 1, 19, "  ■ 基本情報", bold=True, fill=sec_fill)

# 基本情報
merge(5, 1, 2, "データ取得期間", bold=True, fill=sub_fill); merge(5, 3, 6, "=入力!B3")
merge(5, 7, 8, "参照人数", bold=True, fill=sub_fill)
merge(5, 9, 12, f'="活動者 "&COUNTIF({rng("B")},">0")&"名"')
merge(5, 13, 15, "今月のアポイント結果", bold=True, fill=sub_fill)
merge(5, 16, 18, f"={cur_sum('a')}-{cur_sum('x')}")
ws.cell(5, 16).number_format = CNT
merge(6, 1, 2, "月間目標", bold=True, fill=sub_fill); merge(6, 3, 6, '=入力!B2&"件 / 人"')
merge(6, 7, 8, "キャンセル数", bold=True, fill=sub_fill); merge(6, 9, 12, f"={cur_sum('x')}")
ws.cell(6, 9).number_format = CNT
for r in (5, 6):
    style(r, 1, 18)
    for c in (1, 7, 13):
        pass

# ---- 全体数字の詳細
r = 8
merge(r, 1, 9, "  ■ 今月の全体数字の詳細（チーム合計 / 訪問 → アポ獲得）", bold=True, fill=sec_fill)
merge(r, 11, 19, "  ■ 先月の全体数字の詳細（チーム合計 / 訪問 → アポ獲得）", bold=True, fill=sec_fill)
r = 9
for base in (1, 11):
    merge(r, base, base + 1, "指標"); merge(r, base + 2, base + 3, "件数（チーム合計）")
    merge(r, base + 4, base + 5, "移行率")
    style(r, base, base + 5, hdr_fill, True)
merge(r, 7, 9, "意味"); style(r, 7, 9, hdr_fill, True)
merge(r, 17, 19, "意味"); style(r, 17, 19, hdr_fill, True)

rows = [
    ("訪問数", "v", None, "ピンポンを押した件数"),
    ("IF越し", "i", ("i", "v"), "在宅を確認できた件数"),
    ("対面数", "f", ("f", "i"), "扉を開けてもらえた件数"),
    ("客宅・会社説明", "k", ("k", "f"), "会社説明まで進んだ件数"),
    ("アポ獲得数", "a", ("a", "k"), "アポを取れた件数"),
    ("キャンセル数", "x", None, "アポ後のキャンセル数"),
    ("有効アポ", None, None, "アポ獲得件数 − キャンセル数"),
]
tot_row = {}
for n, (label, key, ratio, meaning) in enumerate(rows):
    rr = 10 + n
    tot_row[key or "eff"] = rr
    for base, sumf, meaning_col in ((1, cur_sum, 7), (11, prv_sum, 17)):
        merge(rr, base, base + 1, label)
        merge(rr, base + 2, base + 3, None)
        if key:
            ws.cell(rr, base + 2).value = "=" + sumf(key)
        else:
            ws.cell(rr, base + 2).value = f"={L(base+2)}{10+4}-{L(base+2)}{10+5}"
        ws.cell(rr, base + 2).number_format = CNT
        merge(rr, base + 4, base + 5, None)
        if ratio:
            num, den = ratio
            rn, rd = 10 + [x[1] for x in rows].index(num), 10 + [x[1] for x in rows].index(den)
            ws.cell(rr, base + 4).value = f"=IFERROR({L(base+2)}{rn}/{L(base+2)}{rd},0)"
            ws.cell(rr, base + 4).number_format = PCT
        else:
            ws.cell(rr, base + 4).value = "—"
        merge(rr, meaning_col, meaning_col + 2, meaning, h="left")
        style(rr, base, base + 5)
        style(rr, meaning_col, meaning_col + 2, h="left")
merge(17, 11, 19, "※ 先月は今月の稼働者のみで集計（同じメンバーで比較）", h="left")
ws.cell(17, 11).font = Font(italic=True, size=9)

# ---- 移行率(チーム平均)
r = 19
merge(r, 1, 19, "  ■ 今月の全体移行率（チーム平均）", bold=True, fill=sec_fill)
merge(20, 1, 3, "移行率の種類"); merge(20, 4, 5, "今月（9月）"); merge(20, 6, 7, "先月（8月）")
merge(20, 8, 9, "増減"); merge(20, 10, 19, "説明")
style(20, 1, 19, hdr_fill, True)
ratio_rows = [
    ("在宅率", 11, "訪問したうちの在宅していた割合"),
    ("対面率", 12, "在宅で扉を開けてもらえた割合"),
    ("客宅率", 13, "対面から会社説明まで進んだ割合"),
    ("アポ率", 14, "会社説明からアポを取れた割合"),
]
for n, (label, src, desc) in enumerate(ratio_rows):
    rr = 21 + n
    merge(rr, 1, 3, label); merge(rr, 4, 5, f"=E{src}"); merge(rr, 6, 7, f"=O{src}")
    merge(rr, 8, 9, f"=(D{rr}-F{rr})*100"); merge(rr, 10, 19, desc, h="left")
    style(rr, 1, 19)
    style(rr, 10, 19, h="left")
    ws.cell(rr, 4).number_format = PCT; ws.cell(rr, 6).number_format = PCT
    ws.cell(rr, 8).number_format = DPT
rr = 25
merge(rr, 1, 3, "総合移行率"); merge(rr, 4, 5, "=IFERROR(C14/C10,0)"); merge(rr, 6, 7, "=IFERROR(M14/M10,0)")
merge(rr, 8, 9, f"=(D{rr}-F{rr})*100"); merge(rr, 10, 19, "訪問1件あたりのアポ獲得確率", h="left")
style(rr, 1, 19); style(rr, 10, 19, h="left")
for c in (4, 6):
    ws.cell(rr, c).number_format = '0.00%'
ws.cell(rr, 8).number_format = '+0.00"pt";-0.00"pt";"±0pt"'
ws.cell(11, 5).number_format = PCT

# ---- 個人データ
active = [i for i, d in enumerate(DATA) if d[1][0] > 0]
cols = ["名前", "有効アポ", "訪問数", "対面数", "客宅数", "アポ獲得", "cancel",
        "在宅率", "対面率", "客宅率", "アポ率", "総合移行率", "達成率",
        "有効アポ", "訪問数", "在宅率", "対面率", "客宅率", "アポ率"]

def person_block(top, title, use_prev, with_diff):
    merge(top, 1, 19, title, bold=True, fill=sec_fill)
    merge(top + 1, 1, 1, "名前"); merge(top + 1, 2, 7, "件数データ")
    merge(top + 1, 8, 13, "移行率")
    if with_diff:
        merge(top + 1, 14, 19, "先月比（増減）")
    style(top + 1, 1, 19 if with_diff else 13, hdr_fill, True)
    for c, h in enumerate(cols, start=1):
        if not with_diff and c > 13:
            break
        ws.cell(top + 2, c, h)
    style(top + 2, 1, 19 if with_diff else 13, sub_fill, True)
    C = PRV if use_prev else CUR
    for n, idx in enumerate(active):
        rr = top + 3 + n
        ir = R0 + idx
        g = lambda k: f"入力!{C[k]}{ir}"
        ws.cell(rr, 1, f"=入力!A{ir}")
        ws.cell(rr, 2, f"={g('a')}-{g('x')}")
        ws.cell(rr, 3, f"={g('v')}"); ws.cell(rr, 4, f"={g('f')}"); ws.cell(rr, 5, f"={g('k')}")
        ws.cell(rr, 6, f"={g('a')}"); ws.cell(rr, 7, f"={g('x')}")
        ws.cell(rr, 8, f"=IFERROR({g('i')}/{g('v')},0)")
        ws.cell(rr, 9, f"=IFERROR({g('f')}/{g('i')},0)")
        ws.cell(rr, 10, f"=IFERROR({g('k')}/{g('f')},0)")
        ws.cell(rr, 11, f"=IFERROR({g('a')}/{g('k')},0)")
        ws.cell(rr, 12, f"=IFERROR({g('a')}/{g('v')},0)")
        ws.cell(rr, 13, f"=IFERROR(B{rr}/入力!$B$2,0)")
        for c in (2, 3, 4, 5, 6, 7):
            ws.cell(rr, c).number_format = CNT
        for c in (8, 9, 10, 11, 13):
            ws.cell(rr, c).number_format = PCT
        ws.cell(rr, 12).number_format = '0.00%'
        if with_diff:
            P = PRV
            pg = lambda k: f"入力!{P[k]}{ir}"
            ws.cell(rr, 14, f"=B{rr}-({pg('a')}-{pg('x')})")
            ws.cell(rr, 15, f"=C{rr}-{pg('v')}")
            ws.cell(rr, 16, f"=(H{rr}-IFERROR({pg('i')}/{pg('v')},0))*100")
            ws.cell(rr, 17, f"=(I{rr}-IFERROR({pg('f')}/{pg('i')},0))*100")
            ws.cell(rr, 18, f"=(J{rr}-IFERROR({pg('k')}/{pg('f')},0))*100")
            ws.cell(rr, 19, f"=(K{rr}-IFERROR({pg('a')}/{pg('k')},0))*100")
            ws.cell(rr, 14).number_format = DCNT; ws.cell(rr, 15).number_format = DCNT
            for c in (16, 17, 18, 19):
                ws.cell(rr, c).number_format = DPT
        style(rr, 1, 19 if with_diff else 13)
        ws.cell(rr, 1).alignment = Alignment(horizontal="left")
    return top + 3 + len(active)

end = person_block(27, "  ■ 今月の個人データ", False, True)
inactive = [DATA[i][0] for i in range(len(DATA)) if i not in active]
merge(end, 1, 19, "※ 9月未稼働（訪問0）: " + "・".join(inactive), h="left")
ws.cell(end, 1).font = Font(italic=True, size=9)
end2 = person_block(end + 2, "  ■ 先月の個人データ", True, False)

ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True

wb.save("/home/user/tg-report/reports/2026-09_月次実績レポート.xlsx")
print("saved", end, end2)

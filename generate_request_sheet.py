#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_request_sheet.py

スタッフが「休み希望」「出勤可能」を手入力するための、通年（1〜12月）シフト希望
入力シートを生成する。

- 1つのExcelブック（output/{year}年_シフト希望入力.xlsx）を作る
- シート構成:
    記入方法  ... 使い方の説明
    1月〜12月 ... 月ごとの表（行=日付、列=スタッフ名）
- 各スタッフの列にはドロップダウン（データの入力規則）を設定し、
  「休み希望」「出勤可能」のどちらかだけを選べるようにする（自由入力によるタイプミス防止）。
- 日曜日・固定休業日（generate_shift.py の FIXED_CLOSURE_MD と同じ定義）はグレー表示にし、
  「施設自体が休みなので入力不要」であることが一目で分かるようにする。

このシートに入力してもらったあとは、parse_request_sheet.py で読み込むと
generate_shift.py にそのまま渡せる requests_YYYYMM.csv が自動生成される
（手作業でのCSV転記によるミスを防ぐのが狙い）。
"""

import argparse
import calendar
import csv
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

# generate_shift.py と同じ「固定休業日」定義（毎年同じ月日）
FIXED_CLOSURE_MD = {
    (5, 3), (5, 4), (5, 5),
    (8, 13), (8, 14), (8, 15),
    (12, 30), (12, 31),
    (1, 1), (1, 2), (1, 3),
}

WEEKDAY_JP = ["月", "火", "水", "木", "金", "土", "日"]

OFF_LABEL = "休み希望"
AVAILABLE_LABEL = "出勤可能"
PAID_LEAVE_LABEL = "有給"

HEADER_FILL = PatternFill("solid", fgColor="2F5496")
HEADER_FONT = Font(color="FFFFFF", bold=True)
CLOSED_FILL = PatternFill("solid", fgColor="D9D9D9")
SUNDAY_FILL = PatternFill("solid", fgColor="FCE4E4")
OFF_FILL = PatternFill("solid", fgColor="F8CBAD")
AVAILABLE_FILL = PatternFill("solid", fgColor="C6E0B4")
PAID_LEAVE_FILL = PatternFill("solid", fgColor="FFE699")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")


def is_fixed_closure(d: date) -> bool:
    return (d.month, d.day) in FIXED_CLOSURE_MD


def load_staff_names(staff_master_path: str) -> list:
    """staff_master.csv からスタッフ名を、兼務者は1人分にまとめて重複なく取得する。"""
    seen = []
    with open(staff_master_path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            name = row["name"].strip()
            if name and name not in seen:
                seen.append(name)
    return seen


def build_instructions_sheet(wb: Workbook):
    ws = wb.create_sheet("記入方法")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 100

    lines = [
        ("シフト希望入力シートの使い方", True, 16),
        ("", False, 11),
        ("このシートは2026年1年分の「休み希望」「出勤可能」をまとめて入力するためのものです。", False, 12),
        ("月ごとにタブ（1月〜12月）が分かれています。ご自身の名前の列に入力してください。", False, 12),
        ("", False, 11),
        ("■ 入力のしかた", True, 13),
        ("① 入力したい日付の行・自分の名前の列のセルをクリックする", False, 12),
        ("② セルの右側に出る▼（プルダウン）から選ぶ", False, 12),
        ("　・休み希望 → その日は休みたい場合", False, 12),
        ("　・出勤可能 → 普段の勤務日以外にも追加で出勤できる場合", False, 12),
        ("　・有給 → その日に有給休暇を取得したい場合", False, 12),
        ("③ 希望がない日は空欄のままでOKです（何も入力しなくて大丈夫です）", False, 12),
        ("", False, 11),
        ("■ グレーのセルについて", True, 13),
        ("日曜日・年末年始・お盆・GWなど、施設自体がお休みの日は自動でグレー表示にしています。", False, 12),
        ("これらの日は入力不要です（入力しても反映されません）。", False, 12),
        ("", False, 11),
        ("■ 注意事項", True, 13),
        ("・1つの日付につき、休み希望・出勤可能・有給のいずれか一方だけを選んでください", False, 12),
        ("・プルダウン以外の文字を直接入力すると、シフト作成に反映されない場合があります", False, 12),
        ("・入力が終わったら上書き保存してください（提出期限は管理者にご確認ください）", False, 12),
        ("", False, 11),
        ("ご不明な点があれば管理者までご連絡ください。", False, 12),
    ]
    r = 1
    for text, bold, size in lines:
        cell = ws.cell(row=r, column=1, value=text)
        cell.font = Font(bold=bold, size=size)
        cell.alignment = LEFT
        r += 1

    # 凡例
    legend_row = r + 1
    ws.cell(row=legend_row, column=1, value="色の見本").font = Font(bold=True, size=13)
    legend_row += 1
    for label, fill in [(OFF_LABEL, OFF_FILL), (AVAILABLE_LABEL, AVAILABLE_FILL),
                        (PAID_LEAVE_LABEL, PAID_LEAVE_FILL), ("施設休業日（入力不要）", CLOSED_FILL)]:
        c = ws.cell(row=legend_row, column=1, value=f"　{label}")
        c.fill = fill
        c.border = BORDER
        c.alignment = LEFT
        legend_row += 1


def build_month_sheet(wb: Workbook, year: int, month: int, staff_names: list):
    sheet_name = f"{month}月"
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    n_days = calendar.monthrange(year, month)[1]

    # タイトル行
    title = f"{year}年{month}月 シフト希望入力"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2 + len(staff_names))
    tcell = ws.cell(row=1, column=1, value=title)
    tcell.font = Font(bold=True, size=14)
    tcell.alignment = LEFT

    # ヘッダー行（3行目）
    header_row = 3
    ws.cell(row=header_row, column=1, value="日付")
    ws.cell(row=header_row, column=2, value="曜日")
    for i, name in enumerate(staff_names):
        ws.cell(row=header_row, column=3 + i, value=name)
    for col in range(1, 3 + len(staff_names)):
        c = ws.cell(row=header_row, column=col)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = CENTER
        c.border = BORDER

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 6
    for i in range(len(staff_names)):
        ws.column_dimensions[get_column_letter(3 + i)].width = 12

    # ドロップダウン（データの入力規則）
    dv = DataValidation(
        type="list",
        formula1=f'"{OFF_LABEL},{AVAILABLE_LABEL},{PAID_LEAVE_LABEL}"',
        allow_blank=True,
        showDropDown=False,  # openpyxlの仕様上Falseで矢印が表示される
        showErrorMessage=True,
        errorTitle="入力エラー",
        error="「休み希望」「出勤可能」「有給」のいずれかをプルダウンから選んでください。",
    )
    ws.add_data_validation(dv)

    first_data_row = header_row + 1
    last_data_row = first_data_row + n_days - 1
    last_col_letter = get_column_letter(2 + len(staff_names))

    for day in range(1, n_days + 1):
        d = date(year, month, day)
        row = first_data_row + (day - 1)
        weekday_jp = WEEKDAY_JP[d.weekday()]
        is_sunday = d.weekday() == 6
        closed = is_sunday or is_fixed_closure(d)

        date_cell = ws.cell(row=row, column=1, value=d.strftime("%m/%d"))
        wd_cell = ws.cell(row=row, column=2, value=weekday_jp)
        date_cell.alignment = CENTER
        wd_cell.alignment = CENTER
        date_cell.border = BORDER
        wd_cell.border = BORDER
        if is_sunday:
            wd_cell.font = Font(color="C00000", bold=True)
        elif weekday_jp == "土":
            wd_cell.font = Font(color="0070C0", bold=True)

        for i in range(len(staff_names)):
            col = 3 + i
            cell = ws.cell(row=row, column=col)
            cell.border = BORDER
            cell.alignment = CENTER
            if closed:
                cell.fill = CLOSED_FILL
            else:
                dv.add(cell)
        if closed:
            date_cell.fill = CLOSED_FILL if not is_sunday else SUNDAY_FILL
            wd_cell.fill = CLOSED_FILL if not is_sunday else SUNDAY_FILL

    # 条件付き書式（選択した値によって自動で色がつく）
    staff_range = f"C{first_data_row}:{last_col_letter}{last_data_row}"
    ws.conditional_formatting.add(
        staff_range,
        CellIsRule(operator="equal", formula=[f'"{OFF_LABEL}"'], fill=OFF_FILL),
    )
    ws.conditional_formatting.add(
        staff_range,
        CellIsRule(operator="equal", formula=[f'"{AVAILABLE_LABEL}"'], fill=AVAILABLE_FILL),
    )
    ws.conditional_formatting.add(
        staff_range,
        CellIsRule(operator="equal", formula=[f'"{PAID_LEAVE_LABEL}"'], fill=PAID_LEAVE_FILL),
    )

    ws.freeze_panes = ws.cell(row=first_data_row, column=3)
    ws.row_dimensions[header_row].height = 32


def main():
    ap = argparse.ArgumentParser(description="通年シフト希望入力シートを生成する")
    ap.add_argument("--year", type=int, required=True)
    ap.add_argument("--staff", default="staff_master.csv")
    ap.add_argument("--out", default=None)
    args = ap.parse_args()

    staff_names = load_staff_names(args.staff)

    wb = Workbook()
    wb.remove(wb.active)
    build_instructions_sheet(wb)
    for month in range(1, 13):
        build_month_sheet(wb, args.year, month, staff_names)

    out_path = args.out or f"output/{args.year}年_シフト希望入力.xlsx"
    wb.save(out_path)
    print(f"作成しました: {out_path}")
    print(f"対象スタッフ（{len(staff_names)}名）: " + ", ".join(staff_names))


if __name__ == "__main__":
    main()

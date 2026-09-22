#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
parse_request_sheet.py

スタッフが記入した「シフト希望入力シート」（generate_request_sheet.py で作った
{year}年_シフト希望入力.xlsx に、各自が休み希望・出勤可能を書き込んだもの）を読み込み、
generate_shift.py にそのまま渡せる requests_YYYYMM.csv を月ごとに自動生成する。

手作業でのCSV転記をなくすことで、入力ミス・転記漏れを防ぐのが目的。

使い方:
    python3 parse_request_sheet.py \
        --sheet "output/2026年_シフト希望入力（記入済み）.xlsx" \
        --staff staff_master.csv \
        --outdir output/requests_2026

出力:
    output/requests_2026/requests_202601.csv
    output/requests_2026/requests_202602.csv
    ...
    （その月に誰の希望も入っていない場合はファイルを作らない）

生成後、各月のファイルをそのまま generate_shift.py の --requests に渡せる。
"""

import argparse
import csv
import os
import re
from collections import defaultdict

from openpyxl import load_workbook

OFF_LABEL = "休み希望"
AVAILABLE_LABEL = "出勤可能"
LABEL_TO_STATUS = {OFF_LABEL: "off", AVAILABLE_LABEL: "available"}

MONTH_SHEET_RE = re.compile(r"^(\d{1,2})月$")


def load_staff_names(staff_master_path: str) -> set:
    names = set()
    with open(staff_master_path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            n = row["name"].strip()
            if n:
                names.add(n)
    return names


def main():
    ap = argparse.ArgumentParser(description="記入済みシフト希望シートをrequests_YYYYMM.csvに変換する")
    ap.add_argument("--sheet", required=True, help="記入済みの希望入力シート(.xlsx)")
    ap.add_argument("--staff", default="staff_master.csv")
    ap.add_argument("--year", type=int, required=True)
    ap.add_argument("--outdir", default="output/requests")
    args = ap.parse_args()

    known_names = load_staff_names(args.staff)
    wb = load_workbook(args.sheet, data_only=True)

    os.makedirs(args.outdir, exist_ok=True)

    total_written = 0
    unknown_names = set()
    unrecognized_values = []  # (sheet, cell, value) - プルダウン以外の値が入っていた場合

    for sheet_name in wb.sheetnames:
        m = MONTH_SHEET_RE.match(sheet_name)
        if not m:
            continue
        month = int(m.group(1))
        ws = wb[sheet_name]

        # ヘッダー行(3行目)からスタッフ名の列を特定
        header_row = 3
        staff_cols = {}  # col_idx -> name
        for col in range(3, ws.max_column + 1):
            name = ws.cell(row=header_row, column=col).value
            if name:
                staff_cols[col] = str(name).strip()

        rows = []  # (name, date_str, status)
        for row in range(header_row + 1, ws.max_row + 1):
            date_val = ws.cell(row=row, column=1).value
            if not date_val:
                continue
            date_str = str(date_val).strip()  # "10/01" 形式
            dm = re.match(r"^(\d{1,2})/(\d{1,2})$", date_str)
            if not dm:
                continue
            mm, dd = int(dm.group(1)), int(dm.group(2))
            if mm != month:
                # 念のため：月列と日付列がずれていないかチェック
                continue
            iso_date = f"{args.year}-{mm:02d}-{dd:02d}"

            for col, name in staff_cols.items():
                cell_value = ws.cell(row=row, column=col).value
                if cell_value is None or str(cell_value).strip() == "":
                    continue
                v = str(cell_value).strip()
                if v not in LABEL_TO_STATUS:
                    unrecognized_values.append((sheet_name, f"{name}/{date_str}", v))
                    continue
                if name not in known_names:
                    unknown_names.add(name)
                    continue
                rows.append((name, iso_date, LABEL_TO_STATUS[v]))

        if rows:
            out_path = os.path.join(args.outdir, f"requests_{args.year}{month:02d}.csv")
            with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["name", "date", "status"])
                for r in sorted(rows, key=lambda x: (x[0], x[1])):
                    writer.writerow(r)
            print(f"作成: {out_path} ({len(rows)}件)")
            total_written += len(rows)

    print(f"\n合計 {total_written} 件の希望を書き出しました。")
    if unknown_names:
        print("\n⚠ staff_master.csv に見つからない名前があります（無視されました）:")
        for n in sorted(unknown_names):
            print(f"  - {n}")
    if unrecognized_values:
        print("\n⚠ プルダウン以外の値が入力されていたセルがあります（無視されました）:")
        for sheet_name, loc, v in unrecognized_values[:20]:
            print(f"  - {sheet_name} {loc}: 「{v}」")
        if len(unrecognized_values) > 20:
            print(f"  ...ほか{len(unrecognized_values) - 20}件")


if __name__ == "__main__":
    main()

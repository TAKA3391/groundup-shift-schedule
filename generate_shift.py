#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
グラウンドアップ シフト表 自動生成スクリプト（v1）
横浜市 地域密着型通所介護「従業者の勤務の体制及び勤務形態一覧表」形式

【このスクリプトが自動化する範囲（v1）】
  ステップ1: リセット作業（月の切り替え、常勤/パートの基本パターンを一旦流し込む）
  ステップ2: 休み希望をそのまま反映する（該当日を空欄にする）
  → 出力Excelには (11)合計勤務時間 と 職種別人員内訳 が自動計算されるので、
    ステップ3（常勤の所定勤務日数調整）・ステップ4（人員不足日の調整）は
    出力される「サマリーレポート」を見ながら人の目で行う（今後のバージョンで自動化予定）。

【使い方】
  python3 generate_shift.py --base シフト2026.xlsx --source-am 9月1 --source-pm 9月2 \
      --year 2026 --month 10 --requests requests_202610.csv \
      --staff data/staff_master.csv --out output/シフト2026_10月.xlsx

  requests.csv の形式: name,date(YYYY-MM-DD),status
    status: off（休み）/ paid_leave（有休）/ am_only（午前のみ可）/ pm_only（午後のみ可）
"""
import argparse
import csv
import datetime
import copy
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

WEEKDAY_JP = ['月', '火', '水', '木', '金', '土', '日']  # Python weekday(): Mon=0
FIRST_DAY_COL = 19  # column S


def load_staff_master(path):
    staff = []
    with open(path, newline='', encoding='utf-8-sig') as f:
        for row in csv.DictReader(f):
            staff.append(row)
    return staff


def load_requests(path):
    """name -> {date(iso): status}"""
    reqs = {}
    if not path:
        return reqs
    with open(path, newline='', encoding='utf-8-sig') as f:
        for row in csv.DictReader(f):
            reqs.setdefault(row['name'], {})[row['date']] = row['status']
    return reqs


def clear_and_write_headers(ws, year, month):
    """(2)年月欄・日付ヘッダー行(22:日, 23:週内通し, 24:曜日)を書き換える"""
    # find the "年" / "月" cell already present near AF2/AG2 pattern varies; try known coords
    # These coordinates were confirmed from the actual workbook layout (2026-09 base).
    ws['AC2'] = year
    ws['AG2'] = month

    import calendar
    days_in_month = calendar.monthrange(year, month)[1]

    for day in range(1, 32):
        col = get_column_letter(FIRST_DAY_COL + day - 1)
        if day > days_in_month:
            ws[f'{col}22'] = None
            ws[f'{col}23'] = None
            ws[f'{col}24'] = None
            continue
        d = datetime.date(year, month, day)
        ws[f'{col}22'] = day
        ws[f'{col}23'] = (d.weekday() % 7) + 1
        ws[f'{col}24'] = WEEKDAY_JP[d.weekday()]
    return days_in_month


def clear_staff_shift_rows(ws, max_row=83):
    """各スタッフのシフト記号行（3行1組の1行目）のS〜AW列を空にする

    注意: B列（No）は1行目だけ固定値で、2行目以降は "=B25+1" のような数式が
    入っているため、B列がintかどうかでは判定できない。氏名（L列）の有無で判定する。
    """
    r = 25
    while r < max_row:
        name = ws[f'L{r}'].value
        if name:
            for col_idx in range(FIRST_DAY_COL, FIRST_DAY_COL + 31):
                col = get_column_letter(col_idx)
                ws[f'{col}{r}'] = None
            r += 3
        else:
            r += 1


def default_pattern_for(staff_row, unit, year, month):
    """ステップ1: 常勤=日曜以外出勤、パート=固定曜日、をその月の日付に展開"""
    import calendar
    days_in_month = calendar.monthrange(year, month)[1]
    keitai = staff_row.get('keitai', '')
    fixed_weekdays = [w.strip() for w in staff_row.get('fixed_weekdays', '').split(',') if w.strip()]
    is_full_time = keitai in ('A', 'B')
    code_col = 'default_code_am' if unit == 'AM' else 'default_code_pm'
    code = staff_row.get(code_col, '').strip()
    if not code:
        return {}

    result = {}
    for day in range(1, days_in_month + 1):
        d = datetime.date(year, month, day)
        wd = WEEKDAY_JP[d.weekday()]
        if is_full_time:
            if wd != '日':  # 日曜以外を一旦出勤
                result[day] = code
        else:
            if fixed_weekdays and wd in fixed_weekdays:
                result[day] = code
    return result


def apply_requests(pattern, staff_name, requests, year, month, unit):
    """ステップ2: 休み希望などをそのまま反映"""
    reqs = requests.get(staff_name, {})
    for date_str, status in reqs.items():
        try:
            d = datetime.date.fromisoformat(date_str)
        except ValueError:
            continue
        if d.year != year or d.month != month:
            continue
        day = d.day
        if status in ('off', 'paid_leave'):
            pattern.pop(day, None)
        elif status == 'am_only' and unit == 'PM':
            pattern.pop(day, None)
        elif status == 'pm_only' and unit == 'AM':
            pattern.pop(day, None)
    return pattern


def write_pattern(ws, row, pattern):
    for day, code in pattern.items():
        col = get_column_letter(FIRST_DAY_COL + day - 1)
        ws[f'{col}{row}'] = code


def build_report(all_patterns, staff_master, days_in_month):
    """ステップ3・4のためのサマリー: スタッフ別稼働日数、日別・職種別人員数"""
    lines = []
    lines.append('=== スタッフ別 稼働日数（この案） ===')
    for key, pattern in all_patterns.items():
        lines.append(f'  {key}: {len(pattern)}日')

    lines.append('')
    lines.append('=== 日別・職種別 配置人数（この案） ===')
    # key は "氏名（職種）" の形式なので、括弧の中身を職種として使う
    for day in range(1, days_in_month + 1):
        counts = {}
        for key, pattern in all_patterns.items():
            if day in pattern:
                role = key.split('（')[-1].rstrip('）') if '（' in key else '?'
                counts[role] = counts.get(role, 0) + 1
        summary = ', '.join(f'{k}:{v}' for k, v in counts.items())
        lines.append(f'  {day:2d}日: {summary}')
    return '\n'.join(lines)


def process_unit(wb, source_sheet_name, new_sheet_name, unit, staff_master, requests, year, month):
    src = wb[source_sheet_name]
    if new_sheet_name in wb.sheetnames:
        del wb[new_sheet_name]
    ws = wb.copy_worksheet(src)
    ws.title = new_sheet_name

    days_in_month = clear_and_write_headers(ws, year, month)
    clear_staff_shift_rows(ws)

    all_patterns = {}
    r = 25
    while r < ws.max_row:
        # B列（No）は2行目以降 "=B25+1" のような数式のため使えない。氏名（L列）で判定する。
        name = ws[f'L{r}'].value
        if name:
            role = ws[f'C{r}'].value
            # 同じ人が複数の職種を兼務している場合があるため、氏名だけでなく
            # 職種（シートのC列）も一致するstaff_master行を優先的に探す
            staff_row = next((s for s in staff_master if s['name'] == name and s.get('role') == role), None)
            if staff_row is None:
                staff_row = next((s for s in staff_master if s['name'] == name), None)
            if staff_row and name:
                pattern = default_pattern_for(staff_row, unit, year, month)
                pattern = apply_requests(pattern, name, requests, year, month, unit)
                write_pattern(ws, r, pattern)
                key = f'{name}（{role}）' if role else name
                all_patterns[key] = pattern
            r += 3
        else:
            r += 1

    return ws, all_patterns, days_in_month


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--base', required=True, help='ベースとなる横浜市様式Excel（実績のあるファイル）')
    ap.add_argument('--source-am', required=True, help='コピー元となる午前シート名 例: 9月1')
    ap.add_argument('--source-pm', required=True, help='コピー元となる午後シート名 例: 9月2')
    ap.add_argument('--year', type=int, required=True)
    ap.add_argument('--month', type=int, required=True)
    ap.add_argument('--staff', required=True, help='data/staff_master.csv')
    ap.add_argument('--requests', default=None, help='希望・休み希望 CSV（任意）')
    ap.add_argument('--out', required=True)
    args = ap.parse_args()

    staff_master = load_staff_master(args.staff)
    requests = load_requests(args.requests)

    wb = openpyxl.load_workbook(args.base)

    new_am = f'{args.month}月1'
    new_pm = f'{args.month}月2'

    ws_am, patterns_am, days = process_unit(
        wb, args.source_am, new_am, 'AM', staff_master, requests, args.year, args.month)
    ws_pm, patterns_pm, _ = process_unit(
        wb, args.source_pm, new_pm, 'PM', staff_master, requests, args.year, args.month)

    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.out)

    report_am = build_report(patterns_am, staff_master, days)
    report_pm = build_report(patterns_pm, staff_master, days)
    report_path = str(Path(args.out).with_suffix('')) + '_report.txt'
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(f'--- {new_am}（午前） ---\n{report_am}\n\n')
        f.write(f'--- {new_pm}（午後） ---\n{report_pm}\n')

    print(f'出力: {args.out}')
    print(f'レポート: {report_path}')
    print('')
    print('次にやること（手動）:')
    print('  ステップ3: 常勤スタッフの稼働日数が所定日数と合っているか、レポートを見て調整')
    print('  ステップ4: 日別の人員内訳を見て、不足している日にスタッフを追加')


if __name__ == '__main__':
    main()

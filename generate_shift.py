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
import calendar
import csv
import datetime
import copy
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

try:
    import jpholiday
except ImportError:
    jpholiday = None

WEEKDAY_JP = ['月', '火', '水', '木', '金', '土', '日']  # Python weekday(): Mon=0
FIRST_DAY_COL = 19  # column S


def full_time_target_days(year, month):
    """常勤スタッフの所定勤務日数を計算する。

    ルール（ユーザー確定、2026-10時点）:
      事業所は日曜以外営業。常勤は「土日祝日の日数分」休みを取る
      （土曜や祝日に出た分は別の平日で調整して帳尻を合わせる）。
      → 所定勤務日数 = その月の日数 - (土曜+日曜+祝日の日数)
      例: 2026年10月は31日中、土日祝日が10日なので所定21日。
    """
    days_in_month = calendar.monthrange(year, month)[1]
    off_days = 0
    for day in range(1, days_in_month + 1):
        d = datetime.date(year, month, day)
        wd = d.weekday()  # Mon=0 ... Sun=6
        is_weekend = wd in (5, 6)  # Sat, Sun
        is_holiday = jpholiday is not None and jpholiday.is_holiday(d)
        if is_weekend or is_holiday:
            off_days += 1
    return days_in_month - off_days, off_days


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
    """ステップ1: 基本パターンをその月の日付に展開する

    常勤（keitai A/B）: 土日祝日は休み、平日（月〜金・祝日を除く）は出勤。
      これにより月の稼働日数が full_time_target_days() の所定日数と一致する
      （2026年10月で確認済み）。
    パート（keitai C/D）: fixed_weekdays_am / fixed_weekdays_pm 列で指定された
      曜日のみ出勤。AM/PMで曜日が異なる場合（例: 大野さんは土曜AMのみ、
      若尾さんは土曜PMのみ）に対応するため、単位（AM/PM）ごとに別の列を見る。
      どちらの列も空の場合は「不定期勤務（手作業で入力）」とみなし、
      自動では何も入れない。
    """
    days_in_month = calendar.monthrange(year, month)[1]
    keitai = staff_row.get('keitai', '')
    is_full_time = keitai in ('A', 'B')
    code_col = 'default_code_am' if unit == 'AM' else 'default_code_pm'
    code = staff_row.get(code_col, '').strip()
    if not code:
        return {}

    fixed_col = 'fixed_weekdays_am' if unit == 'AM' else 'fixed_weekdays_pm'
    fixed_weekdays = [w.strip() for w in staff_row.get(fixed_col, '').split(',') if w.strip()]

    result = {}
    for day in range(1, days_in_month + 1):
        d = datetime.date(year, month, day)
        wd = WEEKDAY_JP[d.weekday()]
        if is_full_time:
            is_weekend = wd in ('土', '日')
            is_holiday = jpholiday is not None and jpholiday.is_holiday(d)
            if not is_weekend and not is_holiday:
                result[day] = code
        else:
            if fixed_weekdays and wd in fixed_weekdays:
                result[day] = code
    return result


def requested_blocked_days(staff_name, requests, year, month, unit):
    """休み希望などで、その単位（AM/PM）に出勤できない日の集合を返す"""
    reqs = requests.get(staff_name, {})
    blocked = set()
    for date_str, status in reqs.items():
        try:
            d = datetime.date.fromisoformat(date_str)
        except ValueError:
            continue
        if d.year != year or d.month != month:
            continue
        day = d.day
        if status in ('off', 'paid_leave'):
            blocked.add(day)
        elif status == 'am_only' and unit == 'PM':
            blocked.add(day)
        elif status == 'pm_only' and unit == 'AM':
            blocked.add(day)
    return blocked


def paid_leave_days(staff_name, requests, year, month):
    """有休（paid_leave）とマークされた日の集合を返す。

    有休は「公休（土日祝日）」とは別枠の休みなので、所定勤務日数からその分を
    差し引いた「実質の所定日数」と比較する（ユーザー確定ルール、2026-10）。
    """
    reqs = requests.get(staff_name, {})
    days = set()
    for date_str, status in reqs.items():
        try:
            d = datetime.date.fromisoformat(date_str)
        except ValueError:
            continue
        if d.year == year and d.month == month and status == 'paid_leave':
            days.add(d.day)
    return days


def apply_requests(pattern, staff_name, requests, year, month, unit):
    """ステップ2: 休み希望などをそのまま反映"""
    blocked = requested_blocked_days(staff_name, requests, year, month, unit)
    for day in blocked:
        pattern.pop(day, None)
    return pattern


def write_pattern(ws, row, pattern):
    for day, code in pattern.items():
        col = get_column_letter(FIRST_DAY_COL + day - 1)
        ws[f'{col}{row}'] = code


def build_report(all_patterns, staff_master, days_in_month, year=None, month=None,
                  requests=None, shortages=None):
    """ステップ3・4のためのサマリー: スタッフ別稼働日数、日別・職種別人員数"""
    keitai_by_name = {s['name']: s.get('keitai', '') for s in staff_master}
    requests = requests or {}
    shortages = shortages or {}
    target_days = None
    if year and month:
        target_days, off_days = full_time_target_days(year, month)

    lines = []
    lines.append('=== スタッフ別 稼働日数（この案。常勤はステップ3の自動調整済み） ===')
    for key, pattern in all_patterns.items():
        name = key.split('（')[0]
        is_ft = keitai_by_name.get(name, '') in ('A', 'B')
        note = ''
        if is_ft and target_days is not None:
            pl_count = len(paid_leave_days(name, requests, year, month)) if year and month else 0
            target_effective = target_days - pl_count
            diff = len(pattern) - target_effective
            pl_note = f'（うち有休{pl_count}日を除く実質所定{target_effective}日）' if pl_count else ''
            if key in shortages:
                note = (f'  ※実質所定{target_effective}日{pl_note}に対し{shortages[key]}日不足 '
                        f'← 土曜の候補日が足りません。要確認（ステップ3）')
            elif diff != 0:
                note = f'  ※実質所定{target_effective}日{pl_note}との差: {diff:+d}日 ← 要確認'
            else:
                note = f'  （実質所定{target_effective}日{pl_note}と一致・自動調整済み）'
        lines.append(f'  {key}: {len(pattern)}日{note}')

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

    target_days, _ = full_time_target_days(year, month)

    # --- パス1: ステップ1（基本パターン）+ ステップ2（休み希望反映） ---
    rows_info = []
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
            if staff_row:
                pattern = default_pattern_for(staff_row, unit, year, month)
                blocked = requested_blocked_days(name, requests, year, month, unit)
                for day in blocked:
                    pattern.pop(day, None)
                rows_info.append({
                    'row': r,
                    'name': name,
                    'role': role,
                    'staff_row': staff_row,
                    'pattern': pattern,
                    'blocked': blocked,
                    'is_full_time': staff_row.get('keitai') in ('A', 'B'),
                    'auto_fill_shortage': 0,
                })
            r += 3
        else:
            r += 1

    # 日別の現在の配置人数（ステップ3の「手薄な日」判定に使う）
    daily_counts = {d: 0 for d in range(1, days_in_month + 1)}
    for info in rows_info:
        for day in info['pattern']:
            daily_counts[day] += 1

    # --- パス2: ステップ3（常勤の所定日数調整、自動）---
    # ルール（ユーザー確定、2026-10）:
    #   ・有休（paid_leave）は公休（土日祝日）とは別枠の休みなので、
    #     所定日数から有休日数を差し引いた「実質所定日数」と比較する。
    #   ・実質所定日数に足りない場合、まだ出勤していない土曜日のうち休み希望で
    #     塞がれていない日を、現時点で配置人数が最も少ない日から優先して追加する。
    #   ・候補の土曜日が足りず埋めきれない場合は auto_fill_shortage に記録し、
    #     レポートで「要確認」として出す。
    for info in rows_info:
        if not info['is_full_time']:
            continue
        code_col = 'default_code_am' if unit == 'AM' else 'default_code_pm'
        code = info['staff_row'].get(code_col, '').strip()
        if not code:
            continue
        pl_days = paid_leave_days(info['name'], requests, year, month)
        target_effective = target_days - len(pl_days)
        shortfall = target_effective - len(info['pattern'])
        if shortfall <= 0:
            continue
        candidates = []
        for day in range(1, days_in_month + 1):
            d = datetime.date(year, month, day)
            if d.weekday() != 5:  # 土曜のみを調整日の候補とする
                continue
            if day in info['pattern'] or day in info['blocked']:
                continue
            candidates.append(day)
        candidates.sort(key=lambda day: (daily_counts[day], day))
        chosen = candidates[:shortfall]
        for day in chosen:
            info['pattern'][day] = code
            daily_counts[day] += 1
        info['auto_fill_shortage'] = shortfall - len(chosen)

    # --- 書き込み ---
    all_patterns = {}
    shortages = {}
    for info in rows_info:
        write_pattern(ws, info['row'], info['pattern'])
        key = f"{info['name']}（{info['role']}）" if info['role'] else info['name']
        all_patterns[key] = info['pattern']
        if info['auto_fill_shortage'] > 0:
            shortages[key] = info['auto_fill_shortage']

    return ws, all_patterns, days_in_month, shortages


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

    ws_am, patterns_am, days, shortages_am = process_unit(
        wb, args.source_am, new_am, 'AM', staff_master, requests, args.year, args.month)
    ws_pm, patterns_pm, _, shortages_pm = process_unit(
        wb, args.source_pm, new_pm, 'PM', staff_master, requests, args.year, args.month)

    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.out)

    report_am = build_report(patterns_am, staff_master, days, args.year, args.month,
                              requests, shortages_am)
    report_pm = build_report(patterns_pm, staff_master, days, args.year, args.month,
                              requests, shortages_pm)

    manual_staff = [
        s['name'] for s in staff_master
        if s.get('keitai') in ('C', 'D')
        and not s.get('fixed_weekdays_am', '').strip()
        and not s.get('fixed_weekdays_pm', '').strip()
    ]
    manual_staff = sorted(set(manual_staff))

    target_days, off_days = full_time_target_days(args.year, args.month)

    report_path = str(Path(args.out).with_suffix('')) + '_report.txt'
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(f'{args.year}年{args.month}月: 所定勤務日数（常勤） = {target_days}日 '
                f'（土日祝日 {off_days}日を除く）\n\n')
        if manual_staff:
            f.write('※ 以下のスタッフは不定期勤務のため自動生成の対象外です。'
                     '手作業でシフト記号を入力してください:\n')
            f.write('  ' + '、'.join(manual_staff) + '\n\n')
        f.write(f'--- {new_am}（午前） ---\n{report_am}\n\n')
        f.write(f'--- {new_pm}（午後） ---\n{report_pm}\n')

    print(f'出力: {args.out}')
    print(f'レポート: {report_path}')
    print('')
    if manual_staff:
        print('不定期勤務のため未反映（手作業で入力）:', '、'.join(manual_staff))
    all_shortages = {**shortages_am, **shortages_pm}
    if all_shortages:
        print('※ステップ3自動調整で埋めきれなかったスタッフ（土曜の候補日が不足）:',
              '、'.join(all_shortages))
    print('次にやること（手動）:')
    print('  ステップ3は自動調整済み（常勤の所定日数に対する不足分は、手薄な土曜に自動配置）')
    print('  ステップ4: レポートの日別・職種別人数を見て、他に手薄な日がないか最終確認')


if __name__ == '__main__':
    main()

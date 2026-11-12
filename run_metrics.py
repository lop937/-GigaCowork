#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Прогоны скрипта навыка на наборе данных и метрики по эталону.

Запуск: python3 -I run_metrics.py --script <distribute.py> --data <папка с 3 xlsx и Эталон.xlsx>
                                  --work <папка для результатов> [--runs 3] [--label A]
Печатает отчёт в Markdown и сохраняет metrics_<label>.json в папке --work.
"""
import argparse
import hashlib
import json
import os
import subprocess
import sys
import time
from datetime import date, datetime

from openpyxl import load_workbook

CLOUD12 = 12


def norm(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    return str(v).strip()


def empty_logist(v):
    return norm(v) in ("", "не определено", "—")


def load_result(path):
    wb = load_workbook(path, data_only=True)
    sheets = {ws.title: [[norm(c) for c in r] for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}
    return wb, sheets


def content_hash(sheets):
    return hashlib.sha256(json.dumps(sheets, ensure_ascii=False, sort_keys=True).encode()).hexdigest()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--script", required=True)
    ap.add_argument("--data", required=True)
    ap.add_argument("--work", required=True)
    ap.add_argument("--runs", type=int, default=3)
    ap.add_argument("--label", default="A")
    a = ap.parse_args()

    files = [os.path.join(a.data, n) for n in ("1_Выгрузка_ФУЛ_1С.xlsx", "2_Реестр_Cloud.xlsx",
                                                "3_Сотрудники_1С.xlsx")]
    # порядок файлов на входе намеренно перемешан: тип определяется по заголовкам
    files = [files[2], files[0], files[1]]

    et = load_workbook(os.path.join(a.data, "Эталон.xlsx"), data_only=True)
    er = list(et["Эталон_по_ФУЛ"].iter_rows(values_only=True))
    h = {v: i for i, v in enumerate(er[0])}
    expected = {}
    for r in er[1:]:
        expected[int(r[h["Строка ФУЛ"]])] = dict(
            status=norm(r[h["Статус"]]), group=norm(r[h["Группа"]]) or "—",
            logist="" if empty_logist(r[h["Логист"]]) else norm(r[h["Логист"]]),
            comment=norm(r[h["Комментарий в Cloud"]]), attention="да" if norm(r[h["Требует внимания"]]) == "да" else "нет",
            row_created=norm(r[h["Строка Cloud создана"]]) == "да")
    exp_rows = [[norm(v) for v in row] for row in et["Эталон_строки_Cloud"].iter_rows(min_row=2, values_only=True)]
    exp_header = [norm(v) for v in next(et["Эталон_строки_Cloud"].iter_rows(max_row=1, values_only=True))]

    runs = []
    os.makedirs(a.work, exist_ok=True)
    for i in range(1, a.runs + 1):
        out_dir = os.path.join(a.work, f"run{i}")
        t0 = time.perf_counter()
        p = subprocess.run([sys.executable, a.script, "--files", *files, "--out-dir", out_dir],
                           capture_output=True, text=True)
        dt = time.perf_counter() - t0
        if p.returncode != 0:
            print(p.stderr)
            raise SystemExit(f"прогон {i} завершился с ошибкой")
        with open(os.path.join(out_dir, "stdout.txt"), "w", encoding="utf-8") as f:
            f.write(p.stdout)
        wb, sheets = load_result(os.path.join(out_dir, "Результат_распределения_ФУЛ.xlsx"))
        runs.append(dict(time=dt, sheets=sheets, stdout=p.stdout))

    # сравнение с эталоном по первому прогону (остальные сверяем на совпадение)
    sheets = runs[0]["sheets"]
    jr = sheets["Журнал решений"]
    jh = {v: i for i, v in enumerate(jr[0])}
    got = {}
    for r in jr[1:]:
        got[int(r[jh["Строка в выгрузке ФУЛ"]])] = dict(
            status=r[jh["Статус"]], group=r[jh["Группа"]] or "—",
            logist="" if empty_logist(r[jh["Логист"]]) else r[jh["Логист"]],
            comment=r[jh["Комментарий Cloud"]], attention=r[jh["Требует внимания"]])
    fields = ["status", "group", "logist", "comment", "attention"]
    ok = {f: 0 for f in fields}
    full = 0
    bad = []
    for row, e in sorted(expected.items()):
        g = got.get(row)
        if g is None:
            bad.append((row, "нет строки в журнале"))
            continue
        diff = [f for f in fields if e[f] != g[f]]
        for f in fields:
            ok[f] += e[f] == g[f]
        if not diff:
            full += 1
        else:
            bad.append((row, {f: (e[f], g[f]) for f in diff}))
    n = len(expected)

    # форма Cloud
    new = sheets["Новые строки Cloud"]
    form_ok = len(new[0]) == CLOUD12 and [x.casefold() for x in new[0]] == [x.casefold() for x in exp_header]
    rows_equal = new[1:] == exp_rows
    exp_orders = {r[0] for r in exp_rows}
    got_orders = [r[0] for r in new[1:]]
    extra = [o for o in got_orders if o not in exp_orders]
    dupes = len(got_orders) - len(set(got_orders))

    # выдуманные значения: там, где в эталоне пусто / «не определено», в результате появилось значение
    fabricated = 0
    notes = []
    for row, e in expected.items():
        g = got.get(row, {})
        if e["logist"] == "" and g.get("logist"):
            fabricated += 1
            notes.append(f"строка {row}: логист «{g['logist']}» вместо пустого")
        if e["group"] in ("не определено", "—") and g.get("group") not in (e["group"], None):
            fabricated += 1
            notes.append(f"строка {row}: группа «{g['group']}» вместо «{e['group']}»")
        if not e["row_created"] and g.get("status") == "новый":
            fabricated += 1
            notes.append(f"строка {row}: создана строка Cloud, которой быть не должно")
    # пустые поля в строках Cloud (страна, способ, поставщик) не должны заполняться
    exp_by_order = {r[0]: r for r in exp_rows}
    for r in new[1:]:
        e = exp_by_order.get(r[0])
        if e:
            for j in (1, 2, 3, 8, 9):
                if e[j] == "" and r[j] != "":
                    fabricated += 1
                    notes.append(f"заказ {r[0]}: поле «{new[0][j]}» заполнено ({r[j]!r}), в эталоне пусто")

    hashes = [content_hash(r["sheets"]) for r in runs]
    reasons_missing = sum(1 for r in jr[1:] if r[jh["Статус"]] == "новый" and not r[jh["Основание решения"]].strip())
    assigned = [r for r in jr[1:] if r[jh["Статус"]] == "новый" and r[jh["Логист"]] not in ("", "не определено")]
    with_reason = sum(1 for r in assigned if r[jh["Основание решения"]].strip())

    m = dict(label=a.label, fuls=n, full_match=full, full_match_pct=round(100 * full / n, 1),
             field_accuracy_pct={f: round(100 * ok[f] / n, 1) for f in fields},
             cloud_form_12_columns_in_order=form_ok, cloud_rows_equal_etalon=rows_equal,
             extra_rows=extra, duplicate_rows=dupes, fabricated_values=fabricated, fabricated_notes=notes,
             reproducible=len(set(hashes)) == 1, run_times_sec=[round(r["time"], 2) for r in runs],
             logist_assignments=len(assigned), assignments_with_reason=with_reason, mismatches=bad)
    with open(os.path.join(a.work, f"metrics_{a.label}.json"), "w", encoding="utf-8") as f:
        json.dump(m, f, ensure_ascii=False, indent=2)

    print(f"### Набор {a.label}: {a.runs} прогона(ов)")
    print(f"- ФУЛ в эталоне: {n}; совпало по всем полям: {full} ({m['full_match_pct']}%)")
    print("- Точность по полям: " + ", ".join(f"{k} {v}%" for k, v in m["field_accuracy_pct"].items()))
    print(f"- Форма Cloud: 12 колонок в порядке реестра — {'да' if form_ok else 'НЕТ'}; строки равны эталону — "
          f"{'да' if rows_equal else 'НЕТ'}; лишних строк {len(extra)}, дублей {dupes}")
    print(f"- Выдуманных значений: {fabricated}")
    print(f"- Обоснование у назначений логиста: {with_reason} из {len(assigned)}")
    print(f"- Результаты прогонов одинаковы: {'да' if m['reproducible'] else 'НЕТ'}")
    print(f"- Время прогонов, с: {m['run_times_sec']}")
    for b in bad:
        print(f"  расхождение: {b}")
    for t in notes:
        print(f"  выдумано: {t}")
    return 0 if (full == n and fabricated == 0 and m["reproducible"] and form_ok and rows_equal) else 1


if __name__ == "__main__":
    sys.exit(main())

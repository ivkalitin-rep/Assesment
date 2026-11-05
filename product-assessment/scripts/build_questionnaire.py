#!/usr/bin/env python3
"""Generate references/questionnaire.json and references/questionnaire.md from the Excel template.

Run this only when the template (assets/assessment_template.xlsx) changes:

    python3 scripts/build_questionnaire.py

The template layout (see the "Вводные данные" sheet, item 7):
  * column A                         - the question;
  * green cells in B/D/F/H/J         - the scoring rule for level 1..5;
  * yellow cell right next to a rule - the score for that rule (C/E/G/I/K);
  * light-blue cell after the last
    score cell of a row              - the comment.
"""
import json
import re
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "assets" / "assessment_template.xlsx"

# Order in which the interview goes. "Рынок" is intentionally excluded.
SECTIONS = [
    {"code": "HYP", "sheet": "Гипотезы", "title": "Работа с гипотезами", "radar_row": 2},
    {"code": "CA", "sheet": "ЦА", "title": "Понимание целевой аудитории (ЦА)", "radar_row": 3},
    {"code": "MET", "sheet": "Метрики", "title": "Влияние на ключевые метрики", "radar_row": 4},
    {"code": "GOAL", "sheet": "Цели и стратегия", "title": "Целеполагание и стратегическое планирование", "radar_row": 5},
]
LEVEL_COLS = {"B": 1, "D": 2, "F": 3, "H": 4, "J": 5}
YELLOW = "FFFFFF00"
NUM_RE = re.compile(r"\[\s*(\d+(?:[.,]\d+)?)\s*балл")


def is_yellow(cell):
    f = cell.fill
    return f.patternType == "solid" and f.fgColor.type == "rgb" and f.fgColor.rgb == YELLOW


def merged_end_row(ws, coord):
    for rng in ws.merged_cells.ranges:
        if coord in rng:
            return rng.max_row
    return ws[coord].row


def clean(text):
    return (text or "").replace("\xa0", " ").strip()


def build():
    wb = openpyxl.load_workbook(TEMPLATE)
    result = {"template": "assets/assessment_template.xlsx", "sections": []}
    for sec in SECTIONS:
        ws = wb[sec["sheet"]]
        # locate the totals rows
        total_row = level_row = None
        for r in range(1, ws.max_row + 1):
            v = clean(ws.cell(r, 1).value)
            if v == "Итого по уровням":
                level_row = r
            elif v == "Итого в целом":
                total_row = r
        levels = {str(lvl): clean(ws[f"{col}2"].value) for col, lvl in LEVEL_COLS.items()}

        blocks = []
        covered_until = 0
        for r in range(4, level_row):
            rules = []
            for col, lvl in LEVEL_COLS.items():
                cell = ws[f"{col}{r}"]
                text = clean(cell.value)
                if not text or not text.startswith("["):
                    continue
                score_col = get_column_letter(column_index_from_string(col) + 1)
                score_cell = f"{score_col}{r}"
                assert is_yellow(ws[score_cell]), f"{sec['sheet']}!{score_cell} is not a score cell"
                options = [float(x.replace(",", ".")) for x in NUM_RE.findall(text)]
                rules.append({
                    "id": f"{sec['code']}-{r}-L{lvl}",
                    "level": lvl,
                    "rule_cell": f"{col}{r}",
                    "score_cell": score_cell,
                    "max": max(options) if options else 0.0,
                    "options": sorted(set(options)),
                    "text": text,
                    "_end": merged_end_row(ws, f"{col}{r}"),
                })
            question = clean(ws.cell(r, 1).value)
            if rules:
                end = max(x.pop("_end") for x in rules)
                last_score = max(column_index_from_string(re.sub(r"\d", "", x["score_cell"])) for x in rules)
                blocks.append({
                    "id": f"{sec['code']}-{r}",
                    "rows": [r, end],
                    "questions": [question] if question else [],
                    "rules": rules,
                    "comment_cell": f"{get_column_letter(last_score + 1)}{r}",
                })
                covered_until = end
            elif question and blocks:
                # sub-question that belongs to the previous rule block (merged rule cell or follow-up)
                blocks[-1]["questions"].append(question)
                if r > covered_until:
                    blocks[-1]["rows"][1] = r
        for b in blocks:
            b["max"] = round(sum(x["max"] for x in b["rules"]), 4)
        per_level = {str(l): round(sum(x["max"] for b in blocks for x in b["rules"] if x["level"] == l), 4)
                     for l in range(1, 6)}
        result["sections"].append({
            **sec,
            "max_total": wb["Радар"][f"B{sec['radar_row']}"].value,
            "rules_max_sum": round(sum(per_level.values()), 4),
            "rules_max_per_level": per_level,
            "level_totals_row": level_row,
            "total_cell": f"C{total_row}",
            "levels": levels,
            "blocks": blocks,
        })
    return result


def to_markdown(data):
    out = ["# Опросник продуктового ассесмента (сгенерировано из шаблона)", "",
           "Файл сгенерирован `scripts/build_questionnaire.py`. Не редактируйте вручную: правьте Excel-шаблон и перегенерируйте.",
           "",
           "Формат: каждый **блок** = вопросы + правила начисления баллов. `id` правила используется как ключ в JSON с оценками.",
           "Баллы внутри одной строки складываются по уровням; варианты внутри одного правила взаимоисключающие (выбирается один).",
           ""]
    for s in data["sections"]:
        out += [f"## {s['sheet']} — {s['title']} (max {s['max_total']} баллов)", ""]
        out += ["### Описание уровней", ""]
        for lvl, txt in s["levels"].items():
            out += [f"**Уровень {lvl}.**", "", txt, ""]
        out += [f"Сумма максимумов правил по уровням: " +
                ", ".join(f"L{k}={v:g}" for k, v in s["rules_max_per_level"].items()) +
                f" (итого {s['rules_max_sum']:g})", ""]
        out += ["### Вопросы и правила", ""]
        for b in s["blocks"]:
            out += [f"#### Блок `{b['id']}` (max {b['max']:g}, комментарий → `{b['comment_cell']}`)", ""]
            for q in b["questions"]:
                out += ["> " + q.replace("\n", "\n> "), ""]
            for r in b["rules"]:
                out += [f"- `{r['id']}` · уровень {r['level']} · max **{r['max']:g}** · ячейка `{r['score_cell']}`  ",
                        f"  {r['text'].replace(chr(10), ' ')}"]
            out += [""]
    return "\n".join(out)


if __name__ == "__main__":
    data = build()
    (ROOT / "references" / "questionnaire.json").write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    (ROOT / "references" / "questionnaire.md").write_text(to_markdown(data), encoding="utf-8")
    for s in data["sections"]:
        print(f"{s['sheet']}: {len(s['blocks'])} блоков, сумма правил {s['rules_max_sum']} / max {s['max_total']}",
              s["rules_max_per_level"])

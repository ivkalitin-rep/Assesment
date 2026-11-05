#!/usr/bin/env python3
"""Fill the product-assessment Excel template from a JSON file with the interview results.

    python3 scripts/fill_assessment.py results.json -o "Ассесмент_Иванов.xlsx" [--report report.md]

The script edits the .xlsx XML directly (not via openpyxl) so the template's formatting
and the radar chart stay exactly as they were. What it writes:

  * score sheets (Гипотезы, ЦА, Метрики, Цели и стратегия): every yellow score cell gets
    a number (0 when the JSON has no score for it), the light-blue comment cell of each
    block gets the comment; the "Итого" formulas keep working and get fresh cached values;
  * "Радар": column E ("Оценка данного PM") becomes a formula pointing at the sheet total,
    columns E/K and the radar chart get fresh cached values, so the chart is correct even
    in viewers that do not recalculate;
  * "Зоны роста": one row per growth zone (A..C) and the overall recommendation in D2.

JSON format: see examples/sample_results.json and SKILL.md.
"""
import argparse
import json
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "assets" / "assessment_template.xlsx"
QUESTIONNAIRE = ROOT / "references" / "questionnaire.json"

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
      "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
      "c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
M = "{%s}" % NS["m"]
EPS = 1e-6


# ----------------------------------------------------------------- cell helpers
def col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def split_ref(ref):
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
    return m.group(1), int(m.group(2))


def get_row(sheet, r, template_row=None):
    data = sheet.find("m:sheetData", NS)
    for row in data.findall("m:row", NS):
        rn = int(row.get("r"))
        if rn == r:
            return row
        if rn > r:
            new = etree.Element(M + "row", r=str(r))
            row.addprevious(new)
            return new
    new = etree.SubElement(data, M + "row", r=str(r))
    if template_row is not None:
        for k, v in template_row.attrib.items():
            if k not in ("r",):
                new.set(k, v)
    return new


def get_cell(sheet, ref, style_from=None):
    col, r = split_ref(ref)
    row = get_row(sheet, r)
    for c in row.findall("m:c", NS):
        ccol, _ = split_ref(c.get("r"))
        if ccol == col:
            return c
        if col_to_num(ccol) > col_to_num(col):
            new = etree.Element(M + "c", r=ref)
            c.addprevious(new)
            break
    else:
        new = etree.SubElement(row, M + "c", r=ref)
    if style_from is not None and style_from.get("s"):
        new.set("s", style_from.get("s"))
    return new


def _reset(cell):
    for child in list(cell):
        cell.remove(child)
    if "t" in cell.attrib:
        del cell.attrib["t"]


def fmt_num(x):
    return repr(round(float(x), 6))


def set_number(sheet, ref, value):
    c = get_cell(sheet, ref)
    _reset(c)
    etree.SubElement(c, M + "v").text = fmt_num(value)


def set_text(sheet, ref, text, style_from=None):
    c = get_cell(sheet, ref, style_from)
    _reset(c)
    if not text:
        return
    c.set("t", "inlineStr")
    is_ = etree.SubElement(c, M + "is")
    t = etree.SubElement(is_, M + "t")
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def set_formula(sheet, ref, formula, cached):
    c = get_cell(sheet, ref)
    _reset(c)
    etree.SubElement(c, M + "f").text = formula
    etree.SubElement(c, M + "v").text = fmt_num(cached)


def set_cached(sheet, ref, cached):
    """Keep the formula, refresh only the cached value."""
    c = get_cell(sheet, ref)
    v = c.find("m:v", NS)
    if v is None:
        v = etree.SubElement(c, M + "v")
    v.text = fmt_num(cached)


def read_number(sheet, ref):
    c = get_cell(sheet, ref)
    v = c.find("m:v", NS)
    return float(v.text) if v is not None and v.text else 0.0


# ----------------------------------------------------------------- workbook
class Workbook:
    def __init__(self, path):
        self.zin = zipfile.ZipFile(path)
        self.parts = {n: self.zin.read(n) for n in self.zin.namelist()}
        self.trees = {}
        wb = self.xml("xl/workbook.xml")
        rels = self.xml("xl/_rels/workbook.xml.rels")
        rid2target = {r.get("Id"): r.get("Target") for r in rels.findall("rel:Relationship", NS)}
        self.sheet_parts = {}
        self.sheet_order = []
        for s in wb.find("m:sheets", NS).findall("m:sheet", NS):
            target = rid2target[s.get("{%s}id" % NS["r"])]
            part = "xl/" + target.lstrip("/").replace("xl/", "", 1) if not target.startswith("/") else target[1:]
            self.sheet_parts[s.get("name")] = part
            self.sheet_order.append(s.get("name"))

    def xml(self, part):
        if part not in self.trees:
            self.trees[part] = etree.fromstring(self.parts[part])
        return self.trees[part]

    def sheet(self, name):
        return self.xml(self.sheet_parts[name])

    def drop_part(self, part):
        self.parts.pop(part, None)
        self.trees.pop(part, None)

    def save(self, path):
        for part, tree in self.trees.items():
            self.parts[part] = etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for name in self.zin.namelist():
                if name in self.parts:
                    z.writestr(name, self.parts[name])


# ----------------------------------------------------------------- main logic
def load_questionnaire():
    return json.loads(QUESTIONNAIRE.read_text(encoding="utf-8"))


def validate(results, q):
    errors, warnings = [], []
    rules = {r["id"]: r for s in q["sections"] for b in s["blocks"] for r in b["rules"]}
    blocks = {b["id"] for s in q["sections"] for b in s["blocks"]}
    for rid, val in (results.get("scores") or {}).items():
        if rid not in rules:
            errors.append(f"Неизвестное правило '{rid}'")
            continue
        try:
            v = float(val)
        except (TypeError, ValueError):
            errors.append(f"{rid}: балл '{val}' не число")
            continue
        if v < -EPS or v > rules[rid]["max"] + EPS:
            errors.append(f"{rid}: балл {v} вне диапазона 0..{rules[rid]['max']}")
    for bid in (results.get("comments") or {}):
        if bid not in blocks:
            errors.append(f"Неизвестный блок комментария '{bid}'")
    missing = [rid for rid in rules if rid not in (results.get("scores") or {})]
    if missing:
        warnings.append(f"Нет оценки для {len(missing)} правил — проставлен 0: " + ", ".join(missing))
    return errors, warnings


def fill(results, out_path, template=TEMPLATE):
    q = load_questionnaire()
    errors, warnings = validate(results, q)
    if errors:
        raise SystemExit("Ошибки во входном JSON:\n  - " + "\n  - ".join(errors))

    wb = Workbook(template)
    scores = {k: float(v) for k, v in (results.get("scores") or {}).items()}
    comments = results.get("comments") or {}
    summary = []

    # 1. score sheets
    for sec in q["sections"]:
        sh = wb.sheet(sec["sheet"])
        level_sum = {c: 0.0 for c in "CEGIK"}
        per_level = {str(l): 0.0 for l in range(1, 6)}
        for b in sec["blocks"]:
            for r in b["rules"]:
                v = scores.get(r["id"], 0.0)
                set_number(sh, r["score_cell"], v)
                level_sum[split_ref(r["score_cell"])[0]] += v
                per_level[str(r["level"])] += v
            text = comments.get(b["id"])
            if text:
                set_text(sh, b["comment_cell"], text)
        lr = sec["level_totals_row"]
        for col, val in level_sum.items():
            set_cached(sh, f"{col}{lr}", val)
        total = sum(level_sum.values())
        set_cached(sh, sec["total_cell"], total)
        summary.append({"sheet": sec["sheet"], "title": sec["title"], "total": round(total, 4),
                        "max": sec["max_total"], "per_level": {k: round(v, 4) for k, v in per_level.items()},
                        "per_level_max": sec["rules_max_per_level"], "radar_row": sec["radar_row"]})

    # 2. radar sheet
    radar = wb.sheet("Радар")
    e_total = k_total = 0.0
    k_values = []
    for s in summary:
        row = s["radar_row"]
        set_formula(radar, f"E{row}", f"'{s['sheet']}'!{next(x for x in q['sections'] if x['sheet'] == s['sheet'])['total_cell']}", s["total"])
        b = read_number(radar, f"B{row}") or s["max"]
        k = s["total"] / b if b else 0.0
        set_cached(radar, f"K{row}", k)
        k_values.append(k)
        e_total += s["total"]
        k_total += k
    set_cached(radar, "E6", e_total)
    set_cached(radar, "K6", k_total)

    # radar chart cache for the "Оценка данного PM" series (Радар!$K$2:$K$5)
    for part in [p for p in wb.parts if p.startswith("xl/charts/chart") and p.endswith(".xml")]:
        chart = wb.xml(part)
        for ser in chart.iter("{%s}ser" % NS["c"]):
            f = ser.find("c:val/c:numRef/c:f", NS)
            if f is not None and f.text and f.text.replace("$", "") == "Радар!K2:K5":
                cache = ser.find("c:val/c:numRef/c:numCache", NS)
                for pt in cache.findall("c:pt", NS):
                    pt.find("c:v", NS).text = fmt_num(k_values[int(pt.get("idx"))])

    # 3. growth zones
    gz = wb.sheet("Зоны роста")
    zones = results.get("growth_zones") or []
    style_row = {c: get_cell(gz, f"{c}2") for c in "ABCD"}
    for i, z in enumerate(zones):
        r = 2 + i
        set_text(gz, f"A{r}", z.get("zone", ""), style_row["A"])
        set_text(gz, f"B{r}", z.get("recommendation", ""), style_row["B"])
        set_text(gz, f"C{r}", z.get("argumentation", ""), style_row["C"])
    if results.get("overall_recommendation"):
        set_text(gz, "D2", results["overall_recommendation"], style_row["D"])
    if zones:
        last = 1 + len(zones)
        dim = gz.find("m:dimension", NS)
        if dim is not None and last > 22:
            dim.set("ref", f"A1:D{last}")

    # 4. workbook-level: recalc on open, open on the radar tab, drop calcChain
    wbx = wb.xml("xl/workbook.xml")
    calc = wbx.find("m:calcPr", NS)
    if calc is not None:
        calc.set("fullCalcOnLoad", "1")
    radar_idx = wb.sheet_order.index("Радар")
    view = wbx.find("m:bookViews/m:workbookView", NS)
    if view is not None:
        view.set("activeTab", str(radar_idx))
        view.attrib.pop("firstSheet", None)
    for name in wb.sheet_order:
        for sv in wb.sheet(name).iter(M + "sheetView"):
            if name == "Радар":
                sv.set("tabSelected", "1")
            else:
                sv.attrib.pop("tabSelected", None)
    if "xl/calcChain.xml" in wb.parts:
        wb.drop_part("xl/calcChain.xml")
        rels = wb.xml("xl/_rels/workbook.xml.rels")
        for r in rels.findall("rel:Relationship", NS):
            if r.get("Target", "").endswith("calcChain.xml"):
                rels.remove(r)
        ct = wb.xml("[Content_Types].xml")
        for o in ct.findall("ct:Override", NS):
            if o.get("PartName") == "/xl/calcChain.xml":
                ct.remove(o)

    wb.save(out_path)
    return summary, warnings


def report_md(results, summary):
    total = sum(s["total"] for s in summary)
    mx = sum(s["max"] for s in summary)
    lines = [f"# Итоги продуктового ассесмента — {results.get('employee', '')}".rstrip(" —"), ""]
    meta = [("Команда", results.get("team")), ("Роль", results.get("role")),
            ("Мини-продукт(ы)", results.get("product")), ("PI", results.get("pi")), ("Дата", results.get("date")),
            ("Ассесор", results.get("assessor"))]
    lines += [f"- **{k}:** {v}" for k, v in meta if v]
    if results.get("pi_goals"):
        lines += ["", "## Цели PI", ""] + [f"{i}. {g}" for i, g in enumerate(results["pi_goals"], 1)]
    lines += ["", "## Баллы", "", "| Срез | Баллы | Максимум | % | L1 | L2 | L3 | L4 | L5 |", "|---|---|---|---|---|---|---|---|---|"]
    for s in summary:
        lv = " | ".join(f"{s['per_level'][k]:g}/{s['per_level_max'][k]:g}" for k in "12345")
        lines.append(f"| {s['title']} | {s['total']:g} | {s['max']} | {100 * s['total'] / s['max']:.0f}% | {lv} |")
    lines.append(f"| **Итого** | **{total:g}** | **{mx}** | **{100 * total / mx:.0f}%** | | | | | |")
    lines += ["", "Ориентиры из «Вводных данных»: 3-й уровень по всем срезам ≈ 39 баллов, 5-й уровень ≈ 77 баллов."]
    if results.get("growth_zones"):
        lines += ["", "## Зоны роста", ""]
        for z in results["growth_zones"]:
            lines += [f"### {z.get('zone', '')}", "", f"**Рекомендация.** {z.get('recommendation', '')}", "",
                      f"**Аргументация.** {z.get('argumentation', '')}", ""]
    if results.get("overall_recommendation"):
        lines += ["## Общая рекомендация", "", results["overall_recommendation"], ""]
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("results", help="JSON с результатами интервью")
    ap.add_argument("-o", "--output", required=True, help="куда сохранить заполненный .xlsx")
    ap.add_argument("--report", help="дополнительно сохранить краткий отчет в Markdown")
    ap.add_argument("--template", default=str(TEMPLATE))
    args = ap.parse_args()

    results = json.loads(Path(args.results).read_text(encoding="utf-8"))
    summary, warnings = fill(results, args.output, args.template)
    for w in warnings:
        print("ВНИМАНИЕ:", w, file=sys.stderr)
    for s in summary:
        print(f"{s['sheet']}: {s['total']:g} / {s['max']}")
    print(f"Итого: {sum(s['total'] for s in summary):g} / {sum(s['max'] for s in summary)}")
    print("Файл:", args.output)
    if args.report:
        Path(args.report).write_text(report_md(results, summary), encoding="utf-8")
        print("Отчет:", args.report)


if __name__ == "__main__":
    main()

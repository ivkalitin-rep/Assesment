#!/usr/bin/env python3
"""Regression tests for all three skills.

    python3 tests/regression.py

Needs Python 3 with lxml and openpyxl. LibreOffice (soffice) is optional: if present, the filled workbook is
also recalculated by it to compare formula results with the cached values written by fill_assessment.py.
Exit code 1 if any test fails.
"""
import filecmp
import importlib.util
import re
import shutil
import subprocess
import sys
import tempfile
import traceback
import zipfile
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parent.parent
SINGLE = REPO / "single-skill" / "product-assessment"
INTERVIEW = REPO / "two-skills" / "product-assessment-interview"
SCORING = REPO / "two-skills" / "product-assessment-scoring"
TEMPLATE = REPO / "template" / "assessment_template.xlsx"
SKILLS = [SINGLE, INTERVIEW, SCORING]
EXPECTED_MAX = {"Гипотезы": 25, "ЦА": 16, "Метрики": 24, "Цели и стратегия": 12}
YELLOW = "FFFFFF00"

TMP = Path(tempfile.mkdtemp(prefix="assessment-regression-"))
results = []


def test(fn):
    name = fn.__name__
    try:
        fn()
        results.append((name, True, ""))
        print(f"PASS  {name}")
    except AssertionError as e:
        results.append((name, False, str(e)))
        print(f"FAIL  {name}: {e}")
    except Exception:
        results.append((name, False, traceback.format_exc(limit=2)))
        print(f"ERROR {name}:\n{traceback.format_exc(limit=3)}")
    return fn


def load(path, name):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def run(*args, cwd=None):
    p = subprocess.run([sys.executable, *map(str, args)], capture_output=True, text=True, cwd=cwd)
    return p.returncode, p.stdout + p.stderr


fill_s = load(SCORING / "scripts" / "fill_assessment.py", "fill_scoring")
fill_1 = load(SINGLE / "scripts" / "fill_assessment.py", "fill_single")
check = load(INTERVIEW / "scripts" / "check_protocol.py", "check_protocol")
ART = INTERVIEW / "scripts" / "artifacts.py"


# ------------------------------------------------------------------ A. generator and generated files
@test
def a1_generator_is_idempotent():
    gen = [SINGLE / "references" / "questionnaire.md", SCORING / "references" / "questionnaire.md",
           INTERVIEW / "references" / "questions.md", SINGLE / "assets" / "assessment_template.xlsx",
           SCORING / "assets" / "assessment_template.xlsx"]
    before = {p: p.read_bytes() for p in gen}
    code, out = run(REPO / "tools" / "build_questionnaire.py")
    assert code == 0, out
    changed = [str(p.relative_to(REPO)) for p in gen if p.read_bytes() != before[p]]
    assert not changed, f"генератор изменил закоммиченные файлы: {changed}"


@test
def a2_rule_sums_match_methodology_maxima():
    q = fill_s.load_questionnaire()
    for s in q["sections"]:
        total = sum(r["max"] for b in s["blocks"] for r in b["rules"])
        assert abs(total - EXPECTED_MAX[s["sheet"]]) < 0.1, f"{s['sheet']}: сумма правил {total}"
        assert s["max_total"] == EXPECTED_MAX[s["sheet"]], s
    assert [s["sheet"] for s in q["sections"]] == list(EXPECTED_MAX), "порядок срезов"


@test
def a3_interview_questions_reveal_no_scoring():
    text = (INTERVIEW / "references" / "questions.md").read_text(encoding="utf-8")
    for bad in ["балл", "max", "ячейка", "Уровень 1", "Уровень 5", "Описание уровней", "-L1", "-L2"]:
        assert bad not in text, f"в questions.md есть «{bad}»"
    for path in INTERVIEW.rglob("*"):
        if path.is_file() and path.suffix in (".md", ".py"):
            t = path.read_text(encoding="utf-8")
            assert "questionnaire.md" not in t and "methodology.md" not in t, f"{path.name} ссылается на файлы оценки"
    assert not list(INTERVIEW.rglob("*.xlsx")), "в интервьюере есть Excel-шаблон"


@test
def a4_interview_and_scoring_ids_match():
    q = fill_s.load_questionnaire()
    sid = [b["id"] for s in q["sections"] for b in s["blocks"]]
    iq = check.load_questions()
    iid = [b["id"] for s in iq["sections"] for b in s["blocks"]]
    assert sid == iid, "id блоков у интервьюера и оценщика расходятся"
    assert len(sid) == 78, len(sid)
    na = sum(b["needs_artifact"] for s in iq["sections"] for b in s["blocks"])
    assert na == 21, f"пометок [артефакт]: {na}"


@test
def a5_both_questionnaires_identical():
    assert filecmp.cmp(SINGLE / "references" / "questionnaire.md", SCORING / "references" / "questionnaire.md", shallow=False)


# ------------------------------------------------------------------ B. template
@test
def b1_template_copies_identical():
    for p in [SINGLE / "assets" / "assessment_template.xlsx", SCORING / "assets" / "assessment_template.xlsx"]:
        assert filecmp.cmp(TEMPLATE, p, shallow=False), p


@test
def b2_template_fixes_present():
    wb = openpyxl.load_workbook(TEMPLATE)
    m = wb["Метрики"]
    assert m["A2"].value == "Работа с метриками", m["A2"].value
    assert str(m["A31"].value).startswith("5.3."), m["A31"].value
    assert "0,11 балла, если в работе используются до 30%" in m["H22"].value
    assert wb["ЦА"]["A2"].value == "Понимание целевой аудитории (ЦА)"
    assert wb["Радар"]._charts, "нет диаграммы на «Радаре»"


@test
def b3_questionnaire_cells_match_template_colours():
    wb = openpyxl.load_workbook(TEMPLATE)
    q = fill_s.load_questionnaire()
    for s in q["sections"]:
        ws = wb[s["sheet"]]
        assert ws.cell(s["level_totals_row"], 1).value == "Итого по уровням", s["sheet"]
        assert str(ws[s["total_cell"]].value).startswith("=SUM"), s["total_cell"]
        assert wb["Радар"][f"A{s['radar_row']}"].value, s["radar_row"]
        for b in s["blocks"]:
            for r in b["rules"]:
                f = ws[r["score_cell"]].fill
                assert f.fgColor.rgb == YELLOW, f"{r['id']} {r['score_cell']} не жёлтая"
            f = ws[b["comment_cell"]].fill
            assert f.patternType == "solid" and f.fgColor.type == "theme" and f.fgColor.theme == 8, \
                f"{b['id']} {b['comment_cell']} не голубая"


# ------------------------------------------------------------------ C. fill_assessment
def sample(skill):
    return skill / "examples" / "sample_results.md"


@test
def c1_fill_scripts_identical():
    assert filecmp.cmp(SINGLE / "scripts" / "fill_assessment.py", SCORING / "scripts" / "fill_assessment.py", shallow=False)


@test
def c2_samples_fill_without_warnings():
    for skill, mod in [(SCORING, fill_s), (SINGLE, fill_1)]:
        out = TMP / f"{skill.name}.xlsx"
        code, log = run(skill / "scripts" / "fill_assessment.py", sample(skill), "-o", out, "--report", TMP / f"{skill.name}.md")
        assert code == 0, log
        assert "ВНИМАНИЕ" not in log, f"{skill.name}: {log}"
        assert "Итого: 33.38 / 77" in log, log


@test
def c3_filled_workbook_contents():
    out = TMP / "c3.xlsx"
    res = fill_s.load_results(sample(SCORING))
    summary, _ = fill_s.fill(res, out)
    wb = openpyxl.load_workbook(out, data_only=True)
    q = fill_s.load_questionnaire()
    scores = {k: fill_s.num(v) for k, v in res["scores"].items()}
    for s in q["sections"]:
        ws = wb[s["sheet"]]
        for b in s["blocks"]:
            for r in b["rules"]:
                assert abs((ws[r["score_cell"]].value or 0) - scores.get(r["id"], 0)) < 1e-9, r["id"]
            if b["id"] in res["comments"]:
                assert ws[b["comment_cell"]].value == res["comments"][b["id"]], b["id"]
        total = sum(scores.get(r["id"], 0) for b in s["blocks"] for r in b["rules"])
        assert abs(ws[s["total_cell"]].value - total) < 1e-6, (s["sheet"], ws[s["total_cell"]].value, total)
        radar = wb["Радар"]
        row = s["radar_row"]
        assert abs(radar[f"E{row}"].value - total) < 1e-6, ("Радар E", row)
        assert abs(radar[f"K{row}"].value - total / radar[f"B{row}"].value) < 1e-6, ("Радар K", row)
    gz = wb["Зоны роста"]
    for i, z in enumerate(res["growth_zones"], 2):
        assert gz[f"A{i}"].value == z["zone"] and gz[f"B{i}"].value == z["recommendation"] and gz[f"C{i}"].value == z["argumentation"]
    assert gz["D2"].value == res["overall_recommendation"]
    f = openpyxl.load_workbook(out)["Радар"]["E2"].value
    assert f == "='Гипотезы'!C25", f
    assert openpyxl.load_workbook(out).active.title == "Радар"


@test
def c4_chart_cache_and_package():
    out = TMP / "c3.xlsx"
    with zipfile.ZipFile(out) as z:
        names = z.namelist()
        assert "xl/calcChain.xml" not in names
        assert any(n.startswith("xl/charts/chart") for n in names), "диаграмма пропала"
        chart = z.read("xl/charts/chart1.xml").decode("utf-8")
        wbxml = z.read("xl/workbook.xml").decode("utf-8")
    assert 'fullCalcOnLoad="1"' in wbxml
    m = re.search(r"Радар!\$K\$2:\$K\$5</c:f><c:numCache>(.*?)</c:numCache>", chart, re.S)
    assert m, "нет ряда K2:K5 в диаграмме"
    vals = [float(v) for v in re.findall(r"<c:v>([^<]+)</c:v>", m.group(1))]
    wb = openpyxl.load_workbook(out, data_only=True)
    expect = [wb["Радар"][f"K{r}"].value for r in range(2, 6)]
    assert all(abs(a - b) < 1e-6 for a, b in zip(vals, expect)), (vals, expect)


@test
def c5_libreoffice_recalculation_matches_cache():
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("      (LibreOffice не найден — пропущено)")
        return
    src = TMP / "c3.xlsx"
    outdir = TMP / "lo"
    outdir.mkdir(exist_ok=True)
    subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(outdir), str(src)],
                   capture_output=True, timeout=180)
    rec = outdir / "c3.xlsx"
    assert rec.exists(), "LibreOffice не пересчитал файл"
    a = openpyxl.load_workbook(src, data_only=True)
    b = openpyxl.load_workbook(rec, data_only=True)
    for sheet, cells in [("Гипотезы", ["C24", "E24", "G24", "I24", "K24", "C25"]), ("ЦА", ["C31"]), ("Метрики", ["C34"]),
                         ("Цели и стратегия", ["C25"]), ("Радар", ["E2", "E3", "E4", "E5", "E6", "K2", "K6"])]:
        for c in cells:
            va, vb = a[sheet][c].value or 0, b[sheet][c].value or 0
            assert abs(va - vb) < 1e-6, f"{sheet}!{c}: кэш {va}, пересчёт {vb}"


@test
def c6_validation_errors():
    base = sample(SCORING).read_text(encoding="utf-8")
    cases = {
        "unknown_rule": base.replace("- HYP-4-L1: 1", "- HYP-4-L1: 1\n- HYP-99-L2: 1"),
        "over_max": base.replace("- HYP-4-L1: 1", "- HYP-4-L1: 5"),
        "unknown_block": base.replace("## Оценки\n", "## Оценки\n\n### XYZ-1\n\n**Комментарий:** test\n"),
    }
    for name, text in cases.items():
        p = TMP / f"bad_{name}.md"
        p.write_text(text, encoding="utf-8")
        code, log = run(SCORING / "scripts" / "fill_assessment.py", p, "-o", TMP / "bad.xlsx")
        assert code != 0 and "Ошибки в файле результатов" in log, f"{name}: {log}"


@test
def c7_comma_decimal_and_missing_warning():
    base = sample(SCORING).read_text(encoding="utf-8")
    p = TMP / "comma.md"
    p.write_text(base.replace("- CA-7-L2: 0.5", "- CA-7-L2: 0,5"), encoding="utf-8")
    code, log = run(SCORING / "scripts" / "fill_assessment.py", p, "-o", TMP / "comma.xlsx")
    assert code == 0 and "Итого: 33.38 / 77" in log and "ВНИМАНИЕ" not in log, log
    p.write_text(base.replace("- MET-31-L5: 0\n", ""), encoding="utf-8")
    code, log = run(SCORING / "scripts" / "fill_assessment.py", p, "-o", TMP / "miss.xlsx")
    assert code == 0 and "ВНИМАНИЕ" in log and "MET-31-L5" in log, log


@test
def c8_working_file_with_progress_and_artifacts():
    base = sample(SINGLE).read_text(encoding="utf-8")
    text = base.replace("## Цели PI", "## Прогресс\n\n**Статус:** в процессе\n\n**Следующий вопрос:** CA-4\n\n"
                                      "**Ожидаемые артефакты:** HYP-6: отчёт\n\n## Цели PI\n\nСкриншот: `PI_1.png` — список")
    text = text.replace("### HYP-6\n", "### HYP-6\n\n**Артефакты:**\n- `HYP-6_1.png` — отчёт\n", 1)
    p = TMP / "working.md"
    p.write_text(text, encoding="utf-8")
    res = fill_1.load_results(p)
    assert len(res["pi_goals"]) == 4, res["pi_goals"]
    assert res["scores"]["HYP-6-L2"] == "2", res["scores"].get("HYP-6-L2")
    code, log = run(SINGLE / "scripts" / "fill_assessment.py", p, "-o", TMP / "working.xlsx")
    assert code == 0 and "Итого: 33.38 / 77" in log and "ВНИМАНИЕ" not in log, log


@test
def c9_more_than_21_growth_zones():
    base = sample(SCORING).read_text(encoding="utf-8")
    zones = "\n".join(f"### Зона {i}\n\n**Рекомендация:** r{i}\n\n**Аргументация:** a{i}\n" for i in range(1, 26))
    text = re.sub(r"## Зоны роста\n.*?## Общая рекомендация", "## Зоны роста\n\n" + zones + "\n## Общая рекомендация", base, flags=re.S)
    p = TMP / "zones.md"
    p.write_text(text, encoding="utf-8")
    code, log = run(SCORING / "scripts" / "fill_assessment.py", p, "-o", TMP / "zones.xlsx")
    assert code == 0, log
    gz = openpyxl.load_workbook(TMP / "zones.xlsx")["Зоны роста"]
    assert gz["A26"].value == "Зона 25" and gz["C26"].value == "a25", (gz["A26"].value, gz["C26"].value)


# ------------------------------------------------------------------ D. artifacts.py
@test
def d1_artifact_scripts_identical():
    assert filecmp.cmp(SINGLE / "scripts" / "artifacts.py", ART, shallow=False)


@test
def d2_artifacts_save_pack_unpack():
    w = TMP / "art"
    (w / "uploads").mkdir(parents=True)
    for f in ["goals.PNG", "rep.pdf", "shot.png", "t.xlsx"]:
        (w / "uploads" / f).write_text(f)
    steps = [(["save", "uploads/goals.PNG", "PI"], "PI_1.png"), (["save", "rep.pdf", "hyp-6"], "HYP-6_1.pdf"),
             (["save", "shot.png", "HYP-6"], "HYP-6_2.png"), (["save", "t.xlsx", "MET-12"], "MET-12_1.xlsx")]
    for args, expect in steps:
        code, out = run(ART, *args, "--dir", "a", cwd=w)
        assert code == 0 and out.strip() == expect, (args, out)
    code, out = run(ART, "save", "shot.png", "XX-1", "--dir", "a", cwd=w)
    assert code != 0 and "Неверный id" in out, out
    code, out = run(ART, "save", "nothere.png", "CA-4", "--dir", "a", cwd=w)
    assert code != 0 and "Файл не найден" in out, out
    code, out = run(ART, "pack", "--dir", "a", "-o", "a.zip", cwd=w)
    assert code == 0 and "4 файл" in out, out
    code, out = run(ART, "unpack", "a.zip", "--dir", "b", cwd=w)
    assert code == 0 and sorted(p.name for p in (w / "b").iterdir()) == ["HYP-6_1.pdf", "HYP-6_2.png", "MET-12_1.xlsx", "PI_1.png"], out
    code, out = run(ART, "save", "shot.png", "HYP-6", "--dir", "b", cwd=w)
    assert out.strip() == "HYP-6_3.png", f"нумерация после распаковки: {out}"


# ------------------------------------------------------------------ E. check_protocol.py
SAMPLE_PROTOCOL = INTERVIEW / "examples" / "sample_protocol.md"


def protocol_problems(text, art=None, partial=False):
    return check.check(text, art, partial)


@test
def e1_sample_protocol_complete():
    assert protocol_problems(SAMPLE_PROTOCOL.read_text(encoding="utf-8")) == []
    assert filecmp.cmp(SAMPLE_PROTOCOL, SCORING / "examples" / "sample_protocol.md", shallow=False)


@test
def e2_protocol_negative_cases():
    t = SAMPLE_PROTOCOL.read_text(encoding="utf-8")
    p = protocol_problems(t.replace("- **Команда:** Связи и интеграции (тест)", "- **Команда:** "))
    assert any("Команда" in x for x in p), p
    p = protocol_problems(re.sub(r"### CA-7\n.*?(?=### CA-8)", "", t, flags=re.S))
    assert any("CA-7: нет блока" in x for x in p), p
    p = protocol_problems(t.replace("**Причина пропуска:** MET-15: производственные метрики не прогнозируются", ""))
    assert any("MET-16" in x and "причина" in x for x in p), p
    p = protocol_problems(re.sub(r"(### HYP-19\n\n\*\*Статус:\*\* ответ.*?)\*\*Артефакт не приложен:\*\*[^\n]*\n", r"\1", t, flags=re.S))
    assert any("HYP-19" in x and "артефакт" in x for x in p), p
    p = protocol_problems(t + "\n### ZZZ-1\n\n**Статус:** ответ\n\n**Ответ:** x\n")
    assert any("неизвестный id" in x for x in p), p


@test
def e3_protocol_partial_mode():
    t = SAMPLE_PROTOCOL.read_text(encoding="utf-8")
    cut = t[:t.index("## ЦА")] + "## Заметки интервьюера\n\n—\n"
    assert any("нет блока" in x for x in protocol_problems(cut))
    assert protocol_problems(cut, partial=True) == []


@test
def e4_protocol_artifacts_folder():
    t = SAMPLE_PROTOCOL.read_text(encoding="utf-8")
    d = TMP / "pa"
    d.mkdir()
    for f in ["HYP-6_1.png", "HYP-6_2.png"]:
        (d / f).write_text("x")
    assert protocol_problems(t, d) == [], protocol_problems(t, d)
    (d / "MET-8_1.png").write_text("x")
    assert any("не упомянут" in x for x in protocol_problems(t, d))
    xref = t.replace("**Артефакт не приложен:** Подтверждение", "**Артефакты:**\n- `HYP-6_1.png` — ссылка на файл из HYP-6\n\n**Артефакт не приложен:** Подтверждение")
    (d / "MET-8_1.png").unlink()
    assert protocol_problems(xref, d) == [] or all("MET-12" not in x for x in protocol_problems(xref, d))
    bad = t.replace("`HYP-6_2.png`", "`HYP-6_9.png`")
    assert any("HYP-6_9.png" in x for x in protocol_problems(bad, d))
    (d / "PI_1.png").write_text("x")
    assert any("PI_1.png" in x for x in protocol_problems(t, d))
    with_pi = t.replace("## Цели PI\n", "## Цели PI\n\nСкриншот: `PI_1.png` — список целей\n")
    assert protocol_problems(with_pi, d) == [], protocol_problems(with_pi, d)


# ------------------------------------------------------------------ F. documentation consistency
def skill_text(skill):
    return (skill / "SKILL.md").read_text(encoding="utf-8")


@test
def f1_frontmatter_and_names():
    for skill in SKILLS:
        t = skill_text(skill)
        m = re.match(r"^---\nname: ([a-z0-9-]+)\ndescription: (.+?)\n---\n", t, re.S)
        assert m, f"{skill.name}: нет frontmatter"
        assert m.group(1) == skill.name, (m.group(1), skill.name)
        assert len(m.group(2)) <= 1024, f"{skill.name}: description длиннее 1024"


@test
def f2_files_referenced_in_skill_exist():
    for skill in SKILLS:
        t = skill_text(skill)
        for ref in set(re.findall(r"`((?:references|scripts|assets|examples)/[^`\s]+)`", t)):
            if "<" in ref:
                continue
            assert (skill / ref).exists(), f"{skill.name}: нет файла {ref}"


@test
def f3_no_json_anywhere():
    assert not [p for p in REPO.rglob("*.json") if ".git" not in p.parts], "есть .json файлы"
    for p in list(REPO.rglob("*.md")) + list(REPO.rglob("*.py")):
        if ".git" in p.parts or p.name == "regression.py":
            continue
        t = p.read_text(encoding="utf-8")
        for m in re.finditer(r"json|JSON", t):
            ctx = t[max(0, m.start() - 40):m.end() + 40]
            assert "не используется" in ctx or "только Markdown" in ctx, f"{p.relative_to(REPO)}: «{ctx}»"


@test
def f4_dependency_maps_use_known_ids():
    ids = {b["id"] for s in fill_s.load_questionnaire()["sections"] for b in s["blocks"]}
    for skill in [SINGLE, INTERVIEW]:
        t = skill_text(skill)
        part = t[t.index("Карта зависимых вопросов"):]
        part = part[:part.index("\n\n", part.index("|---|"))]
        for ref in re.findall(r"`((?:HYP|CA|MET|GOAL)-\d+)`", part):
            assert ref in ids, f"{skill.name}: в карте зависимостей неизвестный {ref}"
        for rng in re.findall(r"`((?:HYP|CA|MET|GOAL)-\d+)`…`((?:HYP|CA|MET|GOAL)-\d+)`", part):
            assert rng[0] in ids and rng[1] in ids, rng


@test
def f5_core_rules_present_in_interviewing_skills():
    must = ["Строго один вопрос", "Проверка перед отправкой", "Каждый вопрос должен получить ответ", "мини-продукт",
            "Отвечайте только про свой мини-продукт", "название команды", "**все** мини-продукты", "перерыв",
            "artifacts.py save", "PI_1.png", "не сохранён"]
    for skill in [SINGLE, INTERVIEW]:
        t = skill_text(skill)
        for m in must:
            if skill == SINGLE and m == "не сохранён":
                m = "файлом не сохранён"
            assert m in t, f"{skill.name}: нет «{m}»"


@test
def f6_methodology_copies_consistent():
    import difflib
    a = (SINGLE / "references" / "methodology.md").read_text(encoding="utf-8").splitlines()
    b = (SCORING / "references" / "methodology.md").read_text(encoding="utf-8").splitlines()
    changed = [l for l in difflib.unified_diff(a, b, lineterm="", n=0)
               if l[:1] in "+-" and not l.startswith(("+++", "---"))]
    # единственное допустимое расхождение: откуда берётся список целей PI (интервью vs протокол)
    unexpected = [l for l in changed if "Проценты по целям считаются" not in l and "по всем мини-продуктам вместе" not in l]
    assert not unexpected, f"методички расходятся: {unexpected[:3]}"
    for t in (a, b):
        joined = "\n".join(t)
        for m in ["Разные ячейки суммируются", "засчитывает и все рубежи младших уровней", "### 3.1. Как считать балл"]:
            assert m in joined, f"нет «{m}»"


@test
def f7_readme_paths_exist():
    t = (REPO / "README.md").read_text(encoding="utf-8")
    for ref in set(re.findall(r"`((?:single-skill|two-skills|template|tools)/[^`\s]+)`", t)):
        assert (REPO / ref).exists(), f"README: нет {ref}"


# ------------------------------------------------------------------ G. packaging
@test
def g1_skill_zips():
    for skill in SKILLS:
        z = TMP / f"{skill.name}.zip"
        with zipfile.ZipFile(z, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in skill.rglob("*"):
                if p.is_file() and "__pycache__" not in p.parts:
                    zf.write(p, f"{skill.name}/{p.relative_to(skill)}")
        with zipfile.ZipFile(z) as zf:
            assert f"{skill.name}/SKILL.md" in zf.namelist()
        assert z.stat().st_size < 10 * 1024 * 1024


if __name__ == "__main__":
    failed = [r for r in results if not r[1]]
    print(f"\n{len(results) - len(failed)} из {len(results)} тестов прошли.")
    shutil.rmtree(TMP, ignore_errors=True)
    sys.exit(1 if failed else 0)

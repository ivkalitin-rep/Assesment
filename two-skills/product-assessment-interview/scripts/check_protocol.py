#!/usr/bin/env python3
"""Check an interview protocol for completeness, render it to Markdown and pack it for the assessor.

    python3 scripts/check_protocol.py protocol.json [--pack protocol.zip]

Checks:
  * every question block from references/questions.json is present and is either "answered" (with an answer)
    or "skipped" (with a reason);
  * blocks that need an artifact either have attached artifacts or an artifact_note explaining why not;
  * attached artifact files exist next to the protocol.

Writes <protocol>.md next to the JSON. With --pack, builds a zip with the JSON, the Markdown and the artifact files.
Exit code 1 if something is missing (the Markdown and the zip are still written, so a break can be saved).
"""
import argparse
import json
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
QUESTIONS = ROOT / "references" / "questions.json"


def load(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def check(protocol, base_dir):
    q = load(QUESTIONS)
    blocks = protocol.get("blocks") or {}
    problems = []
    for key in ("employee", "team", "role", "pi"):
        if not protocol.get(key):
            problems.append(f"не заполнено поле «{key}»")
    if not protocol.get("mini_products"):
        problems.append("не перечислены мини-продукты")
    if not protocol.get("pi_goals"):
        problems.append("нет списка целей PI")
    known = set()
    for s in q["sections"]:
        for b in s["blocks"]:
            known.add(b["id"])
            rec = blocks.get(b["id"])
            if not rec:
                problems.append(f"{b['id']}: нет ответа")
                continue
            st = rec.get("status")
            if st == "skipped":
                if not rec.get("skip_reason"):
                    problems.append(f"{b['id']}: пропущен без причины")
                continue
            if st != "answered" or not str(rec.get("answer", "")).strip():
                problems.append(f"{b['id']}: нет ответа")
                continue
            arts = rec.get("artifacts") or []
            if b["needs_artifact"] and not arts and not rec.get("artifact_note"):
                problems.append(f"{b['id']}: нужен артефакт или пояснение, почему его нет")
            for a in arts:
                if not a.get("description"):
                    problems.append(f"{b['id']}: у артефакта {a.get('file', '?')} нет описания")
                f = a.get("file")
                if f and not (base_dir / f).exists():
                    problems.append(f"{b['id']}: файл артефакта не найден: {f}")
    for extra in set(blocks) - known:
        problems.append(f"неизвестный id блока: {extra}")
    return problems


def render(protocol):
    q = load(QUESTIONS)
    p = protocol
    out = [f"# Протокол интервью — {p.get('employee', '')}", ""]
    for k, label in (("team", "Команда"), ("role", "Роль"), ("pi", "PI"), ("date", "Дата"), ("status", "Статус")):
        if p.get(k):
            out.append(f"- **{label}:** {p[k]}")
    if p.get("mini_products"):
        out += ["", "## Мини-продукты", ""]
        out += [f"- **{m.get('name', '')}** — {m.get('description', '')}" for m in p["mini_products"]]
    if p.get("pi_goals"):
        out += ["", "## Цели PI", ""]
        out += [f"{g.get('n', i)}. {g.get('text', '')} *({g.get('mini_product', '')})*"
                for i, g in enumerate(p["pi_goals"], 1)]
    blocks = p.get("blocks") or {}
    for s in q["sections"]:
        out += ["", f"## {s['sheet']} — {s['title']}", ""]
        for b in s["blocks"]:
            rec = blocks.get(b["id"]) or {}
            out += [f"### {b['id']}", ""]
            if rec.get("status") == "skipped":
                out += [f"*Не задавался: {rec.get('skip_reason', '')}*", ""]
                continue
            if not rec:
                out += ["*Нет ответа*", ""]
                continue
            out += [f"**Вопрос:** {rec.get('asked', '')}", "", f"**Ответ:** {rec.get('answer', '')}", ""]
            for f in rec.get("followups") or []:
                out += [f"- **{f.get('q', '')}** — {f.get('a', '')}"]
            if rec.get("followups"):
                out.append("")
            for a in rec.get("artifacts") or []:
                out += [f"- Артефакт `{a.get('file', '')}`: {a.get('description', '')}"]
            if rec.get("artifact_note"):
                out += [f"- Артефакт не приложен: {rec['artifact_note']}"]
            out.append("")
    if p.get("notes"):
        out += ["## Заметки интервьюера", "", p["notes"], ""]
    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("protocol")
    ap.add_argument("--pack", help="путь к zip-архиву для ассесора")
    args = ap.parse_args()

    path = Path(args.protocol)
    protocol = load(path)
    base = path.parent
    problems = check(protocol, base)
    md_path = path.with_suffix(".md")
    md_path.write_text(render(protocol), encoding="utf-8")
    print("Читаемая версия:", md_path)
    if args.pack:
        with zipfile.ZipFile(args.pack, "w", zipfile.ZIP_DEFLATED) as z:
            z.write(path, path.name)
            z.write(md_path, md_path.name)
            for rec in (protocol.get("blocks") or {}).values():
                for a in rec.get("artifacts") or []:
                    f = a.get("file")
                    if f and (base / f).exists():
                        z.write(base / f, f)
        print("Архив:", args.pack)
    if problems:
        print(f"\nНе хватает ({len(problems)}):", *problems, sep="\n  - ")
        sys.exit(1)
    print("Протокол полный.")


if __name__ == "__main__":
    main()

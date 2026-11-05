#!/usr/bin/env python3
"""Check a Markdown interview protocol for completeness.

    python3 scripts/check_protocol.py protocol_<фамилия>.md [--artifacts artifacts_<фамилия>]

The protocol format is described in SKILL.md («Формат протокола»). Checks:
  * header fields (ФИО, Команда, Роль, PI) are filled, mini-products and PI goals are listed;
  * every question id from references/questions.json has a block «### <id>»;
  * each block is either «ответ» (with a non-empty «Ответ») or «не задавался» (with «Причина пропуска»);
  * blocks that need an artifact list artifacts or explain in «Артефакт не приложен» why there are none;
  * every listed artifact has a description;
  * with --artifacts: every saved artifact named in the protocol exists in the folder and is named after its
    question (HYP-6_1.png under «### HYP-6»), and every file in the folder is mentioned in the protocol.
    An artifact that could not be saved is listed as «- `не сохранён` — описание».

Exit code 1 and a list of problems if something is missing.
"""
import argparse
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
QUESTIONS = ROOT / "references" / "questions.json"

HEADER_RE = re.compile(r"^- \*\*(.+?):\*\*\s*(.*)$")
FIELD_RE = re.compile(r"^\*\*(.+?):\*\*\s*(.*)$")
ARTIFACT_RE = re.compile(r"^- `(.+?)`\s*[—-]\s*(.*)$")


def parse(text):
    header, sections, blocks = {}, {}, {}
    section = block = field = None
    for raw in text.splitlines():
        line = raw.rstrip()
        if line.startswith("## "):
            section, block, field = line[3:].strip(), None, None
            sections.setdefault(section, [])
            continue
        if line.startswith("### "):
            block, field = line[4:].strip().strip("`").split()[0], None
            blocks[block] = {}
            continue
        if block:
            m = FIELD_RE.match(line)
            if m:
                field = m.group(1).strip()
                blocks[block][field] = m.group(2).strip()
            elif field and line.strip():
                blocks[block][field] = (blocks[block][field] + "\n" + line.strip()).strip()
            continue
        if section is None:
            m = HEADER_RE.match(line)
            if m:
                header[m.group(1).strip()] = m.group(2).strip()
        elif line.strip():
            sections[section].append(line.strip())
    return header, sections, blocks


def check(text, artifacts_dir=None):
    header, sections, blocks = parse(text)
    q = json.loads(QUESTIONS.read_text(encoding="utf-8"))
    problems = []
    for key in ("ФИО", "Команда", "Роль", "PI", "Статус"):
        if not header.get(key):
            problems.append(f"не заполнено поле «{key}»")
    if not sections.get("Мини-продукты"):
        problems.append("не перечислены мини-продукты (раздел «## Мини-продукты»)")
    if not sections.get("Цели PI"):
        problems.append("нет списка целей PI (раздел «## Цели PI»)")
    known, mentioned = set(), set()
    for s in q["sections"]:
        for b in s["blocks"]:
            known.add(b["id"])
            rec = blocks.get(b["id"])
            if rec is None:
                problems.append(f"{b['id']}: нет блока")
                continue
            status = rec.get("Статус", "").lower()
            if status.startswith("не задавался"):
                if not rec.get("Причина пропуска"):
                    problems.append(f"{b['id']}: не задавался, но причина не указана")
                continue
            if not status.startswith("ответ") or not rec.get("Ответ"):
                problems.append(f"{b['id']}: нет ответа")
                continue
            arts = [ARTIFACT_RE.match(x) for x in rec.get("Артефакты", "").splitlines() if x.startswith("- ")]
            if b["needs_artifact"] and not arts and not rec.get("Артефакт не приложен"):
                problems.append(f"{b['id']}: нужен артефакт или пояснение «Артефакт не приложен»")
            for m in arts:
                if m is None or not m.group(2).strip():
                    problems.append(f"{b['id']}: у артефакта нет описания (формат: - `файл` — описание)")
                    continue
                name = m.group(1).strip()
                mentioned.add(name)
                if artifacts_dir is None or name == "не сохранён":
                    continue
                if not name.startswith(b["id"] + "_"):
                    problems.append(f"{b['id']}: файл {name} назван не по номеру вопроса")
                if not (artifacts_dir / name).is_file():
                    problems.append(f"{b['id']}: файла {name} нет в папке {artifacts_dir}")
    for extra in sorted(set(blocks) - known):
        problems.append(f"неизвестный id блока: {extra}")
    if artifacts_dir is not None and artifacts_dir.is_dir():
        for f in sorted(p.name for p in artifacts_dir.iterdir() if p.is_file()):
            if f not in mentioned:
                problems.append(f"файл {f} есть в папке артефактов, но не упомянут в протоколе")
    return problems


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("protocol")
    ap.add_argument("--artifacts", help="папка с сохранёнными артефактами")
    args = ap.parse_args()
    art = Path(args.artifacts) if args.artifacts else None
    problems = check(Path(args.protocol).read_text(encoding="utf-8"), art)
    if problems:
        print(f"Не хватает ({len(problems)}):", *problems, sep="\n  - ")
        sys.exit(1)
    print("Протокол полный.")


if __name__ == "__main__":
    main()

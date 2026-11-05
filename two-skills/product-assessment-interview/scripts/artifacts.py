#!/usr/bin/env python3
"""Save interview artifacts into a folder, named after the question id, and pack/unpack the folder.

Save a file the employee attached (numbering is automatic: HYP-6_1.png, HYP-6_2.pdf, ...):

    python3 scripts/artifacts.py save <file> <question-id> --dir artifacts_<фамилия>

<question-id> is a question id (HYP-6, CA-20, ...) or PI for the PI goals list (screenshot of goals -> PI_1.png).

If <file> is just a name and is not found as given, it is searched in the usual upload folders
(/mnt/user-data/uploads, /mnt/data, /mnt/user-data, the current folder, the home folder).
Prints the new file name to put into the protocol.

Pack the folder for the employee (at a break and at the end), unpack it when resuming in a new chat:

    python3 scripts/artifacts.py pack --dir artifacts_<фамилия> -o artifacts_<фамилия>.zip
    python3 scripts/artifacts.py unpack artifacts_<фамилия>.zip --dir artifacts_<фамилия>

List what is saved:

    python3 scripts/artifacts.py list --dir artifacts_<фамилия>
"""
import argparse
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path

ID_RE = re.compile(r"^((HYP|CA|MET|GOAL)-\d+|PI)$")
SEARCH_DIRS = ["/mnt/user-data/uploads", "/mnt/data", "/mnt/user-data", ".", str(Path.home())]


def find(src):
    p = Path(src)
    if p.is_file():
        return p
    name = p.name
    for d in SEARCH_DIRS:
        base = Path(d)
        if not base.is_dir():
            continue
        direct = base / name
        if direct.is_file():
            return direct
        for root, _dirs, files in os.walk(base):
            if root.count(os.sep) - str(base).count(os.sep) > 3:
                continue
            if name in files:
                return Path(root) / name
    return None


def cmd_save(args):
    qid = args.question_id.upper()
    if not ID_RE.match(qid):
        sys.exit(f"Неверный id вопроса: {args.question_id} (ожидается, например, HYP-6, CA-20, MET-8, GOAL-9 или PI для списка целей PI)")
    src = find(args.file)
    if src is None:
        sys.exit(f"Файл не найден: {args.file}. Попросите сотрудника приложить его файлом, а не вставкой в сообщение.")
    out = Path(args.dir)
    out.mkdir(parents=True, exist_ok=True)
    n = 1
    while any(out.glob(f"{qid}_{n}.*")) or (out / f"{qid}_{n}").exists():
        n += 1
    dest = out / f"{qid}_{n}{src.suffix.lower()}"
    shutil.copyfile(src, dest)
    print(dest.name)


def cmd_pack(args):
    out = Path(args.dir)
    files = sorted(p for p in out.glob("*") if p.is_file()) if out.is_dir() else []
    with zipfile.ZipFile(args.output, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(f, f"{out.name}/{f.name}")
    print(f"{args.output}: {len(files)} файл(ов)")


def cmd_unpack(args):
    out = Path(args.dir)
    out.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(args.archive) as z:
        for info in z.infolist():
            if info.is_dir():
                continue
            name = Path(info.filename).name
            if not name or name.startswith("."):
                continue
            with z.open(info) as src, open(out / name, "wb") as dst:
                shutil.copyfileobj(src, dst)
    cmd_list(args)


def cmd_list(args):
    out = Path(args.dir)
    files = sorted(p.name for p in out.glob("*") if p.is_file()) if out.is_dir() else []
    print(f"{out}: {len(files)} файл(ов)")
    for f in files:
        print("  " + f)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("save")
    s.add_argument("file")
    s.add_argument("question_id")
    s.add_argument("--dir", required=True)
    s.set_defaults(func=cmd_save)
    p = sub.add_parser("pack")
    p.add_argument("--dir", required=True)
    p.add_argument("-o", "--output", required=True)
    p.set_defaults(func=cmd_pack)
    u = sub.add_parser("unpack")
    u.add_argument("archive")
    u.add_argument("--dir", required=True)
    u.set_defaults(func=cmd_unpack)
    ls = sub.add_parser("list")
    ls.add_argument("--dir", required=True)
    ls.set_defaults(func=cmd_list)
    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()

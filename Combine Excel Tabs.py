"""
Combines selected sheets from the Excel files in this folder into one new workbook, in the order typed.

Author: Yeoh Huey Sun
Email:  greyw901@gmail.com
GitHub: https://github.com/greyw901
"""

from __future__ import annotations

import os
import re
import sys
import time
import traceback
from dataclasses import dataclass, field
from itertools import product

import win32com.client

# ---------------------------------------------------------------- constants
BANNER        = "Combine Excel Tabs"
WIDTH         = 56
EXCEL_EXTS    = {".xls", ".xlsx", ".xlsm", ".xlsb"}
LOCK_PREFIX   = "~$"                        # Excel's lock files - never real inputs
# Deviation from naming convention (run date on a deliverable): Grey's choice.
OUTPUT_NAME   = "Combined_{date}.xlsx"
LOG_NAME      = "error log_{date}.txt"
PLACEHOLDER   = "zz~placeholder~zz"         # blank sheet the copies are inserted before
TEMP_PREFIX   = "~tmp~"                     # sheet names used briefly while relinking
MAX_SHEET_LEN = 31
NO_PASSWORD   = "~no~password~"             # protected file raises instead of prompting

XL_OPENXML_WORKBOOK = 51
XL_WBAT_WORKSHEET   = -4167
XL_WORKSHEET        = -4167
XL_SHEET_VISIBLE    = -1
XL_CALC_MANUAL      = -4135
XL_CALC_AUTOMATIC   = -4105
XL_LINK_EXCEL       = 1
MSO_MACROS_DISABLED = 3                     # msoAutomationSecurityForceDisable

LINKS_KEEP   = "keep"
LINKS_VALUES = "values"

# 'Sheet name'!A1  or  Sheet1!A1   (external '[Book]Sheet' refs are skipped)
REF_RE = re.compile(r"'((?:[^']|'')+)'!|(?<![\w.\]'])([\w.]+)!")


class UserExit(Exception):
    """[0] on the first screen - close immediately."""


class UserCancelled(Exception):
    """[00] - stop and close."""


class GoBack(Exception):
    """[0] on a later screen - redraw the previous one."""


# ---------------------------------------------------------------- data
@dataclass
class SheetInfo:
    code: str
    name: str
    refs: set[str] = field(default_factory=set)   # other sheets its formulas use


@dataclass(eq=False)          # compared/hashed by identity - one object per open file
class Source:
    prefix: str
    path: str
    wb: object
    sheets: list[SheetInfo]
    hidden: set[str]

    @property
    def fname(self) -> str:
        return os.path.basename(self.path)


@dataclass
class Pick:
    source: Source
    sheet: SheetInfo
    final: str = ""


# ---------------------------------------------------------------- helpers
def get_script_dir() -> str:
    return os.path.dirname(os.path.abspath(
        sys.executable if getattr(sys, "frozen", False) else __file__
    ))


def get_excel_files(folder: str) -> list[str]:
    return [
        os.path.join(folder, f) for f in sorted(os.listdir(folder), key=str.lower)
        if os.path.isfile(os.path.join(folder, f))
        and not f.startswith(LOCK_PREFIX)
        and os.path.splitext(f)[1].lower() in EXCEL_EXTS
    ]


def get_unique_path(folder: str, name: str) -> str:
    candidate = os.path.join(folder, name)
    base, ext = os.path.splitext(name)
    counter = 1
    while os.path.exists(candidate):
        candidate = os.path.join(folder, f"{base}_{counter}{ext}")
        counter += 1
    return candidate


def prefix_generator():
    letters = "abcdefghijklmnopqrstuvwxyz"
    length = 1
    while True:
        for combo in product(letters, repeat=length):
            yield "".join(combo)
        length += 1


def as_rows(block: object) -> tuple:
    """Range.Value / .Formula returns a bare value for one cell, else rows."""
    return block if isinstance(block, tuple) else ((block,),)


def find_sheet_refs(formulas: object, sheet_names: list[str], own: str) -> set[str]:
    """Names of other sheets in the same workbook that the formulas point to."""
    lookup = {n.lower(): n for n in sheet_names}
    found = set()
    for row in as_rows(formulas):
        for cell in row:
            if not (isinstance(cell, str) and cell.startswith("=") and "!" in cell):
                continue
            for quoted, plain in REF_RE.findall(cell):
                name = quoted.replace("''", "'") if quoted else plain
                if name.startswith("["):
                    continue
                actual = lookup.get(name.lower())
                if actual and actual.lower() != own.lower():
                    found.add(actual)
    return found


def read_source(xl, path: str, prefix: str) -> Source:
    wb = xl.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True, Password=NO_PASSWORD)
    try:
        all_names, visible, hidden, formulas = [], [], set(), {}
        for sh in wb.Sheets:
            all_names.append(sh.Name)
            if sh.Visible == XL_SHEET_VISIBLE:
                visible.append(sh.Name)
                if sh.Type == XL_WORKSHEET:
                    formulas[sh.Name] = sh.UsedRange.Formula
            else:
                hidden.add(sh.Name)
        sheets = [
            SheetInfo(f"{prefix}{i}", name,
                      find_sheet_refs(formulas.get(name, ""), all_names, name))
            for i, name in enumerate(visible, start=1)
        ]
    except Exception:
        wb.Close(SaveChanges=False)
        raise
    return Source(prefix, path, wb, sheets, hidden)


def parse_files(raw: str, prefixes: list[str]) -> list[str]:
    """File letters (a, c) or letter ranges (a-c), typed order, duplicates dropped."""
    selected: list[str] = []
    for token in (t for t in re.split(r"[\s,]+", raw.strip().lower()) if t):
        if m := re.fullmatch(r"([a-z]+)-([a-z]+)", token):
            for p in m.groups():
                if p not in prefixes:
                    raise ValueError(f"file '{p}' not found")
            i, j = prefixes.index(m.group(1)), prefixes.index(m.group(2))
            if i > j:
                raise ValueError(f"range '{token}' must be ascending")
            letters = prefixes[i:j + 1]
        elif token in prefixes:
            letters = [token]
        else:
            raise ValueError(f"'{token}' not found")
        selected += [p for p in letters if p not in selected]
    return selected


def parse_sheets(raw: str, code_map: dict[str, Pick]) -> list[str]:
    """Sheet codes in typed order, duplicates dropped. Raises ValueError with the reason."""
    selected: list[str] = []
    for token in (t for t in re.split(r"[\s,]+", raw.strip().lower()) if t):
        if re.fullmatch(r"[a-z]+", token):
            codes = [c for c in code_map if re.fullmatch(rf"{token}\d+", c)]
            if not codes:
                raise ValueError(f"file '{token}' not found")
        elif m := re.fullmatch(r"([a-z]+)(\d+)-([a-z]+)?(\d+)", token):
            p1, n1, p2, n2 = m.group(1), int(m.group(2)), m.group(3) or m.group(1), int(m.group(4))
            if p1 != p2:
                raise ValueError(f"range '{token}' spans two files")
            if n1 > n2:
                raise ValueError(f"range '{token}' must be ascending")
            codes = [f"{p1}{n}" for n in range(n1, n2 + 1)]
            missing = [c for c in codes if c not in code_map]
            if missing:
                raise ValueError(f"range '{token}' - '{missing[0]}' not found")
        elif token in code_map:
            codes = [token]
        else:
            raise ValueError(f"'{token}' not found")
        selected += [c for c in codes if c not in selected]
    return selected


def unique_sheet_name(taken: set[str], desired: str) -> str:
    """Excel sheet names are case-insensitive and max 31 chars. Repeats get _1, _2 ..."""
    if desired.lower() not in taken:
        return desired
    counter = 1
    while True:
        suffix = f"_{counter}"
        candidate = desired[:MAX_SHEET_LEN - len(suffix)] + suffix
        if candidate.lower() not in taken:
            return candidate
        counter += 1


def assign_final_names(picks: list[Pick]) -> None:
    taken = {PLACEHOLDER.lower()}
    for p in picks:
        p.final = unique_sheet_name(taken, p.sheet.name)
        taken.add(p.final.lower())


def missing_refs(p: Pick, picks: list[Pick]) -> list[str]:
    """Sheets this pick's formulas use that are not also picked from the same file."""
    picked = {q.sheet.name for q in picks if q.source is p.source}
    return sorted(r for r in p.sheet.refs if r not in picked)


def link_warnings(picks: list[Pick]) -> list[tuple[Pick, str]]:
    out = []
    for p in picks:
        for ref in missing_refs(p, picks):
            if ref in p.source.hidden:
                where = "hidden sheet"
            else:
                where = next(s.code for s in p.source.sheets if s.name == ref)
            out.append((p, f'uses "{ref}" ({where}), not selected'))
    return out


def label(p: Pick) -> str:
    return f'"{p.source.fname}"!{p.sheet.name}'


# ---------------------------------------------------------------- screens
def cls() -> None:
    os.system("cls")


def banner() -> None:
    print("=" * WIDTH)
    print(BANNER)
    print("=" * WIDTH)
    print()


def status_slot(status: str) -> None:
    if status:
        print(status)
        print()


def options(*lines: str) -> None:
    print()
    print("-" * WIDTH)
    for line in lines:
        print(line)
    print()


def print_files(names: list[str]) -> None:
    print(f"Files: {len(names)}")
    for i, name in enumerate(names, 1):
        print(f'   {i}. "{name}"')
    print()


def file_screen(files: list[tuple[str, str]]) -> list[str]:
    """files = [(prefix, path)]. Returns chosen prefixes."""
    prefixes = [p for p, _ in files]
    status = ""
    while True:
        cls()
        banner()
        status_slot(status)
        print("Select FILES:")
        print("e.g. a, c   or   a-c")
        print()
        for prefix, path in files:
            print(f'   [{prefix}] "{os.path.basename(path)}"')
        options("[Enter] Select all", "[#] Select", "[0] Exit")
        raw = input("Files: ").strip()
        if raw == "0":
            raise UserExit
        if not raw:
            return prefixes
        try:
            return parse_files(raw, prefixes)
        except ValueError as e:
            status = f"Error: Invalid Input - {e}"


def sheet_screen(sources: list[Source], unreadable: list[dict],
                 code_map: dict[str, Pick]) -> list[str]:
    status = ""
    while True:
        cls()
        banner()
        print_files([s.fname for s in sources] + [u["item"] for u in unreadable])
        status_slot(status)
        print("Select SHEETS:")
        print("Type codes in the order wanted, e.g. a3-4, b12, c12   or   a, b")
        for src in sources:
            print()
            print(f'   [{src.prefix}] "{src.fname}"')
            for sh in src.sheets:
                print(f"      [{sh.code}] {sh.name}")
        if unreadable:
            print()
            print(f"Unreadable files: {len(unreadable)}")
            for i, f in enumerate(unreadable, 1):
                print(f'   {i}. "{f["item"]}"   :   {f["error"]}')
        if not code_map:
            options("[0] Back")
            raw = input("Sheets: ").strip()
            if raw == "0":
                raise GoBack
            status = "Error: Invalid Input - no readable sheets in these files"
            continue
        options("[Enter] Select all", "[#] Select", "[0] Back")
        raw = input("Sheets: ").strip()
        if raw == "0":
            raise GoBack
        if not raw:
            return list(code_map)
        try:
            return parse_sheets(raw, code_map)
        except ValueError as e:
            status = f"Error: Invalid Input - {e}"


def confirm_screen(sources: list[Source], picks: list[Pick],
                   warns: list[tuple[Pick, str]], out_name: str) -> str:
    """Returns the link handling for sheets that use an unselected sheet."""
    status = ""
    while True:
        cls()
        banner()
        print_files([s.fname for s in sources])
        status_slot(status)
        print(f"Sheets: {len(picks)}")
        for i, p in enumerate(picks, 1):
            renamed = f'   ->   "{p.final}"' if p.final != p.sheet.name else ""
            print(f"   {i}. {label(p)}{renamed}")
        if warns:
            print()
            print(f"Linked to unselected sheets: {len(warns)}")
            for i, (p, msg) in enumerate(warns, 1):
                print(f"   {i}. {label(p)}   :   {msg}")
        print()
        print(f'Output : "{out_name}"')
        if warns:
            # A choice is required here, so no [Enter] default.
            options("[1] Keep the link to the source file",
                    "[2] Turn those formulas into plain values",
                    "[0] Back", "[00] Cancel & exit")
        else:
            options("[Enter] Continue", "[0] Back", "[00] Cancel & exit")
        raw = input("Action: ").strip()
        if raw == "00":
            raise UserCancelled
        if raw == "0":
            raise GoBack
        if warns and raw == "1":
            return LINKS_KEEP
        if warns and raw == "2":
            return LINKS_VALUES
        if not warns and raw == "":
            return LINKS_KEEP
        status = "Error: Invalid Input"


# ---------------------------------------------------------------- work
def copy_sheet(out_wb, before, p: Pick) -> None:
    p.source.wb.Sheets(p.sheet.name).Copy(Before=before)
    copied = out_wb.Sheets(before.Index - 1)
    if copied.Name != p.final:
        copied.Name = p.final


def freeze_missing_refs(out_wb, src: Source, p: Pick, missing: list[str]) -> int:
    """Turn formulas that point at a missing sheet of src into their current value."""
    ws = out_wb.Sheets(p.final)
    pattern = re.compile(
        r"\[" + re.escape(src.fname) + r"\](?:" + "|".join(re.escape(m) for m in missing) + r")'?!",
        re.IGNORECASE)
    ur = ws.UsedRange
    r0, c0 = ur.Row, ur.Column
    count = 0
    for i, row in enumerate(as_rows(ur.Formula)):
        for j, f in enumerate(row):
            if isinstance(f, str) and f.startswith("=") and pattern.search(f):
                cell = ws.Cells(r0 + i, c0 + j)
                cell.Value = cell.Value
                count += 1
    return count


def relink_to_self(out_wb, src: Source, copies: list[Pick]) -> None:
    """Point formulas that still reference src at the copied sheets instead.

    ChangeLink to the output itself resolves by sheet NAME, so the copies take
    their source names for the moment of the switch (anything holding such a
    name steps aside), then everything is renamed back. Renames update formulas.
    """
    originals = {p.sheet.name.lower() for p in copies}
    finals = {p.final.lower() for p in copies}
    blockers = [s.Name for s in out_wb.Sheets
                if s.Name.lower() in originals and s.Name.lower() not in finals]
    for i, p in enumerate(copies):
        out_wb.Sheets(p.final).Name = f"{TEMP_PREFIX}c{i}"
    for i, name in enumerate(blockers):
        out_wb.Sheets(name).Name = f"{TEMP_PREFIX}b{i}"
    for i, p in enumerate(copies):
        out_wb.Sheets(f"{TEMP_PREFIX}c{i}").Name = p.sheet.name
    try:
        out_wb.ChangeLink(src.wb.FullName, out_wb.FullName, XL_LINK_EXCEL)
    finally:
        for i, p in enumerate(copies):
            out_wb.Sheets(p.sheet.name).Name = f"{TEMP_PREFIX}c{i}"
        for i, p in enumerate(copies):
            out_wb.Sheets(f"{TEMP_PREFIX}c{i}").Name = p.final
        for i, name in enumerate(blockers):
            out_wb.Sheets(f"{TEMP_PREFIX}b{i}").Name = name


def failure(item: str, step: str, e: Exception, context: str = "") -> dict:
    return {"item": item, "step": step, "error": f"{type(e).__name__} - {e}",
            "context": context, "trace": traceback.format_exc()}


def write_log(folder: str, failures: list[dict]) -> str:
    path = get_unique_path(folder, LOG_NAME.format(date=time.strftime("%Y%m%d")))
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"{BANNER} - error log\n")
        f.write(f"Input folder: {folder}\nOutput folder: {folder}\n\n")
        for fl in failures:
            f.write(f'"{fl["item"]}"\n')
            f.write(f'  Step: {fl["step"]}\n')
            f.write(f'  Error: {fl["error"]}\n')
            if fl["context"]:
                f.write(f'  Context: {fl["context"]}\n')
            f.write("  Traceback:\n")
            for line in fl["trace"].rstrip().splitlines():
                f.write(f"    {line}\n")
            f.write("\n")
    return path


def run(xl, folder: str, picks: list[Pick], link_mode: str, failures: list[dict]) -> None:
    start = time.time()
    cls()
    banner()
    used = list(dict.fromkeys(p.source for p in picks))
    print(f"Inputs: {len(used)}")
    for i, src in enumerate(used, 1):
        print(f'   {i}. "{src.fname}"')
    print()
    print(f"Sheets : {len(picks)}")
    if link_warnings(picks):
        print("Links to unselected sheets : "
              + ("kept" if link_mode == LINKS_KEEP else "turned into values"))
    print()
    print("Processing")

    print("Copying sheets ...", flush=True)
    out_wb = xl.Workbooks.Add(XL_WBAT_WORKSHEET)
    outputs: list[str] = []
    try:
        placeholder = out_wb.Sheets(1)
        placeholder.Name = PLACEHOLDER
        copied: list[Pick] = []
        for i, p in enumerate(picks, 1):
            try:
                copy_sheet(out_wb, placeholder, p)
                copied.append(p)
                renamed = f'   ->   "{p.final}"' if p.final != p.sheet.name else ""
                print(f"   {i}. {label(p)}{renamed}", flush=True)
            except Exception as e:
                failures.append(failure(label(p), "copying sheet", e,
                                        f"target name = {p.final!r}"))
                print(f"   {i}. {label(p)}   :   FAILED", flush=True)

        if copied:
            out_path = get_unique_path(
                folder, OUTPUT_NAME.format(date=time.strftime("%Y%m%d")))
            print("Saving ...", flush=True)
            placeholder.Delete()
            out_wb.SaveAs(out_path, FileFormat=XL_OPENXML_WORKBOOK)

            print("Relinking formulas ...", flush=True)
            frozen = relinked = 0
            for src in used:
                mine = [p for p in copied if p.source is src]
                gaps = {p.final: missing_refs(p, copied) for p in mine}
                if any(gaps.values()):
                    if link_mode == LINKS_KEEP:
                        continue                 # whole file stays linked to source
                    for p in mine:
                        if gaps[p.final]:
                            try:
                                frozen += freeze_missing_refs(out_wb, src, p, gaps[p.final])
                            except Exception as e:
                                failures.append(failure(label(p), "turning formulas into values",
                                                        e, f"missing sheets = {gaps[p.final]}"))
                links = {s.lower() for s in (out_wb.LinkSources(XL_LINK_EXCEL) or ())}
                if src.wb.FullName.lower() not in links:
                    continue
                try:
                    relink_to_self(out_wb, src, mine)
                    relinked += 1
                except Exception as e:
                    failures.append(failure(src.fname, "relinking formulas", e,
                                            f"sheets = {[p.sheet.name for p in mine]}"))
            kept = out_wb.LinkSources(XL_LINK_EXCEL) or ()
            print(f"   Files relinked : {relinked}")
            if link_mode == LINKS_VALUES:
                print(f"   Formulas turned into values : {frozen}")
            print(f"   External links kept : {len(kept)}")
            for k in kept:
                print(f'      "{os.path.basename(k)}"')

            xl.Calculation = XL_CALC_AUTOMATIC   # saved with the file - must be automatic
            out_wb.Sheets(1).Activate()
            out_wb.Save()
            outputs.append(os.path.basename(out_path))
    finally:
        out_wb.Close(SaveChanges=False)

    print()
    print(f"Outputs: {len(outputs)}")
    for i, name in enumerate(outputs, 1):
        print(f'   {i}. "{name}"')
    if failures:
        print()
        print(f"Failed: {len(failures)}")
        for i, fl in enumerate(failures, 1):
            print(f'   {i}. "{fl["item"]}"   :   {fl["error"]}')
    print()
    print(f"Elapsed: {time.time() - start:.1f}s")


def open_sources(xl, files: list[tuple[str, str]], chosen: list[str],
                 cache: dict[str, Source], bad: dict[str, dict]) -> None:
    """Open each chosen file once; later visits reuse it."""
    todo = [(p, path) for p, path in files if p in chosen and p not in cache and p not in bad]
    if not todo:
        return
    cls()
    banner()
    print("Reading files ...", flush=True)
    for prefix, path in todo:
        try:
            cache[prefix] = read_source(xl, path, prefix)
        except Exception as e:
            bad[prefix] = failure(os.path.basename(path), "opening file", e)
            continue
        if len(cache) == 1:
            xl.ScreenUpdating = False            # only valid once a workbook is open
            xl.Calculation = XL_CALC_MANUAL


def main() -> None:
    folder = get_script_dir()
    os.chdir(folder)

    paths = get_excel_files(folder)
    if not paths:
        cls()
        banner()
        print("Error: No Excel files found in this folder.")
        options("[Enter] to close...")
        input()
        return
    files = list(zip(prefix_generator(), paths))

    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = MSO_MACROS_DISABLED
    cache: dict[str, Source] = {}
    bad: dict[str, dict] = {}
    failures: list[dict] = []
    try:
        while True:
            chosen = file_screen(files)
            open_sources(xl, files, chosen, cache, bad)
            sources = [cache[p] for p in chosen if p in cache]
            unreadable = [bad[p] for p in chosen if p in bad]
            code_map = {sh.code: Pick(src, sh) for src in sources for sh in src.sheets}
            try:
                while True:
                    codes = sheet_screen(sources, unreadable, code_map)
                    picks = [Pick(code_map[c].source, code_map[c].sheet) for c in codes]
                    assign_final_names(picks)
                    out_name = os.path.basename(get_unique_path(
                        folder, OUTPUT_NAME.format(date=time.strftime("%Y%m%d"))))
                    try:
                        link_mode = confirm_screen(sources, picks, link_warnings(picks), out_name)
                        break
                    except GoBack:
                        continue
                break
            except GoBack:
                continue

        failures += unreadable
        run(xl, folder, picks, link_mode, failures)
    finally:
        if cache:
            try:
                xl.Calculation = XL_CALC_AUTOMATIC
                xl.ScreenUpdating = True
            except Exception:
                pass
        for src in cache.values():
            try:
                src.wb.Close(SaveChanges=False)
            except Exception:
                pass
        xl.Quit()

    options("[Enter] to close...", *(["[0] Save error log & close"] if failures else []))
    if input("Action: ").strip() == "0" and failures:
        path = write_log(folder, failures)
        print(f'Saved: "{os.path.basename(path)}"')
        time.sleep(1.5)


if __name__ == "__main__":
    try:
        main()
    except UserExit:
        pass
    except UserCancelled:
        print()
        print("Cancelled.")
        print()
        input("[Enter] to close...")
    except Exception as e:
        print()
        print("Error: An unexpected error occurred.")
        print(f"   {type(e).__name__} - {e}")
        print()
        traceback.print_exc()
        print()
        input("[Enter] to close...")

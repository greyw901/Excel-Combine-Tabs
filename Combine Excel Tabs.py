"""
Combine Excel Tabs.py
---------------------
Lists all tabs from every Excel file in the folder,
lets you select and order them, then copies into a new workbook.

Uses win32com — requires Microsoft Excel installed.
Output: Combined_YYYYMMDD.xlsx
"""

import os
import sys
import re
import datetime
from itertools import product
import win32com.client

EXCEL_EXTS = {".xls", ".xlsx", ".xlsm", ".xlsb"}


def get_script_dir():
    return os.path.dirname(os.path.abspath(
        sys.executable if getattr(sys, 'frozen', False) else __file__
    ))


def get_excel_files(folder):
    files = []
    for fname in sorted(os.listdir(folder)):
        fpath = os.path.join(folder, fname)
        if os.path.isfile(fpath) and os.path.splitext(fname)[1].lower() in EXCEL_EXTS:
            files.append(fpath)
    return files


def get_unique_path(folder, name):
    candidate = os.path.join(folder, name)
    if not os.path.exists(candidate):
        return candidate
    base, ext = os.path.splitext(name)
    counter = 1
    while True:
        candidate = os.path.join(folder, f"{base}_{counter}{ext}")
        if not os.path.exists(candidate):
            return candidate
        counter += 1


def prefix_generator():
    letters = "abcdefghijklmnopqrstuvwxyz"
    length = 1
    while True:
        for combo in product(letters, repeat=length):
            yield "".join(combo)
        length += 1


def get_visible_sheets(xl_app, fpath):
    wb = xl_app.Workbooks.Open(fpath, ReadOnly=True, UpdateLinks=False)
    try:
        names = []
        for sheet in wb.Sheets:
            if sheet.Visible == -1:
                names.append(sheet.Name)
        return names
    finally:
        wb.Close(SaveChanges=False)


def parse_selection(raw, code_map):
    """Returns (selected_codes, invalid_tokens)."""
    tokens = re.split(r'[\s,]+', raw.strip().lower())
    selected = []
    invalid  = []
    seen     = set()

    for token in tokens:
        if not token:
            continue

        if re.match(r'^[a-z]+$', token):
            matches = sorted(
                [c for c in code_map if re.match(rf'^{re.escape(token)}\d+$', c)],
                key=lambda c: int(re.search(r'\d+', c).group())
            )
            if matches:
                for code in matches:
                    if code not in seen:
                        selected.append(code)
                        seen.add(code)
            else:
                invalid.append(token)
            continue

        range_match = re.match(r'^([a-z]+)(\d+)-([a-z]+)?(\d+)$', token)
        if range_match:
            p1 = range_match.group(1)
            n1 = int(range_match.group(2))
            p2 = range_match.group(3) or p1
            n2 = int(range_match.group(4))
            if p1 == p2:
                range_valid = True
                for n in range(n1, n2 + 1):
                    code = f"{p1}{n}"
                    if code in code_map:
                        if code not in seen:
                            selected.append(code)
                            seen.add(code)
                    else:
                        range_valid = False
                if not range_valid:
                    invalid.append(token)
            else:
                invalid.append(token)
            continue

        if token in code_map:
            if token not in seen:
                selected.append(token)
                seen.add(token)
        else:
            invalid.append(token)

    return selected, invalid


def unique_sheet_name(existing_names, desired):
    if desired not in existing_names:
        return desired
    counter = 2
    while True:
        candidate = f"{desired}_{counter}"
        if candidate not in existing_names:
            return candidate
        counter += 1


def main():
    script_dir = get_script_dir()
    os.chdir(script_dir)

    print("=" * 50)
    print("  Combine Excel Tabs")
    print("=" * 50)
    print("Selects sheets from multiple Excel files and combines them into one workbook.")
    print()

    excel_files = get_excel_files(script_dir)
    if not excel_files:
        print("No Excel files found in this folder.")
        return

    print("Reading workbooks...\n")
    xl_app = win32com.client.Dispatch("Excel.Application")
    xl_app.Visible       = False
    xl_app.DisplayAlerts = False

    try:
        while True:
            code_map      = {}
            display_lines = []
            gen           = prefix_generator()

            for fpath in excel_files:
                fname  = os.path.basename(fpath)
                prefix = next(gen)
                try:
                    sheets = get_visible_sheets(xl_app, fpath)
                except Exception as e:
                    print(f"Unreadable Excel file: \"{fname}\"")
                    print("[Enter] - Skip & Continue")
                    print("[0] - Cancel & Return to Menu")
                    choice = input().strip()
                    if choice == "0":
                        return
                    continue

                display_lines.append(f"  {prefix} - {fname}")
                for i, sheet in enumerate(sheets, start=1):
                    code = f"{prefix}{i}"
                    code_map[code] = (fpath, sheet)
                    display_lines.append(f"    {code} - {sheet}")

            if not code_map:
                print("No readable Excel files found.")
                return

            print("\n".join(display_lines))
            print()

            print("Enter selection (e.g. a1, b1-b4, a3)")
            print("[Enter] - Select all sheets from all files")
            print("[0] - Cancel & Return to Menu")
            raw = input().strip().lower()

            if raw == "0":
                return

            if not raw:
                selected_codes = list(code_map.keys())
                invalid_tokens = []
            else:
                selected_codes, invalid_tokens = parse_selection(raw, code_map)

            if invalid_tokens:
                quoted = ", ".join(f'"{t}"' for t in invalid_tokens)
                print(f"\nInvalid Input: {quoted}")
                print("[Enter] - Skip Invalid Input & Continue")
                print("[N] - Re-enter")
                print("[0] - Cancel & Return to Menu")
                choice = input().strip().lower()
                if choice == "0":
                    continue
                if choice == "n":
                    print()
                    continue

            if not selected_codes:
                print("No valid sheets selected. Aborted.")
                break

            print(f"\nSelection summary ({len(selected_codes)} sheet(s)):")
            prev_file = None
            for code in selected_codes:
                fpath, sheet = code_map[code]
                fname = os.path.basename(fpath)
                if fname != prev_file:
                    print(f"  {fname}")
                    prev_file = fname
                print(f"    {code} - {sheet}")

            print()
            print("[Enter] - Proceed")
            print("[N] - Re-enter")
            print("[0] - Cancel & Return to Menu")
            proceed = input().strip().lower()
            if proceed == "0":
                continue
            if proceed == "n":
                print()
                continue

            date_str = datetime.datetime.now().strftime("%Y%m%d")
            out_name = f"Combined_{date_str}.xlsx"
            out_path = get_unique_path(script_dir, out_name)

            out_wb = xl_app.Workbooks.Add()
            while out_wb.Sheets.Count > 1:
                out_wb.Sheets(out_wb.Sheets.Count).Delete()
            dummy          = out_wb.Sheets(1)
            dummy.Name     = "_placeholder_"
            existing_names = ["_placeholder_"]
            open_wbs       = {}

            print("\nCopying sheets...")
            for code in selected_codes:
                fpath, sheet_name = code_map[code]
                if fpath not in open_wbs:
                    open_wbs[fpath] = xl_app.Workbooks.Open(
                        fpath, ReadOnly=True, UpdateLinks=False
                    )
                src_wb    = open_wbs[fpath]
                src_sheet = src_wb.Sheets(sheet_name)
                final_name = unique_sheet_name(existing_names, sheet_name)
                existing_names.append(final_name)
                src_sheet.Copy(Before=dummy)
                out_wb.Sheets(out_wb.Sheets.Count - 1).Name = final_name
                print(f"  {code} - {sheet_name}" +
                      (f"  -> renamed to {final_name}" if final_name != sheet_name else ""))

            dummy.Delete()
            for wb in open_wbs.values():
                wb.Close(SaveChanges=False)

            out_wb.SaveAs(out_path, FileFormat=51)
            out_wb.Close(SaveChanges=False)
            print(f"\nSaved: {os.path.basename(out_path)}")
            print("\nDone.")
            break

    finally:
        xl_app.Quit()


try:
    main()
except Exception as e:
    print("ERROR: An unexpected error occurred.")
    print(f"Error Details: {e}")

print("Enter to Close.")
input()

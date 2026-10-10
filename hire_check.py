"""
Hire report checker - ACS in-house hire report vs supplier hire reports.

Usage:
    python3 hire_check.py INHOUSE.xlsx SUPPLIER1.xlsx [SUPPLIER2.xlsx ...] [--code-col C]

For each supplier report it writes SUPPLIER_checked.xlsx where:
  - lines whose plant code is on any sheet of the in-house report -> green + bold
  - lines NOT on the in-house report are left untouched and listed on screen
It also writes INHOUSE_checked.xlsx: ACS lines charged on the supplier reports -> green + bold
"""
import sys, os, re, subprocess, tempfile, warnings
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import column_index_from_string, get_column_letter

warnings.filterwarnings("ignore")

GREEN = PatternFill("solid", start_color="C6EFCE", end_color="C6EFCE")
NOT_CODES = {"CODE", "PLANTCODE", "PLANT", "ITEM", "ITEMCODE", "PRODUCT", "PRODUCTCODE",
             "REF", "REFERENCE", "STOCK", "STOCKCODE", "DESCRIPTION", "QTY", "QUANTITY", "DAYS", "RATE", "CHARGE", "AMOUNT", "TOTAL", "SUBTOTAL", "VAT", "NET", "GROSS"}


def norm(v):
    """Tidy a cell into a comparable code: uppercase, no spaces."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"\s+", "", str(v)).upper()


def looks_like_code(v):
    """A plant code is one 'word' (no spaces) with a letter in it, or a long number."""
    if v is None:
        return False
    s = str(v).strip()
    if not s or " " in s or len(s) > 30:
        return False
    n = norm(s)
    if n in NOT_CODES:
        return False
    return bool(re.search(r"[A-Z]", n)) or (n.isdigit() and len(n) >= 4)


def open_xlsx(path):
    """Old .xls files get converted to .xlsx first."""
    if path.lower().endswith(".xls"):
        out = tempfile.mkdtemp()
        subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", out, path],
                       capture_output=True, check=True)
        path = os.path.join(out, os.path.splitext(os.path.basename(path))[0] + ".xlsx")
    return path


def read_inhouse(path):
    """Every code-like value on every sheet -> which job sheet(s) it's on. Same as Ctrl+F across the workbook."""
    wb = load_workbook(open_xlsx(path), data_only=True)
    codes = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for v in row:
                if looks_like_code(v):
                    codes.setdefault(norm(v), set()).add(ws.title)
                elif isinstance(v, str) and re.search(r"[+,;&]", v):
                    # several asset numbers in one cell, e.g. "A1459741 + A1156216"
                    for part in re.split(r"[+,;&]", v):
                        if looks_like_code(part) and re.search(r"\d", part):
                            codes.setdefault(norm(part), set()).add(ws.title)
    return codes


def cell_codes(v):
    """Codes in one cell: the cell itself, or each part of "A1459741 + A1156216"."""
    if looks_like_code(v):
        return [norm(v)]
    if isinstance(v, str) and re.search(r"[+,;&]", v):
        return [norm(p) for p in re.split(r"[+,;&]", v) if looks_like_code(p) and re.search(r"\d", p)]
    return []


def mark_inhouse(path, charged):
    """Copy of the ACS report with every plant line that is charged on a supplier report in green + bold.
    A line listing several assets is only marked when all of them are charged."""
    src = open_xlsx(path)
    wb = load_workbook(src)
    vals = load_workbook(src, data_only=True)
    marked = 0
    for ws in wb.worksheets:
        vws = vals[ws.title]
        asset_col, width = None, 0
        for r, row in enumerate(vws.iter_rows(values_only=True), start=1):
            hdr = [v.strip().lower() if isinstance(v, str) else "" for v in row]
            a = next((i for i, h in enumerate(hdr) if re.search(r"asset|plant\s*code|^code$", h)), None)
            if a is not None and any(re.search(r"company|supplier", h) for h in hdr):
                asset_col, width = a, max(i for i, h in enumerate(hdr) if h) + 1   # header row (can repeat)
                continue
            if asset_col is None:
                continue
            found = cell_codes(row[asset_col] if asset_col < len(row) else None)
            if found and all(c in charged for c in found):
                for c in range(1, width + 1):
                    cell = ws.cell(r, c)
                    if isinstance(cell, MergedCell):
                        continue
                    cell.fill = GREEN
                    cell.font = cell.font.copy(bold=True)
                marked += 1
    base = os.path.splitext(os.path.basename(path))[0]
    out = os.path.join(os.path.dirname(os.path.abspath(path)), f"{base}_checked.xlsx")
    wb.save(out)
    return out, marked


def find_code_column(ws, codes):
    """The plant-code column is whichever column has the most hits against the in-house report."""
    best_col, best_hits = None, 0
    for col in ws.iter_cols():
        hits = sum(1 for c in col if looks_like_code(c.value) and norm(c.value) in codes)
        if hits > best_hits:
            best_col, best_hits = col[0].column, hits
    if best_col:
        return best_col
    # nothing matched at all - fall back to a column headed like a code column
    for row in ws.iter_rows(max_row=30):
        for c in row:
            if isinstance(c.value, str) and re.search(r"code|plant|item|product", c.value, re.I):
                return c.column
    return None


def style_row(ws, r, fill, bold):
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(r, c)
        if isinstance(cell, MergedCell):
            continue
        cell.fill = fill
        if bold:
            cell.font = cell.font.copy(bold=True)


def check_supplier(path, codes, code_col=None):
    src = open_xlsx(path)
    wb = load_workbook(src)                    # keeps formulas/formatting for the output
    vals = load_workbook(src, data_only=True)  # calculated values for reading
    matched, flagged = [], []

    for ws in wb.worksheets:
        vws = vals[ws.title]
        col = column_index_from_string(code_col) if code_col else find_code_column(vws, codes)
        if not col:
            continue
        for r in range(1, vws.max_row + 1):
            v = vws.cell(r, col).value
            if not looks_like_code(v):
                continue
            line = [vws.cell(r, c).value for c in range(1, vws.max_column + 1)]
            if norm(v) in codes:
                style_row(ws, r, GREEN, bold=True)
                matched.append((ws.title, r, v, ", ".join(sorted(codes[norm(v)]))))
            else:
                # not on our report - left untouched in the file, reported on screen instead
                flagged.append((ws.title, r, v, line))

    base = os.path.splitext(os.path.basename(path))[0]
    out = os.path.join(os.path.dirname(os.path.abspath(path)), f"{base}_checked.xlsx")
    wb.save(out)
    return out, matched, flagged


if __name__ == "__main__":
    args = sys.argv[1:]
    code_col = None
    if "--code-col" in args:
        i = args.index("--code-col")
        code_col = args[i + 1]
        args = args[:i] + args[i + 2:]
    inhouse, suppliers = args[0], args[1:]

    codes = read_inhouse(inhouse)
    all_flagged = 0
    charged = set()
    for s in suppliers:
        out, m, f = check_supplier(s, codes, code_col)
        all_flagged += len(f)
        charged.update(norm(v) for _, _, v, _ in m)
        name = os.path.basename(s)
        if not f:
            print(f"OK  {name}: all {len(m)} charged items are on the ACS report")
            continue
        print(f"!!  {name}: {len(f)} charged item(s) NOT on the ACS report ({len(m)} matched)")
        for sheet, r, v, line in f:
            detail = "  |  ".join(str(x) for x in line if x not in (None, "") and x != v)
            print(f"      row {r}: {v}  -  {detail}")
    print(f"\nTOTAL NOT ON ACS REPORT: {all_flagged}")
    acs_out, acs_marked = mark_inhouse(inhouse, charged)
    print(f"ACS report: {acs_marked} line(s) charged on the supplier reports marked green -> {os.path.basename(acs_out)}")

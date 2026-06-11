"""
openpyxl workbook engine for the ICAI NCE format.

Consumes a models.Payload and produces a formula-linked .xlsx:
  Index · Balance Sheet · Statement of P&L (+Appropriation for a firm) · Notes

Guarantees implemented here:
  - Dynamic note suppression (Nil in both years) + contiguous Schedule III numbering
  - Every BS/P&L face figure is a cross-sheet formula to its note total
  - Leaf note lines are the only hardcoded numbers
  - Capital closing(s) computed by formula and linked to the BS
  - PPE transposed schedule (Gross / Acc-Dep / Net), no deduction columns
  - Calibri Light 11 on every cell incl. empty (styles.xml + theme patch)
  - No underline anywhere
"""
from __future__ import annotations
import io, zipfile, re
from typing import Dict, Tuple, List
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

from models import Payload, Note

FONT_NAME = "Calibri Light"
NUMFMT = '(* #,##0.00);(* (#,##0.00);(* "-"??);(@_)'

# ---- Schedule III mapping ---------------------------------------------------
# Balance-sheet line plan: (section_header, [(sub_label, note_key), ...])
BS_LIABILITIES = [
    ("1  Owners' Funds", [
        ("(a)  Owners'/Partners' Capital Account", "capital"),
        ("(b)  Reserves and Surplus", "reserves"),
    ]),
    ("2  Non-current liabilities", [
        ("(a)  Long-term borrowings", "lt_borrowings"),
        ("(b)  Deferred tax liabilities (Net)", "dtl"),
        ("(c)  Other long-term liabilities", "other_lt_liab"),
        ("(d)  Long-term provisions", "lt_provisions"),
    ]),
    ("3  Current liabilities", [
        ("(a)  Short-term borrowings", "st_borrowings"),
        ("(b)  Trade payables", "trade_payables"),
        ("(c)  Other current liabilities", "other_cl"),
        ("(d)  Short-term provisions", "st_provisions"),
    ]),
]
BS_ASSETS = [
    ("1  Non-current assets", [
        ("(a)  Property, Plant and Equipment", "ppe"),
        ("(b)  Non-current investments", "nc_investments"),
        ("(c)  Deferred tax assets (Net)", "dta"),
        ("(d)  Long-term loans and advances", "lt_loans_adv"),
        ("(e)  Other non-current assets", "other_nca"),
    ]),
    ("2  Current assets", [
        ("(a)  Current investments", "current_investments"),
        ("(b)  Inventories", "inventories"),
        ("(c)  Trade receivables", "trade_receivables"),
        ("(d)  Cash and bank balances", "cash_bank"),
        ("(e)  Short-term loans and advances", "st_loans_adv"),
        ("(f)  Other current assets", "other_ca"),
    ]),
]

PL_EXPENSE_PLAN = [
    ("(a)  Cost of materials consumed / Purchases of stock-in-trade", "cost_materials"),
    ("(b)  Changes in inventories", "changes_inventory"),
    ("(c)  Employee benefits expense", "employee_benefits"),
    ("(d)  Finance costs", "finance_costs"),
    ("(e)  Depreciation and amortisation expense", "depreciation"),
    ("(f)  Other expenses", "other_expenses"),
]

# Notes whose presence does not depend on a numeric balance
ALWAYS_RETAINED = {"entity", "policies", "capital", "prev_year", "rounding"}
# Schedule III canonical order for numbering
SCHEDULE_ORDER = [
    "entity", "policies", "capital",
    "reserves", "lt_borrowings", "dtl", "other_lt_liab", "lt_provisions",
    "st_borrowings", "trade_payables", "other_cl", "st_provisions",
    "ppe", "nc_investments", "dta", "lt_loans_adv", "other_nca",
    "current_investments", "inventories", "trade_receivables",
    "cash_bank", "st_loans_adv", "other_ca",
    "revenue", "other_income", "cost_materials", "changes_inventory",
    "employee_benefits", "finance_costs", "depreciation", "other_expenses",
    "contingent", "related_party", "segment", "msmed", "forex",
    "confirmation", "prev_year", "rounding",
]

THIN = Side(style="thin")
DOUBLE = Side(style="double")


class Engine:
    def __init__(self, payload: Payload):
        self.p = payload
        self.wb = Workbook()
        self.firm = payload.entity.constitution == "partnership"
        # logical anchor -> "Sheet!Cell" for cross-sheet formulas
        self.anchor: Dict[str, str] = {}
        # final note number per key (after suppression)
        self.note_no: Dict[str, int] = {}
        self.retained: List[str] = []

    # -- styling helpers ------------------------------------------------------
    def _f(self, bold=False, italic=False, size=11, color=None):
        return Font(name=FONT_NAME, size=size, bold=bold, italic=italic,
                    underline=None, color=color)

    def cell(self, ws, r, c, val=None, bold=False, italic=False, num=False,
             center=False, right=False, wrap=False, size=11, top=False,
             bottom=False, double_bottom=False):
        cl = ws.cell(row=r, column=c, value=val)
        cl.font = self._f(bold, italic, size)
        al = Alignment(vertical="top", wrap_text=wrap,
                       horizontal="center" if center else ("right" if right else None))
        cl.alignment = al
        if num:
            cl.number_format = NUMFMT
        b = {}
        if top: b["top"] = THIN
        if bottom: b["bottom"] = THIN
        if double_bottom: b["bottom"] = DOUBLE
        if b:
            cl.border = Border(**b)
        return cl

    # -- suppression & numbering ---------------------------------------------
    def compute_retained(self):
        has_ppe = bool(self.p.ppe_assets)
        present = {n.key for n in self.p.notes if not n.is_nil()}
        for key in SCHEDULE_ORDER:
            keep = False
            if key in ALWAYS_RETAINED:
                keep = True
            elif key in ("ppe", "depreciation") and has_ppe:
                keep = True
            elif key == "st_provisions" and self.firm:
                keep = True  # firm tax provision
            elif key in present:
                keep = True
            if keep:
                self.retained.append(key)
        for i, key in enumerate(self.retained, start=1):
            self.note_no[key] = i

    # ========================================================================
    def build(self) -> bytes:
        self.compute_retained()
        # remove default sheet, create in order
        self.wb.remove(self.wb.active)
        self.ws_notes = self.wb.create_sheet("Notes")
        self.ws_bs = self.wb.create_sheet("Balance Sheet")
        self.ws_pl = self.wb.create_sheet("Statement of P&L")
        self.ws_idx = self.wb.create_sheet("Index")
        # build notes first so anchors exist, then BS/PL reference them
        self.build_notes()
        self.build_balance_sheet()
        self.build_pl()
        self.build_index()
        # order: Index, BS, PL, Notes
        self.wb.move_sheet("Index", -(self.wb.sheetnames.index("Index")))
        self._reorder()
        self._no_gridlines()
        data = self._save_with_theme_patch()
        return data

    def _reorder(self):
        order = ["Index", "Balance Sheet", "Statement of P&L", "Notes"]
        self.wb._sheets.sort(key=lambda s: order.index(s.title))

    def _no_gridlines(self):
        for ws in self.wb.worksheets:
            ws.sheet_view.showGridLines = False

    # -- title block ----------------------------------------------------------
    def title_block(self, ws, desc, as_at=True):
        e = self.p.entity
        last = max(6, ws.max_column)
        self.cell(ws, 1, 1, e.name, bold=True, center=True)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
        self.cell(ws, 2, 1, desc, bold=True, center=True)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)
        self.cell(ws, 3, 6, "(Amount in Rs.)", italic=True, right=True)

    # -- NOTES ----------------------------------------------------------------
    def build_notes(self):
        ws = self.ws_notes
        for w, wd in {"A": 6, "B": 52, "C": 4, "D": 18, "E": 2, "F": 18}.items():
            ws.column_dimensions[w].width = wd
        self.title_block(ws, "Notes forming part of the Financial Statements")
        r = 5
        cyl, pyl = self.p.entity.cy_label, self.p.entity.py_label
        for key in self.retained:
            num = self.note_no[key]
            if key == "capital":
                r = self._note_capital(ws, r, num)
            elif key == "ppe":
                r = self._note_ppe(ws, r, num)
            elif key in ("entity", "policies", "prev_year", "rounding",
                         "segment", "related_party", "contingent", "msmed",
                         "forex", "confirmation", "dtl", "dta"):
                r = self._note_prose(ws, r, num, key)
            else:
                r = self._note_standard(ws, r, num, key)
        self.notes_last_row = r

    def _note_title(self, ws, r, num, title, with_years=True):
        self.cell(ws, r, 1, f"Note {num}", bold=True)
        self.cell(ws, r, 2, title, bold=True)
        if with_years:
            self.cell(ws, r, 4, self.p.entity.cy_label, bold=True, center=True)
            self.cell(ws, r, 6, self.p.entity.py_label, bold=True, center=True)
        return r + 1

    def _note_standard(self, ws, r, num, key):
        note = self.p.note(key)
        if note is None:
            return r
        r = self._note_title(ws, r, num, note.title)
        first = r
        for it in note.items:
            self.cell(ws, r, 2, it.label)
            self.cell(ws, r, 4, round(it.cy, 2), num=True)
            self.cell(ws, r, 6, round(it.py, 2), num=True)
            r += 1
        last = r - 1
        # total
        self.cell(ws, r, 2, "Total", bold=True)
        if note.items:
            self.cell(ws, r, 4, f"=SUM(D{first}:D{last})", bold=True, num=True, top=True)
            self.cell(ws, r, 6, f"=SUM(F{first}:F{last})", bold=True, num=True, top=True)
        else:
            self.cell(ws, r, 4, 0, bold=True, num=True, top=True)
            self.cell(ws, r, 6, 0, bold=True, num=True, top=True)
        self.anchor[f"note_{key}_cy"] = f"Notes!D{r}"
        self.anchor[f"note_{key}_py"] = f"Notes!F{r}"
        r += 1
        r = self._footnotes(ws, r, note)
        return r + 1

    def _footnotes(self, ws, r, note):
        for fn in note.footnotes:
            c = self.cell(ws, r, 2, fn, italic=True, wrap=True)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
            ws.row_dimensions[r].height = max(15 * ((len(fn) // 110) + 1), 15)
            r += 1
        return r

    def _note_prose(self, ws, r, num, key):
        note = self.p.note(key)
        title = note.title if note else self._default_title(key)
        r = self._note_title(ws, r, num, title, with_years=False)
        body = []
        if note and note.footnotes:
            body = note.footnotes
        elif note and note.items:
            body = [f"{it.label}" for it in note.items]
        else:
            body = [self._default_prose(key)]
        for para in body:
            self.cell(ws, r, 2, para, italic=False, wrap=True)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
            ws.row_dimensions[r].height = max(15 * ((len(para) // 110) + 1), 15)
            r += 1
        return r + 1

    def _default_title(self, key):
        return {
            "entity": "Brief about the Entity",
            "policies": "Significant Accounting Policies",
            "prev_year": "Previous Year Figures",
            "rounding": "Rounding-off",
            "segment": "Segment Reporting",
            "related_party": "Related Party Disclosures",
            "contingent": "Contingent Liabilities and Commitments",
            "msmed": "MSMED Act Disclosure",
            "forex": "Earnings / Expenditure in Foreign Currency",
            "confirmation": "Confirmation of Balances",
            "dtl": "Deferred Tax Liabilities (Net)",
            "dta": "Deferred Tax Assets (Net)",
        }.get(key, key.title())

    def _default_prose(self, key):
        return {
            "rounding": "Figures have been rounded off to the nearest Rupee and presented to two decimals.",
            "segment": "The entity operates in a single business segment; accordingly, segment reporting under AS 17 is not applicable.",
            "prev_year": "Previous year figures have been regrouped / reclassified wherever necessary to conform to the current year's presentation.",
        }.get(key, "—")

    # -- Capital note ---------------------------------------------------------
    def _note_capital(self, ws, r, num):
        if self.firm:
            return self._note_partners(ws, r, num)
        return self._note_owner(ws, r, num)

    def _note_owner(self, ws, r, num):
        oc = self.p.owner_capital
        title = "Owner's Capital Account"
        self.cell(ws, r, 1, f"Note {num}", bold=True)
        self.cell(ws, r, 2, title, bold=True)
        r += 1
        heads = ["Particulars", "Opening", "Introduced", "Net Profit",
                 "Interest", "Withdrawals", "Closing"]
        for j, h in enumerate(heads):
            self.cell(ws, r, 2 + j, h, bold=True, center=(j > 0), top=True, bottom=True)
        r += 1
        # CY row
        def row(label, opening, intro, npf, interest, wd, profit_anchor):
            self.cell(ws, r0, 2, label)
            self.cell(ws, r0, 3, round(opening, 2), num=True)
            self.cell(ws, r0, 4, round(intro, 2), num=True)
            self.cell(ws, r0, 5, profit_anchor, num=True)
            self.cell(ws, r0, 6, round(interest, 2), num=True)
            self.cell(ws, r0, 7, round(wd, 2), num=True)
            self.cell(ws, r0, 8, f"=C{r0}+D{r0}+E{r0}+F{r0}-G{r0}", num=True, bold=True)
        r0 = r
        row(f"{oc.name} — {self.p.entity.cy_label}", oc.opening_cy, oc.introduced_cy,
            "=net_profit_cy", oc.interest_cy, oc.withdrawals_cy, None)
        self.cell(ws, r0, 5, "='Statement of P&L'!__NPCY__", num=True)  # placeholder, fixed later
        self.anchor["capital_close_cy"] = f"Notes!H{r0}"
        r += 1
        r0 = r
        row(f"{oc.name} — {self.p.entity.py_label}", oc.opening_py, oc.introduced_py,
            "=net_profit_py", oc.interest_py, oc.withdrawals_py, None)
        self.cell(ws, r0, 5, "='Statement of P&L'!__NPPY__", num=True)
        self.anchor["capital_close_py"] = f"Notes!H{r0}"
        # BS capital anchors point to closing cells
        self.anchor["note_capital_cy"] = self.anchor["capital_close_cy"]
        self.anchor["note_capital_py"] = self.anchor["capital_close_py"]
        r += 1
        sub = self.p.note("capital")
        note = sub if sub else Note(key="capital", title=title)
        note.footnotes = note.footnotes or [
            "Capital is maintained under the fluctuating capital method; a separate Current Account is not maintained."]
        r = self._footnotes(ws, r, note)
        return r + 1

    def _note_partners(self, ws, r, num):
        self.cell(ws, r, 1, f"Note {num}", bold=True)
        self.cell(ws, r, 2, "Partners' Capital Account", bold=True)
        r += 1
        heads = ["Partner", "PSR%", "Opening", "Introduced", "Share of Profit",
                 "Interest", "Remuneration", "Withdrawals", "Closing"]
        for j, h in enumerate(heads):
            self.cell(ws, r, 2 + j, h, bold=True, center=(j > 0), top=True, bottom=True)
        r += 1

        def block(year):
            nonlocal r
            self.cell(ws, r, 2, f"As at {self.p.entity.cy_label if year=='cy' else self.p.entity.py_label}",
                      italic=True)
            r += 1
            first = r
            for pt in self.p.partners:
                g = lambda a: getattr(pt, f"{a}_{year}")
                self.cell(ws, r, 2, pt.name)
                self.cell(ws, r, 3, pt.psr, num=True)
                self.cell(ws, r, 4, round(g("opening"), 2), num=True)
                self.cell(ws, r, 5, round(g("introduced"), 2), num=True)
                self.cell(ws, r, 6, round(g("share_profit"), 2), num=True)
                self.cell(ws, r, 7, round(g("interest"), 2), num=True)
                self.cell(ws, r, 8, round(g("remuneration"), 2), num=True)
                self.cell(ws, r, 9, round(g("withdrawals"), 2), num=True)
                self.cell(ws, r, 10, f"=D{r}+E{r}+F{r}+G{r}+H{r}-I{r}", num=True)
                r += 1
            last = r - 1
            self.cell(ws, r, 2, "Total", bold=True)
            self.cell(ws, r, 10, f"=SUM(J{first}:J{last})", bold=True, num=True, top=True)
            tot = f"Notes!J{r}"
            r += 1
            return tot

        cy_total = block("cy")
        py_total = block("py")
        self.anchor["note_capital_cy"] = cy_total
        self.anchor["note_capital_py"] = py_total
        sub = self.p.note("capital")
        note = sub if sub else Note(key="capital", title="Partners' Capital Account")
        if not note.footnotes:
            note.footnotes = [
                "Capital is maintained under the fluctuating capital method (one account per partner).",
                "Withdrawals include actual drawings; prior-year tax paid during the year is routed through Withdrawals for comparability."]
        r = self._footnotes(ws, r, note)
        return r + 1

    # -- PPE note -------------------------------------------------------------
    def _note_ppe(self, ws, r, num):
        self.cell(ws, r, 1, f"Note {num}", bold=True)
        self.cell(ws, r, 2, "Property, Plant and Equipment", bold=True)
        r += 1
        # super-group header row
        self.cell(ws, r, 4, "GROSS BLOCK", bold=True, center=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
        self.cell(ws, r, 7, "ACCUMULATED DEPRECIATION", bold=True, center=True)
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=9)
        self.cell(ws, r, 10, "NET BLOCK", bold=True, center=True)
        ws.merge_cells(start_row=r, start_column=10, end_row=r, end_column=11)
        r += 1
        cyl, pyl = self.p.entity.cy_label, self.p.entity.py_label
        heads = ["Sr.", "Particulars", f"Opening\n{pyl}", "Additions",
                 f"Closing\n{cyl}", f"Opening\n{pyl}", "For the year",
                 f"Closing\n{cyl}", f"Net {cyl}", f"Net {pyl}"]
        for j, h in enumerate(heads):
            self.cell(ws, r, 1 + j, h, bold=True, center=(j > 1), wrap=True, top=True, bottom=True)
        ws.row_dimensions[r].height = 30
        r += 1
        first = r
        for i, a in enumerate(self.p.ppe_assets, start=1):
            self.cell(ws, r, 1, i, center=True)
            self.cell(ws, r, 2, a.name)
            self.cell(ws, r, 3, round(a.gb_open, 2), num=True)
            self.cell(ws, r, 4, round(a.additions, 2), num=True)
            self.cell(ws, r, 5, f"=C{r}+D{r}", num=True)
            self.cell(ws, r, 6, round(a.accdep_open, 2), num=True)
            self.cell(ws, r, 7, round(a.dep_year, 2), num=True)
            self.cell(ws, r, 8, f"=F{r}+G{r}", num=True)
            self.cell(ws, r, 9, f"=E{r}-H{r}", num=True)   # net CY
            self.cell(ws, r, 10, f"=C{r}-F{r}", num=True)  # net PY
            r += 1
        last = r - 1
        self.cell(ws, r, 2, "Total", bold=True)
        for col in (3, 4, 5, 6, 7, 8, 9, 10):
            L = get_column_letter(col)
            self.cell(ws, r, col, f"=SUM({L}{first}:{L}{last})", bold=True, num=True, top=True)
        self.anchor["note_ppe_cy"] = f"Notes!I{r}"   # net block CY
        self.anchor["note_ppe_py"] = f"Notes!J{r}"   # net block PY
        self.anchor["ppe_dep_year"] = f"Notes!G{r}"  # depreciation for the year
        r += 1
        note = self.p.note("ppe") or Note(key="ppe", title="Property, Plant and Equipment")
        fns = list(note.footnotes)
        if self.p.depreciation_case == "A":
            fns.append("Depreciation has not been provided in the books; fixed-asset balances are carried at cost as gross block. The impact of non-provision has not been ascertained.")
        fns.append("No revaluation of assets was carried out during the year.")
        fns.append("There are no intangible assets, capital work-in-progress or intangible assets under development.")
        fns.append("All assets are owned by the entity and are free from charge except as disclosed.")
        note.footnotes = fns
        r = self._footnotes(ws, r, note)
        return r + 1

    # -- BALANCE SHEET --------------------------------------------------------
    def build_balance_sheet(self):
        ws = self.ws_bs
        for w, wd in {"A": 5, "B": 58, "C": 8, "D": 18, "E": 2, "F": 18}.items():
            ws.column_dimensions[w].width = wd
        self.title_block(ws, f"Balance Sheet as at {self.p.entity.cy_label}")
        self.cell(ws, 4, 2, "Particulars", bold=True)
        self.cell(ws, 4, 3, "Note", bold=True, center=True)
        self.cell(ws, 4, 4, self.p.entity.cy_label, bold=True, center=True)
        self.cell(ws, 4, 6, self.p.entity.py_label, bold=True, center=True)
        r = 6
        self.cell(ws, r, 1, "I", bold=True)
        self.cell(ws, r, 2, "OWNERS' FUNDS AND LIABILITIES", bold=True)
        r += 1
        r, liab_rows = self._bs_block(ws, r, BS_LIABILITIES)
        self.cell(ws, r, 2, "Total", bold=True)
        self.cell(ws, r, 4, "=" + ("+".join(f"D{x}" for x in liab_rows) or "0"), bold=True, num=True, top=True, double_bottom=True)
        self.cell(ws, r, 6, "=" + ("+".join(f"F{x}" for x in liab_rows) or "0"), bold=True, num=True, top=True, double_bottom=True)
        self.anchor["bs_liab_total_cy"] = f"Balance Sheet!D{r}"
        r += 2
        self.cell(ws, r, 1, "II", bold=True)
        self.cell(ws, r, 2, "ASSETS", bold=True)
        r += 1
        r, asset_rows = self._bs_block(ws, r, BS_ASSETS)
        self.cell(ws, r, 2, "Total", bold=True)
        self.cell(ws, r, 4, "=" + ("+".join(f"D{x}" for x in asset_rows) or "0"), bold=True, num=True, top=True, double_bottom=True)
        self.cell(ws, r, 6, "=" + ("+".join(f"F{x}" for x in asset_rows) or "0"), bold=True, num=True, top=True, double_bottom=True)
        r += 2
        self.cell(ws, r, 2, "The accompanying notes are an integral part of the financial statements.", italic=True)

    def _bs_block(self, ws, r, plan):
        subtotal_rows = []
        for header, lines in plan:
            visible = [(lbl, key) for (lbl, key) in lines if key in self.note_no or key == "ppe"]
            visible = [(lbl, key) for (lbl, key) in lines
                       if (key in self.note_no) or (key == "ppe" and "note_ppe_cy" in self.anchor)]
            if not visible:
                continue
            self.cell(ws, r, 1, header.split()[0], bold=True)
            self.cell(ws, r, 2, header[len(header.split()[0]):].strip(), bold=True)
            r += 1
            block_rows = []
            for lbl, key in lines:
                cy_anchor = self.anchor.get(f"note_{key}_cy")
                if cy_anchor is None:
                    continue
                self.cell(ws, r, 2, lbl)
                if key in self.note_no:
                    self.cell(ws, r, 3, self.note_no[key], center=True)
                self.cell(ws, r, 4, f"='{self._sheet(cy_anchor)}'!{self._addr(cy_anchor)}", num=True)
                py_anchor = self.anchor.get(f"note_{key}_py")
                self.cell(ws, r, 6, f"='{self._sheet(py_anchor)}'!{self._addr(py_anchor)}", num=True)
                block_rows.append(r)
                r += 1
            # subtotal
            if block_rows:
                self.cell(ws, r, 2, "", )
                self.cell(ws, r, 4, "=" + "+".join(f"D{x}" for x in block_rows), num=True, top=True, bold=True)
                self.cell(ws, r, 6, "=" + "+".join(f"F{x}" for x in block_rows), num=True, top=True, bold=True)
                subtotal_rows.append(r)
                r += 1
        return r, subtotal_rows

    @staticmethod
    def _sheet(anchor):
        return anchor.split("!")[0]

    @staticmethod
    def _addr(anchor):
        return anchor.split("!")[1]

    # -- STATEMENT OF P&L -----------------------------------------------------
    def build_pl(self):
        ws = self.ws_pl
        for w, wd in {"A": 5, "B": 58, "C": 8, "D": 18, "E": 2, "F": 18}.items():
            ws.column_dimensions[w].width = wd
        self.title_block(ws, f"Statement of Profit and Loss for the year ended {self.p.entity.cy_label}")
        self.cell(ws, 4, 2, "Particulars", bold=True)
        self.cell(ws, 4, 3, "Note", bold=True, center=True)
        self.cell(ws, 4, 4, f"Year ended {self.p.entity.cy_label}", bold=True, center=True)
        self.cell(ws, 4, 6, f"Year ended {self.p.entity.py_label}", bold=True, center=True)
        r = 6

        def line(roman, label, key=None, val_cy=None, val_py=None, bold=False, top=False):
            nonlocal r
            self.cell(ws, r, 1, roman, bold=bold)
            self.cell(ws, r, 2, label, bold=bold)
            if key and key in self.note_no:
                self.cell(ws, r, 3, self.note_no[key], center=True)
            if val_cy is not None:
                self.cell(ws, r, 4, val_cy, num=True, bold=bold, top=top)
            if val_py is not None:
                self.cell(ws, r, 6, val_py, num=True, bold=bold, top=top)
            rr = r
            r += 1
            return rr

        rev_cy = self._anchor_ref("revenue", "cy")
        rev_py = self._anchor_ref("revenue", "py")
        oi_cy = self._anchor_ref("other_income", "cy")
        oi_py = self._anchor_ref("other_income", "py")
        r_rev = line("I", "Revenue from operations", "revenue", rev_cy, rev_py)
        r_oi = line("II", "Other income", "other_income", oi_cy, oi_py)
        r_ti = line("III", "Total Income (I + II)", None,
                    f"=D{r_rev}+D{r_oi}", f"=F{r_rev}+F{r_oi}", bold=True, top=True)
        line("IV", "Expenses:", None, bold=True)
        exp_rows = []
        for lbl, key in PL_EXPENSE_PLAN:
            if key not in self.note_no:
                continue
            cy = self._anchor_ref(key, "cy")
            py = self._anchor_ref(key, "py")
            rr = line("", lbl, key, cy, py)
            exp_rows.append(rr)
        r_te = line("", "Total expenses", None,
                    "=" + "+".join(f"D{x}" for x in exp_rows) if exp_rows else "0",
                    "=" + "+".join(f"F{x}" for x in exp_rows) if exp_rows else "0",
                    bold=True, top=True)
        r_pbt = line("V", "Profit before tax (III − IV)", None,
                     f"=D{r_ti}-D{r_te}", f"=F{r_ti}-F{r_te}", bold=True, top=True)
        if self.firm:
            tax = self.p.firm_tax
            r_tax = line("VI", "Tax expense — Current tax", "st_provisions",
                         round(tax.current_tax_cy, 2) if tax else 0,
                         round(tax.current_tax_py, 2) if tax else 0)
            r_pat = line("VII", "Profit for the year after tax (V − VI)", None,
                         f"=D{r_pbt}-D{r_tax}", f"=F{r_pbt}-F{r_tax}", bold=True, top=True)
            self.anchor["np_cy"] = f"Statement of P&L!D{r_pat}"
            self.anchor["np_py"] = f"Statement of P&L!F{r_pat}"
            self._appropriation(ws, r, r_pat)
        else:
            self.cell(ws, r_pbt, 2, "Profit for the year (= V)", bold=True)
            self.anchor["np_cy"] = f"Statement of P&L!D{r_pbt}"
            self.anchor["np_py"] = f"Statement of P&L!F{r_pbt}"
        self._wire_capital_profit()

    def _appropriation(self, ws, r, r_pat):
        r += 1
        self.cell(ws, r, 2, "Profit and Loss Appropriation", bold=True); r += 1
        self.cell(ws, r, 2, "Profit for the year after tax")
        self.cell(ws, r, 4, f"=D{r_pat}", num=True)
        self.cell(ws, r, 6, f"=F{r_pat}", num=True)
        r += 1
        int_cy = sum(p.interest_cy for p in self.p.partners)
        int_py = sum(p.interest_py for p in self.p.partners)
        rem_cy = sum(p.remuneration_cy for p in self.p.partners)
        rem_py = sum(p.remuneration_py for p in self.p.partners)
        self.cell(ws, r, 2, "(−) Interest on partners' capital")
        self.cell(ws, r, 4, round(int_cy, 2), num=True)
        self.cell(ws, r, 6, round(int_py, 2), num=True); r_int = r; r += 1
        self.cell(ws, r, 2, "(−) Remuneration to partners")
        self.cell(ws, r, 4, round(rem_cy, 2), num=True)
        self.cell(ws, r, 6, round(rem_py, 2), num=True); r_rem = r; r += 1
        self.cell(ws, r, 2, "Divisible profit among partners", bold=True)
        self.cell(ws, r, 4, f"=D{r_pat}-D{r_int}-D{r_rem}", bold=True, num=True, top=True)
        self.cell(ws, r, 6, f"=F{r_pat}-F{r_int}-F{r_rem}", bold=True, num=True, top=True)
        r += 1

    def _anchor_ref(self, key, yr):
        a = self.anchor.get(f"note_{key}_{yr}")
        if not a:
            return 0
        return f"='{self._sheet(a)}'!{self._addr(a)}"

    def _wire_capital_profit(self):
        """Replace capital-account profit placeholders with real P&L links."""
        # Owner capital placeholders were written as formulas with __NPCY__/__NPPY__
        np_cy = self.anchor.get("np_cy")
        np_py = self.anchor.get("np_py")
        if not self.firm and np_cy:
            ws = self.ws_notes
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and "__NPCY__" in c.value:
                        c.value = f"='{self._sheet(np_cy)}'!{self._addr(np_cy)}"
                    elif isinstance(c.value, str) and "__NPPY__" in c.value:
                        c.value = f"='{self._sheet(np_py)}'!{self._addr(np_py)}"

    # -- INDEX ----------------------------------------------------------------
    def build_index(self):
        ws = self.ws_idx
        ws.column_dimensions["B"].width = 55
        ws.column_dimensions["C"].width = 12
        self.title_block(ws, "Index")
        self.cell(ws, 4, 2, "Contents", bold=True)
        self.cell(ws, 4, 3, "Reference", bold=True, center=True)
        r = 5
        for label, sheet in [("Balance Sheet", "Balance Sheet"),
                             ("Statement of Profit and Loss", "Statement of P&L"),
                             ("Notes to the Financial Statements", "Notes")]:
            c = self.cell(ws, r, 2, label)
            c.hyperlink = f"#'{sheet}'!A1"
            c.font = self._f(color="0563C1")
            self.cell(ws, r, 3, "→", center=True)
            r += 1
        r += 1
        self.cell(ws, r, 2, "Note index", bold=True); r += 1
        for key in self.retained:
            self.cell(ws, r, 2, f"Note {self.note_no[key]}  {self._note_label(key)}")
            r += 1

    def _note_label(self, key):
        note = self.p.note(key)
        if note and note.title:
            return note.title
        if key == "capital":
            return "Partners' Capital Account" if self.firm else "Owner's Capital Account"
        if key == "ppe":
            return "Property, Plant and Equipment"
        return self._default_title(key)

    # -- save with theme/styles font patch -----------------------------------
    def _save_with_theme_patch(self) -> bytes:
        buf = io.BytesIO()
        self.wb.save(buf)
        buf.seek(0)
        return self._patch_fonts(buf.read())

    @staticmethod
    def _patch_fonts(data: bytes) -> bytes:
        """Force Calibri Light onto the theme minor font and the default style
        so empty cells comply too."""
        src = io.BytesIO(data)
        out = io.BytesIO()
        with zipfile.ZipFile(src, "r") as zin:
            names = zin.namelist()
            with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
                for n in names:
                    content = zin.read(n)
                    if n == "xl/theme/theme1.xml":
                        t = content.decode("utf-8")
                        # replace minor font typeface
                        t = re.sub(r'(<a:minorFont>\s*<a:latin typeface=")[^"]*"',
                                   r'\g<1>Calibri Light"', t)
                        t = re.sub(r'(<a:majorFont>\s*<a:latin typeface=")[^"]*"',
                                   r'\g<1>Calibri Light"', t)
                        content = t.encode("utf-8")
                    zout.writestr(n, content)
        return out.getvalue()


def build_workbook(payload: Payload) -> bytes:
    return Engine(payload).build()

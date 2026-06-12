"""
Reconciliation Engine — zero-tolerance accounting checks.

Recomputes every total from the payload in pure Python (mirroring builder.py's
arithmetic, independent of Excel formulas) and asserts the converted statements
reconcile to the SOURCE control totals the LLM read from the printed statements.

Returns a list of discrepancies. A workbook is only delivered when there are none.
Also provides deterministic auto-fixes for the two most common extraction slips
(inventory double-count, bank-interest mis-placed in the capital interest column).
"""
from __future__ import annotations
from typing import List, Dict
from models import Payload

EPS = 0.5  # half-rupee: absorbs float noise, flags any real (>=Re.1) difference

ASSET_KEYS = ["nc_investments", "dta", "lt_loans_adv", "other_nca",
              "current_investments", "inventories", "trade_receivables",
              "cash_bank", "st_loans_adv", "other_ca"]
LIAB_KEYS = ["reserves", "lt_borrowings", "dtl", "other_lt_liab", "lt_provisions",
             "st_borrowings", "trade_payables", "other_cl", "st_provisions"]
EXPENSE_KEYS = ["cost_materials", "changes_inventory", "employee_benefits",
                "finance_costs", "depreciation", "other_expenses"]


def _note_total(p: Payload, key: str, yr: str) -> float:
    n = p.note(key)
    if not n:
        return 0.0
    return round(sum((getattr(i, yr) or 0.0) for i in n.items), 2)


def ppe_net(p: Payload, yr: str) -> float:
    tot = 0.0
    for a in p.ppe_assets:
        if yr == "cy":
            tot += (a.gb_open + a.additions) - (a.accdep_open + a.dep_year)
        else:
            tot += a.gb_open - a.accdep_open
    return round(tot, 2)


def dep_for_year_total(p: Payload, yr: str) -> float:
    # CY depreciation = sum of dep_year; PY not modelled per-asset (returns note value)
    if yr == "cy":
        return round(sum(a.dep_year for a in p.ppe_assets), 2)
    return _note_total(p, "depreciation", "py")


def pl_profit(p: Payload, yr: str) -> float:
    income = _note_total(p, "revenue", yr) + _note_total(p, "other_income", yr)
    expense = sum(_note_total(p, k, yr) for k in EXPENSE_KEYS)
    return round(income - expense, 2)


def capital_closing(p: Payload, yr: str) -> float:
    if p.entity.constitution == "partnership":
        tot = 0.0
        for pt in p.partners:
            g = lambda a: getattr(pt, f"{a}_{yr}") or 0.0
            tot += (g("opening") + g("introduced") + g("share_profit")
                    + g("interest") + g("remuneration") - g("withdrawals"))
        return round(tot, 2)
    oc = p.owner_capital
    if not oc:
        return 0.0
    g = lambda a: getattr(oc, f"{a}_{yr}") or 0.0
    # net profit in the capital account links to the P&L profit (engine behaviour)
    return round(g("opening") + g("introduced") + pl_profit(p, yr)
                 + g("interest") - g("withdrawals"), 2)


def assets_total(p: Payload, yr: str) -> float:
    return round(ppe_net(p, yr) + sum(_note_total(p, k, yr) for k in ASSET_KEYS), 2)


def liabilities_total(p: Payload, yr: str) -> float:
    return round(capital_closing(p, yr) + sum(_note_total(p, k, yr) for k in LIAB_KEYS), 2)


def compute(p: Payload) -> Dict:
    out = {}
    for yr in ("cy", "py"):
        out[yr] = {
            "assets": assets_total(p, yr),
            "liabilities": liabilities_total(p, yr),
            "capital_closing": capital_closing(p, yr),
            "pl_profit": pl_profit(p, yr),
            "ppe_net": ppe_net(p, yr),
        }
    return out


def reconcile(p: Payload) -> List[Dict]:
    """Return a list of discrepancies. Empty list = fully reconciled."""
    d: List[Dict] = []
    c = p.controls
    firm = p.entity.constitution == "partnership"
    for yr, ylabel in (("cy", p.entity.cy_label), ("py", p.entity.py_label)):
        at = assets_total(p, yr)
        lt = liabilities_total(p, yr)
        cap = capital_closing(p, yr)
        src_bs = getattr(c, f"bs_total_{yr}") or 0.0
        src_cap = getattr(c, f"capital_close_{yr}") or 0.0

        # 1. internal tally: assets == liabilities
        if abs(at - lt) > EPS:
            d.append({"check": "Balance Sheet does not tally (Assets ≠ Liabilities)",
                      "year": ylabel, "expected": at, "got": lt, "diff": round(at - lt, 2)})
        # 2. assets total == source BS total
        if src_bs and abs(at - src_bs) > EPS:
            d.append({"check": "Total Assets do not match source Balance Sheet total",
                      "year": ylabel, "expected": src_bs, "got": at, "diff": round(at - src_bs, 2)})
        # 3. capital closing == source capital
        if src_cap and abs(cap - src_cap) > EPS:
            d.append({"check": "Capital closing balance does not match source Balance Sheet capital",
                      "year": ylabel, "expected": src_cap, "got": cap, "diff": round(cap - src_cap, 2)})
        # 4. Case B depreciation reconciliation (CY only, per-asset modelled)
        if p.depreciation_case == "B" and p.ppe_assets:
            dep_sched = dep_for_year_total(p, yr) if yr == "cy" else _note_total(p, "depreciation", "py")
            dep_note = _note_total(p, "depreciation", yr)
            if yr == "cy" and abs(dep_sched - dep_note) > EPS:
                d.append({"check": "Depreciation for the year (PPE schedule) ≠ Depreciation note",
                          "year": ylabel, "expected": dep_sched, "got": dep_note,
                          "diff": round(dep_sched - dep_note, 2)})
        # 5. firm appropriation tie: interest + remuneration + share = PAT
        if firm:
            int_t = sum(getattr(pt, f"interest_{yr}") or 0 for pt in p.partners)
            rem_t = sum(getattr(pt, f"remuneration_{yr}") or 0 for pt in p.partners)
            shr_t = sum(getattr(pt, f"share_profit_{yr}") or 0 for pt in p.partners)
            tax = 0.0
            if p.firm_tax:
                tax = getattr(p.firm_tax, f"current_tax_{yr}") or 0.0
            pat = round(pl_profit(p, yr) - tax, 2)
            if abs((int_t + rem_t + shr_t) - pat) > EPS:
                d.append({"check": "Appropriation does not tie (Interest + Remuneration + Profit share ≠ Profit after tax)",
                          "year": ylabel, "expected": pat, "got": round(int_t + rem_t + shr_t, 2),
                          "diff": round((int_t + rem_t + shr_t) - pat, 2)})
    return d


# --------------------------------------------------------------------------
# Deterministic auto-fixes for known extraction slips
# --------------------------------------------------------------------------
def _set_note(p: Payload, key: str, label: str, cy: float, py: float):
    from models import Note, LineItem
    n = p.note(key)
    if n is None:
        n = Note(key=key, title=label)
        p.notes.append(n)
    n.items = [LineItem(label=label, cy=round(cy, 2), py=round(py, 2))]


def auto_fix(p: Payload) -> List[str]:
    """Apply safe, rule-based corrections. Returns a list of fixes applied."""
    fixes = []
    c = p.controls
    # (a) Inventory netting: if opening/closing/purchases known, force the
    #     combined method: Cost of materials = Opening + Purchases - Closing,
    #     and Changes in inventories = 0. Eliminates the double-count.
    for yr in ("cy", "py"):
        op = getattr(c, f"opening_stock_{yr}") or 0.0
        cl = getattr(c, f"closing_stock_{yr}") or 0.0
        pu = getattr(c, f"purchases_{yr}") or 0.0
        if pu > 0 or op > 0 or cl > 0:
            target = round(op + pu - cl, 2)
            cur = _note_total(p, "cost_materials", yr) + _note_total(p, "changes_inventory", yr)
            if abs(cur - target) > EPS:
                # rebuild cost_materials note for BOTH years coherently
                cm_cy = (getattr(c, "opening_stock_cy") or 0) + (getattr(c, "purchases_cy") or 0) - (getattr(c, "closing_stock_cy") or 0)
                cm_py = (getattr(c, "opening_stock_py") or 0) + (getattr(c, "purchases_py") or 0) - (getattr(c, "closing_stock_py") or 0)
                _set_note(p, "cost_materials", "Cost of materials consumed", cm_cy, cm_py)
                ci = p.note("changes_inventory")
                if ci:
                    ci.items = []
                fixes.append("Inventory netted into Cost of materials consumed (Opening + Purchases - Closing); Changes in inventories set to nil")
                break
    # (b) Bank interest mis-placed in proprietor capital interest column:
    #     if the capital 'interest' equals the Other Income total, it is bank
    #     interest (already in profit) wrongly duplicated -> zero it.
    if p.entity.constitution == "proprietorship" and p.owner_capital:
        oc = p.owner_capital
        for yr in ("cy", "py"):
            oi = _note_total(p, "other_income", yr)
            iv = getattr(oc, f"interest_{yr}") or 0.0
            if iv > 0 and abs(iv - oi) <= EPS:
                setattr(oc, f"interest_{yr}", 0.0)
                fixes.append(f"Removed bank interest wrongly placed in the capital 'interest on own capital' column ({yr.upper()})")
    return fixes


def report_text(discrepancies: List[Dict]) -> str:
    if not discrepancies:
        return "RECONCILED: the converted statements tie exactly to the source. No differences."
    lines = ["RECONCILIATION FAILED — the following figures do not tie to the source "
             "(no workbook has been produced):", ""]
    for x in discrepancies:
        lines.append(f"• {x['check']} [{x['year']}]: expected {x['expected']:,.2f}, "
                     f"got {x['got']:,.2f}, difference {x['diff']:,.2f}")
    return "\n".join(lines)

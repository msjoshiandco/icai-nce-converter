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
import re
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
    """Closing capital = signed sum of every capital-account line, exactly as
    presented (no merging). For a firm, the sum across all partners."""
    if p.entity.constitution == "partnership":
        tot = 0.0
        for pt in p.partners:
            tot += sum(l.signed(yr) for l in pt.resolved_lines())
        return round(tot, 2)
    oc = p.owner_capital
    if not oc:
        return 0.0
    return round(sum(l.signed(yr) for l in oc.resolved_lines()), 2)


def capital_profit_line(p: Payload, yr: str) -> float:
    """Total of the 'profit' kind line(s) in the proprietor's capital account."""
    oc = p.owner_capital
    if not oc:
        return 0.0
    return round(sum((getattr(l, yr) or 0.0) for l in oc.resolved_lines()
                     if l.kind == "profit"), 2)


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
        # 3a. (proprietorship) Net Profit must equal the source P&L net profit.
        #     Items the source posted directly to capital (FD/bank interest, LIC
        #     premium, etc.) stay in the capital account and are NOT reclassified
        #     into the P&L, so business Net Profit always ties to the T-format.
        if not firm:
            src_np = getattr(c, f"net_profit_{yr}") or 0.0
            plp = pl_profit(p, yr)
            if src_np and abs(plp - src_np) > EPS:
                d.append({"check": "Statement of P&L net profit does not match the source Profit & Loss net profit",
                          "year": ylabel, "expected": src_np, "got": plp, "diff": round(plp - src_np, 2)})
            # the capital-account 'Net Profit' line must equal the P&L profit
            cap_np = capital_profit_line(p, yr)
            if abs(cap_np - plp) > EPS:
                d.append({"check": "Net Profit shown in the Capital Account ≠ Statement of P&L profit",
                          "year": ylabel, "expected": plp, "got": cap_np, "diff": round(cap_np - plp, 2)})
        # 4. Case B depreciation reconciliation (CY only, per-asset modelled)
        if p.depreciation_case == "B" and p.ppe_assets and dep_for_year_total(p, "cy") > EPS:
            dep_sched = dep_for_year_total(p, yr) if yr == "cy" else _note_total(p, "depreciation", "py")
            dep_note = _note_total(p, "depreciation", yr)
            if yr == "cy" and abs(dep_sched - dep_note) > EPS:
                d.append({"check": "Depreciation for the year (PPE schedule) ≠ Depreciation note",
                          "year": ylabel, "expected": dep_sched, "got": dep_note,
                          "diff": round(dep_sched - dep_note, 2)})
        # 5. firm appropriation tie. Two valid presentations exist and we accept
        #    EITHER (flag only when neither ties):
        #    (1) interest on capital / remuneration are CHARGED in the P&L, so they
        #        are already inside Net Profit; the whole Net Profit is then shared
        #        by PSR  ->  sum(profit-share) == PAT.
        #    (2) interest / remuneration are appropriated BELOW Net Profit
        #        (book-profit style)  ->  interest + remuneration + share == PAT.
        if firm:
            int_t = sum(pt.kind_total("interest", yr) for pt in p.partners)
            rem_t = sum(pt.kind_total("remuneration", yr) for pt in p.partners)
            shr_t = sum(pt.kind_total("profit", yr) for pt in p.partners)
            tax = 0.0
            if p.firm_tax:
                tax = getattr(p.firm_tax, f"current_tax_{yr}") or 0.0
            pat = round(pl_profit(p, yr) - tax, 2)
            conv1 = abs(shr_t - pat)                       # interest/rem expensed in P&L
            conv2 = abs((int_t + rem_t + shr_t) - pat)     # interest/rem appropriated
            if min(conv1, conv2) > EPS:
                d.append({"check": "Appropriation does not tie (profit share, or interest + remuneration + share, must equal Profit after tax)",
                          "year": ylabel, "expected": pat, "got": round(shr_t, 2),
                          "diff": round(shr_t - pat, 2)})
        # Trading-account Gross Profit checkpoint: Sales - Cost of materials -
        # Direct expenses must equal the printed Gross Profit (if supplied).
        gp = getattr(c, f"gross_profit_{yr}") or 0.0
        if gp:
            calc_gp = round(_note_total(p, "revenue", yr) - _note_total(p, "cost_materials", yr)
                            - (getattr(c, f"direct_exp_{yr}") or 0.0), 2)
            if abs(calc_gp - gp) > EPS:
                d.append({"check": "Gross Profit does not tie (Sales - Cost of materials - Direct expenses != printed Gross Profit)",
                          "year": ylabel, "expected": gp, "got": calc_gp, "diff": round(calc_gp - gp, 2)})
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


def _rebuild_from_controls(p: Payload) -> List[str]:
    """Deterministically rebuild the P&L and the fixed-asset block from the PRINTED
    bold sub-totals in the Controls object. This removes reliance on the model
    correctly summing dozens of Trading-account sub-lines: the figures that MUST
    reconcile (revenue, other income, cost of materials, total expenses, depreciation,
    fixed-asset movement) come straight from the printed sub-totals the model read.
    Only runs for a year when those sub-totals are supplied (else leaves the year as-is)."""
    from models import Note, LineItem, PPEAsset
    c = p.controls
    fixes = []

    for yr in ("cy", "py"):
        # revenue / other income: prefer the printed control, else the model's note total
        rev = getattr(c, f"revenue_{yr}") or _note_total(p, "revenue", yr)
        oin = getattr(c, f"other_income_{yr}") or _note_total(p, "other_income", yr)
        dir_e = getattr(c, f"direct_exp_{yr}") or 0.0
        ind_e = getattr(c, f"indirect_exp_{yr}") or 0.0
        dep = getattr(c, f"depreciation_{yr}") or _note_total(p, "depreciation", yr)
        op = getattr(c, f"opening_stock_{yr}") or 0.0
        cl = getattr(c, f"closing_stock_{yr}") or 0.0
        pu = getattr(c, f"purchases_{yr}") or 0.0
        cost = round(op + pu - cl, 2) if (op or pu or cl) else _note_total(p, "cost_materials", yr)

        # Net Profit: prefer the printed control; else derive it from the capital-account
        # 'profit' lines (these tie to the source capital, which is independently checked).
        np_ = getattr(c, f"net_profit_{yr}") or 0.0
        if np_ <= 0:
            if p.entity.constitution == "partnership":
                np_ = round(sum(pt.kind_total("profit", yr) for pt in p.partners), 2)
            elif p.owner_capital:
                np_ = round(sum((getattr(l, yr) or 0.0) for l in p.owner_capital.resolved_lines()
                                if l.kind == "profit"), 2)

        # Decide TOTAL expenses:
        #  - if Net Profit + Revenue are known, anchor expenses so the P&L profit equals
        #    Net Profit EXACTLY (independent of how the model summed the sub-lines);
        #  - else fall back to the printed Direct + Indirect expense sub-totals.
        if np_ > 0 and rev > 0:
            total_exp = round(rev + oin - np_, 2)
        elif dir_e > 0 or ind_e > 0:
            total_exp = round(cost + dir_e + ind_e, 2)
        else:
            continue   # not enough reliable anchors for this year - leave as-is

        def _put(key, label, val):
            n = p.note(key)
            # preserve the OTHER year's figure label-agnostically (sum of existing items),
            # so renaming the note to the canonical label never wipes the other year.
            exist_cy = round(sum((i.cy or 0.0) for i in n.items), 2) if (n and n.items) else 0.0
            exist_py = round(sum((i.py or 0.0) for i in n.items), 2) if (n and n.items) else 0.0
            if n is None:
                n = Note(key=key, title=label); p.notes.append(n)
            if not n.title:
                n.title = label
            cyv = val if yr == "cy" else exist_cy
            pyv = val if yr == "py" else exist_py
            n.items = [LineItem(label=label, cy=round(cyv, 2), py=round(pyv, 2))]

        _put("revenue", "Revenue from operations", round(rev, 2))
        _put("other_income", "Other income", round(oin, 2))
        _put("cost_materials", "Cost of materials consumed", cost)
        ci = p.note("changes_inventory")
        if ci:
            ci.items = []
        if dep > 0:
            _put("depreciation", "Depreciation", round(dep, 2))

        # keep employee-benefit / finance-cost classification ONLY if sane, else drop
        emp = _note_total(p, "employee_benefits", yr)
        fin = _note_total(p, "finance_costs", yr)
        if emp < 0 or fin < 0 or (cost + dep + emp + fin) > total_exp + EPS:
            emp = fin = 0.0
            for k in ("employee_benefits", "finance_costs"):
                n = p.note(k)
                if n:
                    n.items = []
        # other_expenses is the residual that makes total expenses (hence profit) tie
        other = round(total_exp - cost - dep - emp - fin, 2)
        _put("other_expenses", "Other expenses", other)
        fixes.append(f"{yr.upper()}: P&L anchored to Net Profit {np_:,.0f} "
                     f"(revenue {rev:,.0f}); profit ties to source")

    # --- fixed-asset block movement from the two Balance-Sheet FA totals ---
    fa_cy = getattr(c, "fixed_assets_cy") or 0.0
    fa_py = getattr(c, "fixed_assets_py") or 0.0
    dep_cy = getattr(c, "depreciation_cy") or 0.0
    if fa_cy > 0 and fa_py > 0:
        additions = round(fa_cy - fa_py + dep_cy, 2)
        p.ppe_assets = [PPEAsset(name="Fixed Assets (net block)", rate="",
                                 gb_open=round(fa_py, 2), additions=max(0.0, additions),
                                 accdep_open=0.0, dep_year=round(dep_cy, 2))]
        # if additions came out negative (net disposals), fold the sign into dep_year
        if additions < 0:
            p.ppe_assets[0].additions = 0.0
            p.ppe_assets[0].dep_year = round(dep_cy - additions, 2)
        p.depreciation_case = "A"   # movement modelled at block level, skip per-asset check
        fixes.append(f"Fixed assets reconciled via movement: open {fa_py:,.0f} + additions "
                     f"- depreciation {dep_cy:,.0f} = close {fa_cy:,.0f}")
    return fixes


def _clean_capital_lines(p: Payload) -> List[str]:
    """A T-format capital account lists balancing rows ("To Closing Balance", "Total",
    "Balance c/d") that are RESULTS, not movements. If extraction captured those as
    capital lines (or merged both years' accounts) the signed closing balance balloons.
    The builder computes the closing itself, so strip any balancing/total row here."""
    drop = {"closing balance", "closing bal", "closing capital", "balance cd",
            "balance carried down", "balance carried forward", "total", "grand total",
            "total rs", "balance"}
    def _norm(lbl: str) -> str:
        t = (lbl or "").strip().lower()
        for pre in ("to ", "by "):
            if t.startswith(pre):
                t = t[len(pre):]
        t = t.replace(".", "").replace("/", "").replace(":", "")
        return " ".join(t.split())
    holders = ([p.owner_capital] if p.owner_capital else []) + list(p.partners or [])
    removed = 0
    for h in holders:
        lines = getattr(h, "lines", None)
        if not lines:
            continue
        kept = [l for l in lines if _norm(l.label) not in drop]
        removed += len(lines) - len(kept)
        h.lines = kept
    return [f"Removed {removed} balancing/total row(s) wrongly captured in the capital "
            f"account(s); closing balance is computed by the engine"] if removed else []


def _anchor_capital(p: Payload) -> List[str]:
    """Tie each capital account to its PRINTED CLOSING BALANCE. The year's movements
    (profit share, interest, introductions, drawings) are individually reliable; opening
    is the balancing figure the model most often misreads (e.g. a 10x slip). So:
      opening = printed closing - sum(movements)   [per partner, when closing is given]
    Falls back to tying the FIRM total to the printed capital control (distributing the
    correction across the openings that have no per-partner closing). Guarantees capital
    ties to source and recovers correct openings even under non-uniform slips."""
    c = p.controls
    holders = ([p.owner_capital] if p.owner_capital else []) + list(p.partners or [])
    if not holders:
        return []
    fixes = []
    for yr in ("cy", "py"):
        anchored = 0
        anchored_close = 0.0
        free_opens = []          # opening lines of holders without a printed closing
        free_move = 0.0
        for h in holders:
            closing = getattr(h, f"closing_{yr}", 0.0) or 0.0
            opens = [l for l in h.resolved_lines() if l.kind == "opening"]
            move = sum(l.signed(yr) for l in h.resolved_lines() if l.kind != "opening")
            if closing > 0 and opens:
                needed = round(closing - move, 2)        # opening that makes this account tie
                cur = sum((getattr(l, yr) or 0.0) for l in opens)
                if abs(cur - needed) > EPS:
                    if abs(cur) > EPS:
                        sc = needed / cur
                        for l in opens:
                            setattr(l, yr, round((getattr(l, yr) or 0.0) * sc, 2))
                    else:
                        setattr(opens[0], yr, needed)
                    anchored += 1
                anchored_close += closing
            else:
                free_opens += opens
                free_move += move
        if anchored:
            fixes.append(f"{yr.upper()}: {anchored} capital account(s) anchored to their "
                         f"printed closing balance")
        # firm-total fallback: tie the overall capital to the printed control using the
        # openings that were NOT individually anchored.
        target = getattr(c, f"capital_close_{yr}") or 0.0
        if target <= 0 or not free_opens:
            continue
        needed_free = round(target - anchored_close - free_move, 2)
        cur_free = sum((getattr(l, yr) or 0.0) for l in free_opens)
        if abs(cur_free) < EPS or abs(cur_free - needed_free) <= EPS:
            continue
        sc = needed_free / cur_free
        for l in free_opens:
            setattr(l, yr, round((getattr(l, yr) or 0.0) * sc, 2))
        fixes.append(f"{yr.upper()}: capital total tied to printed capital {target:,.0f} "
                     f"(factor {sc:.4g})")
    return fixes


def _anchor_inventories(p: Payload) -> List[str]:
    """Inventories must equal the printed Closing Stock for each year. If the extracted
    inventories note does not tie to the closing-stock control (e.g. stock got
    double-counted, or opening stock was wrongly included as an asset), reset it to a
    single Closing Stock line equal to the control. Deterministic; only acts when wrong."""
    from models import Note, LineItem
    c = p.controls
    cs_cy = getattr(c, "closing_stock_cy") or 0.0
    cs_py = getattr(c, "closing_stock_py") or 0.0
    if cs_cy <= 0 and cs_py <= 0:
        return []
    cur_cy = _note_total(p, "inventories", "cy")
    cur_py = _note_total(p, "inventories", "py")
    if abs(cur_cy - cs_cy) <= EPS and abs(cur_py - cs_py) <= EPS:
        return []                      # already correct - leave the breakup as-is
    n = p.note("inventories")
    if n is None:
        n = Note(key="inventories", title="Inventories"); p.notes.append(n)
    n.items = [LineItem(label="Closing Stock (at cost or NRV, whichever is lower)",
                        cy=round(cs_cy, 2), py=round(cs_py, 2))]
    return [f"Inventories anchored to printed Closing Stock (CY {cs_cy:,.0f}, PY {cs_py:,.0f}); "
            f"removed any double-count / mis-inclusion of stock"]


def _expense_head(label: str) -> str:
    """Deterministically map a P&L expense ledger to its NCE head, by rulebook keywords.
    Order matters: depreciation, then interest-on-LATE-payment (-> other), then finance,
    then employee benefits, else other expenses."""
    t = " " + (label or "").lower() + " "
    def has(words): return any(re.search(r"\b" + re.escape(w) + r"\b", t) for w in words)
    if has(["depreciation", "deprecation", "amortization", "amortisation"]):
        return "depreciation"
    if ("late payment" in t) or ("late fee" in t) or ("penalty" in t) or ("penal " in t):
        return "other_expenses"           # interest on late payment of TDS/GST/IT -> Other
    if has(["interest", "bank charge", "bank charges", "finance cost", "finance costs",
            "loan processing", "processing fee", "commitment charge"]):
        return "finance_costs"
    if has(["salary", "wages", "wage", "bonus", "staff welfare", "staff medical",
            "provident fund", "pf", "epf", "esic", "gratuity", "labour", "labor",
            "remuneration", "stipend", "ex gratia"]):
        return "employee_benefits"
    return "other_expenses"


def _classify_expenses(p: Payload) -> List[str]:
    """Re-bucket every expense ledger into Employee Benefits / Finance Costs /
    Depreciation / Other Expenses by the deterministic rulebook keywords - so the P&L
    classification is identical every run, independent of how the model first split them.
    Cost of materials, revenue and other income are left untouched (anchored elsewhere)."""
    from models import Note
    keys = ["employee_benefits", "finance_costs", "depreciation", "other_expenses"]
    pool = []
    for k in keys:
        n = p.note(k)
        if n and n.items:
            pool += list(n.items)
            n.items = []
    if not pool:
        return []
    buckets = {k: [] for k in keys}
    for it in pool:
        buckets[_expense_head(it.label)].append(it)
    titles = {"employee_benefits": "Employee Benefits Expense", "finance_costs": "Finance Costs",
              "depreciation": "Depreciation and Amortisation Expense", "other_expenses": "Other Expenses"}
    moved = 0
    for k in keys:
        n = p.note(k)
        if n is None:
            n = Note(key=k, title=titles[k]); p.notes.append(n)
        n.items = buckets[k]
        moved += len(buckets[k])
    return [f"Expenses re-classified deterministically by rulebook into Employee Benefits / "
            f"Finance Costs / Depreciation / Other ({moved} ledgers)"]


_GROUP_ANCHOR_KEYS = [
    # liabilities / equity groups
    "reserves", "lt_borrowings", "st_borrowings", "trade_payables", "other_cl",
    "st_provisions", "lt_provisions", "other_lt_liab", "dtl",
    # asset groups
    "trade_receivables", "cash_bank", "st_loans_adv", "other_ca",
    "nc_investments", "current_investments", "lt_loans_adv", "other_nca", "dta",
]


def _anchor_groups(p: Payload) -> List[str]:
    """Tie every Balance-Sheet group note to its PRINTED FACE TOTAL (from group_totals).
    A single-line group is set to the total; a multi-line group keeps its breakup and
    gets a clearly-labelled balancing line for any difference (dropped line, double-count
    or misread) so the Balance Sheet always tallies and the gap is surfaced for review.
    Deterministic; only acts when a group does not already tie to its printed total."""
    from models import Note, LineItem
    gt = {g.key: g for g in (p.controls.group_totals or []) if getattr(g, "key", "")}
    if not gt:
        return []
    fixes = []
    for key in _GROUP_ANCHOR_KEYS:
        g = gt.get(key)
        if not g:
            continue
        t_cy = getattr(g, "cy", 0.0) or 0.0
        t_py = getattr(g, "py", 0.0) or 0.0
        if t_cy <= 0 and t_py <= 0:
            continue
        n = p.note(key)
        cur_cy = _note_total(p, key, "cy") if n else 0.0
        cur_py = _note_total(p, key, "py") if n else 0.0
        d_cy = round(t_cy - cur_cy, 2) if t_cy > 0 else 0.0
        d_py = round(t_py - cur_py, 2) if t_py > 0 else 0.0
        if abs(d_cy) <= EPS and abs(d_py) <= EPS:
            continue                         # already ties to printed total
        if n is None:
            n = Note(key=key, title=key.replace("_", " ").title()); p.notes.append(n)
        if len(n.items) <= 1:
            # single-line group: set it straight to the printed total
            label = n.items[0].label if n.items else "As per source"
            n.items = [LineItem(label=label,
                                cy=round(t_cy, 2) if t_cy > 0 else (n.items[0].cy if n.items else 0.0),
                                py=round(t_py, 2) if t_py > 0 else (n.items[0].py if n.items else 0.0))]
            fixes.append(f"{key}: set to printed group total (CY {t_cy:,.0f}, PY {t_py:,.0f})")
        else:
            # multi-line group: preserve breakup, add a balancing line for the difference
            n.items.append(LineItem(label="Balancing difference vs printed group total (please verify)",
                                    cy=d_cy if abs(d_cy) > EPS else 0.0,
                                    py=d_py if abs(d_py) > EPS else 0.0))
            fixes.append(f"{key}: balanced to printed group total (CY diff {d_cy:,.0f}, PY diff {d_py:,.0f}) - flagged")
    return fixes


def auto_fix(p: Payload) -> List[str]:
    """Apply safe, rule-based corrections. Returns a list of fixes applied."""
    fixes = []
    fixes += _clean_capital_lines(p)
    fixes += _anchor_capital(p)
    fixes += _anchor_inventories(p)
    fixes += _anchor_groups(p)
    fixes += _classify_expenses(p)
    fixes += _rebuild_from_controls(p)
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
    # NOTE: income/expense items the source posted directly in the Capital
    # Account (FD/bank interest, LIC premium, personal drawings, etc.) are kept
    # in the Capital Account verbatim and are deliberately NOT moved into the
    # P&L. No auto-reclassification is performed on capital-account lines.
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

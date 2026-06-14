"""
Data model — the contract between the LLM extraction layer and the openpyxl
workbook engine. Pure stdlib dataclasses (no third-party dependency) plus a
from_dict parser so the LLM can return plain JSON.

All amounts are floats (Rupees, two decimals). CY = current year, PY = previous year.
"""
from __future__ import annotations
import typing
from dataclasses import dataclass, field, fields, is_dataclass
from typing import List, Optional, Union, get_origin, get_args, get_type_hints


# ---- parsing helper ---------------------------------------------------------
def _coerce(tp, val):
    if val is None:
        return None
    origin = get_origin(tp)
    if origin is list:
        (inner,) = get_args(tp)
        return [_coerce(inner, v) for v in (val or [])]
    if origin is Union:  # Optional[X]
        args = [a for a in get_args(tp) if a is not type(None)]
        return _coerce(args[0], val) if args else val
    if is_dataclass(tp):
        return from_dict(tp, val)
    if tp is float:
        try:
            return float(val)
        except (TypeError, ValueError):
            return 0.0
    if tp is int:
        try:
            return int(val)
        except (TypeError, ValueError):
            return 0
    return val


def from_dict(cls, data: dict):
    if data is None:
        return None
    hints = get_type_hints(cls)
    kwargs = {}
    for f in fields(cls):
        if f.name in data:
            kwargs[f.name] = _coerce(hints.get(f.name, f.type), data[f.name])
    return cls(**kwargs)


# ICAI Schedule III note keys (canonical sequence used for numbering)
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


@dataclass
class EntityInfo:
    name: str = ""
    constitution: str = "proprietorship"   # or "partnership"
    cy_label: str = "31 March 20XX"
    py_label: str = "31 March 20XX"
    cy_fy: str = ""
    nature_of_business: str = ""
    nce_level: str = "IV"
    pan: str = "—"
    address: str = "—"
    gstin: str = "—"
    date_of_commencement: str = "—"
    ca_firm: str = "M S Joshi & Co"
    frn: str = "0138082W"
    ca_partner: str = "CA Mehulkumar S. Joshi"
    membership_no: str = "152333"
    udin: str = "—"
    place: str = "—"
    report_date: str = "—"
    deed_date: str = "—"


@dataclass
class LineItem:
    label: str = ""
    cy: float = 0.0
    py: float = 0.0


# ---- Capital Account line model --------------------------------------------
# A capital account is now a free list of lines, mirroring the T-format exactly.
# Each line carries a semantic "kind" that controls (a) whether it is added or
# subtracted to reach the closing balance and (b) the Add/Less prefix shown.
#   added   : opening, profit, introduced, interest, remuneration, add
#   subtracted: withdrawals, less
ADD_KINDS = {"opening", "profit", "introduced", "interest", "remuneration", "add"}
SUB_KINDS = {"withdrawals", "less"}


@dataclass
class CapitalLine:
    label: str = ""
    kind: str = "add"          # see ADD_KINDS / SUB_KINDS
    cy: float = 0.0
    py: float = 0.0

    def signed(self, yr: str) -> float:
        v = getattr(self, yr) or 0.0
        return -v if self.kind in SUB_KINDS else v

    def prefix(self) -> str:
        if self.kind == "opening":
            return ""
        return "Less" if self.kind in SUB_KINDS else "Add"


@dataclass
class SubNote:
    """A numbered sub-section of a note (rendered as N.1, N.2, ...).
    Use for movement schedules, AS-29 disclosures, sub-group breakups, etc.
    If `items` is given it renders as its own small table (no auto-total -
    any closing/total must be supplied as an item). `footnotes` render as prose."""
    title: str = ""
    items: List[LineItem] = field(default_factory=list)
    footnotes: List[str] = field(default_factory=list)


@dataclass
class Note:
    key: str = ""
    title: str = ""
    items: List[LineItem] = field(default_factory=list)
    footnotes: List[str] = field(default_factory=list)
    subnotes: List["SubNote"] = field(default_factory=list)

    def total_cy(self) -> float:
        return round(sum(i.cy for i in self.items), 2)

    def total_py(self) -> float:
        return round(sum(i.py for i in self.items), 2)

    def is_nil(self) -> bool:
        return abs(self.total_cy()) < 0.005 and abs(self.total_py()) < 0.005


@dataclass
class OwnerCapital:
    name: str = ""
    # Preferred: a verbatim list of the proprietor's capital-account lines.
    lines: List[CapitalLine] = field(default_factory=list)
    # Legacy scalar fields (used only as a fallback when `lines` is empty).
    opening_cy: float = 0.0
    introduced_cy: float = 0.0
    net_profit_cy: float = 0.0
    interest_cy: float = 0.0
    withdrawals_cy: float = 0.0
    opening_py: float = 0.0
    introduced_py: float = 0.0
    net_profit_py: float = 0.0
    interest_py: float = 0.0
    withdrawals_py: float = 0.0

    def resolved_lines(self) -> List[CapitalLine]:
        if self.lines:
            return self.lines
        out = [CapitalLine("Opening Balance", "opening", self.opening_cy, self.opening_py)]
        if self.introduced_cy or self.introduced_py:
            out.append(CapitalLine("Capital Introduced", "introduced",
                                   self.introduced_cy, self.introduced_py))
        out.append(CapitalLine("Net Profit for the year", "profit",
                               self.net_profit_cy, self.net_profit_py))
        if self.interest_cy or self.interest_py:
            out.append(CapitalLine("Interest on Capital", "interest",
                                   self.interest_cy, self.interest_py))
        out.append(CapitalLine("Withdrawals", "withdrawals",
                               self.withdrawals_cy, self.withdrawals_py))
        return out


@dataclass
class PartnerRow:
    name: str = ""
    psr: float = 0.0
    # Preferred: a verbatim list of this partner's capital-account lines.
    lines: List[CapitalLine] = field(default_factory=list)
    # Legacy scalar fields (used only as a fallback when `lines` is empty).
    opening_cy: float = 0.0
    introduced_cy: float = 0.0
    share_profit_cy: float = 0.0
    interest_cy: float = 0.0
    remuneration_cy: float = 0.0
    withdrawals_cy: float = 0.0
    opening_py: float = 0.0
    introduced_py: float = 0.0
    share_profit_py: float = 0.0
    interest_py: float = 0.0
    remuneration_py: float = 0.0
    withdrawals_py: float = 0.0

    def resolved_lines(self) -> List[CapitalLine]:
        if self.lines:
            return self.lines
        out = [CapitalLine("Opening Balance", "opening", self.opening_cy, self.opening_py)]
        if self.introduced_cy or self.introduced_py:
            out.append(CapitalLine("Capital Introduced", "introduced",
                                   self.introduced_cy, self.introduced_py))
        out.append(CapitalLine("Share of Profit", "profit",
                               self.share_profit_cy, self.share_profit_py))
        if self.interest_cy or self.interest_py:
            out.append(CapitalLine("Interest on Capital", "interest",
                                   self.interest_cy, self.interest_py))
        if self.remuneration_cy or self.remuneration_py:
            out.append(CapitalLine("Remuneration", "remuneration",
                                   self.remuneration_cy, self.remuneration_py))
        out.append(CapitalLine("Withdrawals", "withdrawals",
                               self.withdrawals_cy, self.withdrawals_py))
        return out

    def kind_total(self, kind: str, yr: str) -> float:
        return round(sum((getattr(l, yr) or 0.0) for l in self.resolved_lines()
                         if l.kind == kind), 2)


@dataclass
class PPEAsset:
    name: str = ""
    rate: str = ""
    gb_open: float = 0.0
    additions: float = 0.0
    accdep_open: float = 0.0
    dep_year: float = 0.0


@dataclass
class FirmTax:
    current_tax_cy: float = 0.0
    current_tax_py: float = 0.0
    restatement_note: str = ""


@dataclass
class PLMeta:
    net_profit_cy: float = 0.0
    net_profit_py: float = 0.0


@dataclass
class Controls:
    """Source control totals read directly from the printed source statements.
    Used by the reconciliation engine to assert the converted output matches source."""
    bs_total_cy: float = 0.0
    bs_total_py: float = 0.0
    capital_close_cy: float = 0.0
    capital_close_py: float = 0.0
    net_profit_cy: float = 0.0
    net_profit_py: float = 0.0
    opening_stock_cy: float = 0.0
    opening_stock_py: float = 0.0
    closing_stock_cy: float = 0.0
    closing_stock_py: float = 0.0
    purchases_cy: float = 0.0
    purchases_py: float = 0.0


@dataclass
class Payload:
    entity: EntityInfo = field(default_factory=EntityInfo)
    notes: List[Note] = field(default_factory=list)
    owner_capital: Optional[OwnerCapital] = None
    partners: List[PartnerRow] = field(default_factory=list)
    ppe_assets: List[PPEAsset] = field(default_factory=list)
    depreciation_case: str = "A"
    firm_tax: Optional[FirmTax] = None
    pl: PLMeta = field(default_factory=PLMeta)
    controls: Controls = field(default_factory=Controls)

    def note(self, key: str) -> Optional[Note]:
        for n in self.notes:
            if n.key == key:
                return n
        return None

    @staticmethod
    def parse(data: dict) -> "Payload":
        return from_dict(Payload, data)

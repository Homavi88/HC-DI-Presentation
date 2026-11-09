#!/usr/bin/env python3
"""Rebuild slide-6 PARETO_DATA in hc_di_deck_v2_single.html from HC checks Excel.

Example:
  python3 tools/hc_deck_pareto.py \\
    --excel hc_checks_Secondary_20261001_20261031.xlsx \\
    --apply
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError as exc:  # pragma: no cover
    raise SystemExit("openpyxl is required: pip install openpyxl") from exc

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_DECK = ROOT / "hc_di_deck_v2_single.html"

SHORT = {
    1: "1440 мин / нет",
    2: "399 не назначен",
    3: "600 ручн. корректир.",
    4: "611 старт",
    5: "611 остановок",
    6: "602/609 меж PO",
    7: "ремонт 603/604",
    8: "608 IFMS",
    9: "скорость ≤ проект",
    10: "пит-стоп / CO — CIL",
    11: "причина план. остановок",
    12: "690 Unavoidable",
}

MONTH_RU = {
    1: "январь",
    2: "февраль",
    3: "март",
    4: "апрель",
    5: "май",
    6: "июнь",
    7: "июль",
    8: "август",
    9: "сентябрь",
    10: "октябрь",
    11: "ноябрь",
    12: "декабрь",
}
MONTH_SHORT = {
    1: "Янв",
    2: "Фев",
    3: "Мар",
    4: "Апр",
    5: "Май",
    6: "Июн",
    7: "Июл",
    8: "Авг",
    9: "Сен",
    10: "Окт",
    11: "Ноя",
    12: "Дек",
}


def _parse_period(text: str) -> tuple[str, str, int]:
    """Return display period, Russian label, and month count."""
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})\s*[—-]\s*(\d{4})-(\d{2})-(\d{2})", text or "")
    if not m:
        raise ValueError(f"Cannot parse period from {text!r}")
    y1, mo1, d1, y2, mo2, d2 = map(int, m.groups())
    period = f"{d1:02d}.{mo1:02d}.{y1}–{d2:02d}.{mo2:02d}.{y2}"
    months = (y2 - y1) * 12 + (mo2 - mo1) + 1
    if months == 1 and y1 == y2 and mo1 == mo2:
        label = f"{MONTH_RU[mo1]} {y1}"
    else:
        label = period
    return period, label, months


def build_pareto(excel_path: Path) -> dict:
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    period_raw = wb["Сводка"]["B2"].value
    period, period_label, months = _parse_period(str(period_raw))

    machines = []
    for row in wb["Сводка"].iter_rows(min_row=8, values_only=True):
        if not row or not row[1]:
            continue
        machines.append(
            {
                "name": row[1],
                "strict": int(row[4]),
                "reg": int(row[5]),
                "k": int(row[6]),
            }
        )
    if not machines:
        raise ValueError("No machines found on Сводка")

    m = sorted([[x["name"], x["strict"], x["k"]] for x in machines], key=lambda t: (-t[1], t[0]))
    m_p = sorted([[x["name"], x["reg"]] for x in machines], key=lambda t: (-t[1], t[0]))

    qrows = []
    for row in wb["Анализ (деп.)"].iter_rows(min_row=5, max_row=16, values_only=True):
        if row[0] is None:
            continue
        qn = int(row[0])
        qrows.append([f"Q{qn}", SHORT[qn], int(row[3])])
    q = sorted(qrows, key=lambda t: (-t[2], int(t[0][1:])))
    total_fails = sum(x[2] for x in q)
    top4 = q[:4]
    top4_pct = round(100 * sum(x[2] for x in top4) / total_fails) if total_fails else 0

    def fails_for(mach: str, qn: int) -> int:
        ws = wb[f"{mach} — Анализ HC"]
        for row in ws.iter_rows(min_row=5, max_row=16, values_only=True):
            if row[0] == qn:
                return int(row[3])
        return 0

    qm = []
    for qid, text, tot in top4:
        qn = int(qid[1:])
        per = []
        for mach in (x["name"] for x in machines):
            f = fails_for(mach, qn)
            if f > 0:
                per.append([mach, f])
        per.sort(key=lambda t: (-t[1], t[0]))
        got = sum(x[1] for x in per)
        if got != tot:
            raise ValueError(f"{qid}: machine fails sum {got} != department {tot}")
        qm.append([qid, text, tot, per])

    itog = str(wb["Департамент"]["B4"].value or "")
    mpark = re.search(
        r"строгий HC DI:\s*(\d+)%\s*·\s*HC DI \(период\):\s*(\d+)%\s*·\s*вопросо-дни:\s*(\d+)/(\d+)\s*·\s*машин:\s*(\d+)",
        itog,
    )
    if not mpark:
        raise ValueError(f"Cannot parse department summary: {itog!r}")
    park_strict = int(mpark.group(1))
    park_reg = int(mpark.group(2))
    passed, total_qd, nmach = map(int, mpark.groups()[2:])
    if nmach != len(machines):
        raise ValueError(f"machine count mismatch: summary {nmach} vs Сводка {len(machines)}")
    if total_fails != total_qd - passed:
        raise ValueError(
            f"fail count mismatch: analysis {total_fails} vs question-days {total_qd - passed}"
        )

    mo = int(re.search(r"-(\d{2})-", str(period_raw)).group(1))
    return {
        "m": m,
        "mP": m_p,
        "q": q,
        "qm": qm,
        "mth": [[MONTH_SHORT[mo], park_strict, park_reg]],
        "meta": {
            "period": period,
            "periodLabel": period_label,
            "months": months,
            "nmach": nmach,
            "parkStrict": park_strict,
            "parkReg": park_reg,
            "totalFails": total_fails,
            "top4Pct": top4_pct,
        },
    }


def apply_to_deck(deck_path: Path, pareto: dict) -> None:
    html = deck_path.read_text(encoding="utf-8")
    marker_s = "/*PARETO_DATA*/"
    marker_e = "/*END_PARETO*/"
    i0 = html.find(marker_s)
    if i0 < 0:
        raise ValueError(f"Missing {marker_s} in {deck_path}")
    j0 = i0 + len(marker_s)
    j1 = html.find(marker_e, j0)
    if j1 < 0:
        raise ValueError(f"Missing {marker_e} in {deck_path}")
    raw = json.dumps(pareto, ensure_ascii=False, separators=(",", ":"))
    deck_path.write_text(html[:j0] + raw + html[j1:], encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--excel", type=Path, required=True, help="HC checks .xlsx")
    p.add_argument("--deck", type=Path, default=DEFAULT_DECK)
    p.add_argument("--apply", action="store_true", help="Write PARETO_DATA into the deck HTML")
    p.add_argument("--months", type=int, default=None, help="Override months count in meta")
    args = p.parse_args(argv)

    pareto = build_pareto(args.excel)
    if args.months is not None:
        pareto["meta"]["months"] = args.months

    print(json.dumps(pareto["meta"], ensure_ascii=False, indent=2))
    print("mP:", pareto["mP"])
    print("q top4:", pareto["q"][:4])

    if args.apply:
        apply_to_deck(args.deck, pareto)
        print(f"Applied to {args.deck}")
    else:
        print("Dry-run only (pass --apply to write).")
    return 0


if __name__ == "__main__":
    sys.exit(main())

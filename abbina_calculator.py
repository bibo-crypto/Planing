"""
abbina_calculator.py — Automatic Abbina assignment for Delta Dyeing order rows.

Business rules (from spec)
--------------------------
1.  Never merge rows with different Colour values.
2.  Only consecutive rows with the same Colour may form one machine group.
3.  Colour comparison is exact (case-sensitive, leading-zeros preserved).
4.  Walk rows top → bottom, grouping while Colour stays identical.
5.  Sum Quantity /Cones for the group.
6.  Assign the *smallest* available machine capacity that is >= the group total.
7.  Write "Machine <capacity> Cones" into every row of that group.

Available machine capacities (ascending order — must stay sorted):
    6, 24, 32, 56, 72, 128, 192, 384, 672

If the group total exceeds 672 (the largest machine), the largest machine is
assigned and a warning is logged.  The caller is responsible for deciding
whether to split oversized groups.
"""

from __future__ import annotations

from pdf_parser import OrderRow
from utils import logger
from constants import ABBINA_MACHINE_CODES, MACHINE_CAPACITIES


# ---------------------------------------------------------------------------
# Machine capacity table — must remain sorted ascending
# ---------------------------------------------------------------------------

# Backward-compatible alias; the canonical values live in constants.py.
MACHINE_CAPACITIES: list[int] = list(MACHINE_CAPACITIES)

# Physical machine number for each cone capacity — used by the "Ordini
# ELVY" export sheet to turn an Abbina value like "Machine 72 Cones" into
# a MACCHINA code.
MACHINE_CODES: dict[int, int] = dict(ABBINA_MACHINE_CODES)


def _smallest_fitting_machine(total_cones: float) -> int:
    """Return 0 for an empty group, otherwise the smallest fitting capacity."""
    if total_cones <= 0:
        return 0
    for capacity in MACHINE_CAPACITIES:
        if capacity >= total_cones:
            return capacity
    # total_cones exceeds every machine — use the largest and warn
    logger.warning(
        "    Group total %.2f cones exceeds largest machine (%d). "
        "Assigning largest machine.",
        total_cones, MACHINE_CAPACITIES[-1],
    )
    return MACHINE_CAPACITIES[-1]


# ---------------------------------------------------------------------------
# Calculator
# ---------------------------------------------------------------------------

class AbbinaCalculator:
    """
    Assign Abbina values to a list of :class:`~pdf_parser.OrderRow` objects
    **in-place** using the automatic grouping rules.

    Usage::

        rows = PDFParser(path).parse()
        AbbinaCalculator().calculate(rows)
        # rows[i].abbina is now set for every row
    """

    def calculate(self, rows: list[OrderRow], only_if_missing: bool = False) -> None:
        """
        Walk *rows* and set ``abbina`` on each row.

        Groups are formed by consecutive rows that share the same ``colour``
        value (exact string match).  Each group is assigned the smallest
        machine that can hold its total ``quantity_cones``.

        only_if_missing
        ----------------
        When True, a group is only touched if NONE of its rows already have
        an ``abbina`` value (e.g. extracted from a PDF annotation) — groups
        that already have one are left exactly as they are. This powers the
        mandatory hybrid behaviour: PDF-provided Abbina wins where present,
        automatic calculation fills in the gaps.
        """
        if not rows:
            return

        i = 0
        while i < len(rows):
            colour = rows[i].colour
            group_start = i
            total_cones: float = 0.0

            # Extend the group while colour stays identical
            while i < len(rows) and rows[i].colour == colour:
                total_cones += rows[i].quantity_cones or 0.0
                i += 1

            group_end = i  # exclusive slice index

            if only_if_missing and any(rows[j].abbina for j in range(group_start, group_end)):
                continue

            machine_capacity = _smallest_fitting_machine(total_cones)
            label = f"Machine {machine_capacity} Cones"

            for j in range(group_start, group_end):
                rows[j].abbina = label

            logger.debug(
                "    Abbina group colour=%r rows %d–%d  "
                "total=%.2f → %s",
                colour, group_start + 1, group_end,
                total_cones, label,
            )

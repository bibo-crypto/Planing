"""Shared business constants used by the planning and export modules."""

_DEFAULT_MACHINE_ROWS = (
    ("3301", "11", "6"), ("3310", "12", "24"), ("3306", "9", "32"),
    ("3302", "10", "56"), ("3307", "7", "72"), ("3303", "8", "128"),
    ("3308", "5", "192"), ("3304", "6", "384"), ("3309", "3", "672"),
)

try:
    from utility.master_data import load as _load_master_data
    _machine_rows = _load_master_data().get("machines", [])
    _machine_rows = tuple((str(r["code"]), str(r["number"]), int(r["rocche"])) for r in _machine_rows)
except (KeyError, TypeError, ValueError, OSError):
    _machine_rows = ()
if not _machine_rows:
    _machine_rows = tuple((code, number, int(rocche)) for code, number, rocche in _DEFAULT_MACHINE_ROWS)

# Capacity values are ordered from the smallest to the largest machine.
MACHINE_CAPACITIES: tuple[int, ...] = tuple(sorted(row[2] for row in _machine_rows))

# Machine identifiers depend on the output domain, so they remain separate.
ABBINA_MACHINE_CODES: dict[int, int] = {capacity: int(code) for code, _number, capacity in _machine_rows}
SUGGESTION_MACHINE_NUMBERS: dict[int, int] = {capacity: int(number) for _code, number, capacity in _machine_rows}

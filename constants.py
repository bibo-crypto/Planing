"""Shared business constants used by the planning and export modules."""

# Capacity values are ordered from the smallest to the largest machine.
MACHINE_CAPACITIES: tuple[int, ...] = (6, 24, 32, 56, 72, 128, 192, 384, 672)

# Machine identifiers depend on the output domain, so they remain separate.
ABBINA_MACHINE_CODES: dict[int, int] = {
    6: 3301, 24: 3310, 32: 3306, 56: 3302, 72: 3307,
    128: 3303, 192: 3308, 384: 3304, 672: 3305,
}
SUGGESTION_MACHINE_NUMBERS: dict[int, int] = {
    6: 11, 24: 12, 32: 9, 56: 10, 72: 7,
    128: 8, 192: 5, 384: 6, 672: 3,
}

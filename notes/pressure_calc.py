"""Duct leakage calculator.

Formula:
	Q_allow = CL * P^0.65 * (A / 100)

Where:
	Q_allow: allowed leakage (CFM)
	CL: leakage class constant
	P: test pressure (in. w.g.)
	A: tested duct surface area (ft^2)
"""

from __future__ import annotations

import math


# Update these values if your company standard uses a different class mapping.
SEAL_CLASS_TO_LEAKAGE_CLASS = {
    "A": 2.0,
    "B": 4.0,
    "C": 6.0,
}


def round_surface_area_per_foot(diameter_in: float) -> float:
    """Return round duct surface area per linear foot (ft^2/ft)."""
    return math.pi * (diameter_in / 12.0)


def total_round_surface_area(diameter_in: float, length_ft: float) -> float:
    """Return total round duct surface area (ft^2)."""
    return round_surface_area_per_foot(diameter_in) * length_ft


def rectangular_surface_area_per_foot(width_in: float, height_in: float) -> float:
    """Return rectangular duct surface area per linear foot (ft^2/ft)."""
    return 2.0 * ((width_in / 12.0) + (height_in / 12.0))


def total_rectangular_surface_area(width_in: float, height_in: float, length_ft: float) -> float:
    """Return total rectangular duct surface area (ft^2)."""
    return rectangular_surface_area_per_foot(width_in, height_in) * length_ft


def allowed_leakage_cfm(leakage_class: float, pressure_in_wg: float, area_ft2: float) -> float:
    """Compute allowed duct leakage in CFM."""
    return leakage_class * (pressure_in_wg ** 0.65) * (area_ft2 / 100.0)


def _read_positive_float(prompt: str) -> float:
    while True:
        raw = input(prompt).strip()
        try:
            value = float(raw)
            if value <= 0:
                print("Value must be greater than 0.")
                continue
            return value
        except ValueError:
            print("Enter a valid number.")


def _read_seal_class(prompt: str) -> str:
    while True:
        raw = input(prompt).strip().upper()
        if raw in SEAL_CLASS_TO_LEAKAGE_CLASS:
            return raw
        print("Enter A, B, or C.")


def _read_shape(prompt: str) -> str:
    while True:
        raw = input(prompt).strip().lower()
        if raw in {"round", "r"}:
            return "round"
        if raw in {"square", "s", "rect", "rectangle", "rectangular"}:
            return "square"
        print("Enter round or square.")


def main() -> None:
    print("Duct Leakage Calculator")
    print("Formula: Q_allow = CL * P^0.65 * (A / 100)")
    print()
    shape = _read_shape("Duct shape (round or square): ")

    if shape == "round":
        diameter_in = _read_positive_float("Round duct diameter (in): ")
        sfa_per_ft = round_surface_area_per_foot(diameter_in)
    else:
        width_in = _read_positive_float("Square/rect width (in): ")
        height_in = _read_positive_float("Square/rect height (in): ")
        sfa_per_ft = rectangular_surface_area_per_foot(width_in, height_in)

    length_ft = _read_positive_float("Duct length (ft): ")
    area_ft2 = sfa_per_ft * length_ft

    seal_class = _read_seal_class("Seal class (A, B, or C): ")
    leakage_class = SEAL_CLASS_TO_LEAKAGE_CLASS[seal_class]
    pressure_in_wg = _read_positive_float("Test pressure (in. w.g., e.g. 4): ")

    if shape == "round":
        area_ft2 = total_round_surface_area(diameter_in, length_ft)
    else:
        area_ft2 = total_rectangular_surface_area(
            width_in, height_in, length_ft)

    leakage_cfm = allowed_leakage_cfm(leakage_class, pressure_in_wg, area_ft2)

    print("\nResults")
    print("-------")
    print(f"Shape: {shape}")
    if shape == "round":
        print(f"Diameter: {diameter_in:.2f} in")
    else:
        print(f"Width x Height: {width_in:.2f} in x {height_in:.2f} in")

    print(f"Surface area/ft: {sfa_per_ft:.2f} ft^2/ft")
    print(f"Length used: {length_ft:.2f} ft")
    print(f"Total area (A): {area_ft2:.2f} ft^2")

    print(f"Seal class: {seal_class}")
    print(f"Leakage class constant (CL): {leakage_class:.2f}")
    print(f"Allowed leakage: {leakage_cfm:.2f} CFM")


if __name__ == "__main__":
    main()

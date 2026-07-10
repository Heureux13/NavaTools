# -*- coding: utf-8 -*-
"""SMACNA duct leakage calculation helpers."""

SMACNA_REFERENCE_PRESSURE_IN_WC = 1.0
SMACNA_PRESSURE_EXPONENT = 0.65
SMACNA_SURFACE_AREA_BASE_FT2 = 100.0


def allowable_leakage_factor(leakage_class, test_pressure_in_wg):
    """Return allowable leakage factor in CFM per ft^2."""
    if leakage_class < 0:
        raise ValueError('Leakage class must be non-negative.')
    if test_pressure_in_wg < 0:
        raise ValueError('Test pressure must be non-negative.')

    return (
        leakage_class
        * (test_pressure_in_wg / SMACNA_REFERENCE_PRESSURE_IN_WC) ** SMACNA_PRESSURE_EXPONENT
        / SMACNA_SURFACE_AREA_BASE_FT2
    )


def max_allowable_leakage(leakage_class, test_pressure_in_wg, surface_area_ft2):
    """Return maximum allowable leakage in CFM."""
    if surface_area_ft2 <= 0:
        raise ValueError('Surface area must be greater than zero.')

    return allowable_leakage_factor(leakage_class, test_pressure_in_wg) * surface_area_ft2


def measured_leakage_factor(measured_leakage_cfm, surface_area_ft2):
    """Return measured leakage factor in CFM per ft^2."""
    if measured_leakage_cfm < 0:
        raise ValueError('Measured leakage must be non-negative.')
    if surface_area_ft2 <= 0:
        raise ValueError('Surface area must be greater than zero.')

    return measured_leakage_cfm / surface_area_ft2


def leakage_passes(leakage_class, test_pressure_in_wg, surface_area_ft2, measured_leakage_cfm):
    """Return True when measured leakage is within the allowable limit."""
    return measured_leakage_cfm <= max_allowable_leakage(
        leakage_class,
        test_pressure_in_wg,
        surface_area_ft2,
    )


def leakage_summary(leakage_class, test_pressure_in_wg, surface_area_ft2, measured_leakage_cfm):
    """Return a calculation summary for the duct leakage test."""
    allowable_f = allowable_leakage_factor(leakage_class, test_pressure_in_wg)
    allowable_leakage = max_allowable_leakage(
        leakage_class,
        test_pressure_in_wg,
        surface_area_ft2,
    )
    measured_f = measured_leakage_factor(
        measured_leakage_cfm, surface_area_ft2)

    return {
        'allowable_f': allowable_f,
        'allowable_leakage': allowable_leakage,
        'measured_f': measured_f,
        'measured_leakage': measured_leakage_cfm,
        'passes': measured_leakage_cfm <= allowable_leakage,
    }

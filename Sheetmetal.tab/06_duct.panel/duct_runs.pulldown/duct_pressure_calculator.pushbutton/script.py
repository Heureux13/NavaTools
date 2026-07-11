# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import forms, script

from constants.delete import leakage_summary


__title__ = 'Duct Pressure'
__doc__ = """
Calculate SMACNA allowable duct leakage values from leakage class, test pressure,
surface area, and measured leakage.
"""


output = script.get_output()


def prompt_float(prompt, title, default):
    value = forms.ask_for_string(
        prompt=prompt,
        title=title,
        default=default,
    )
    if value is None:
        script.exit()

    text = value.strip()
    if not text:
        forms.alert('A numeric value is required.', title=title)
        script.exit()

    try:
        return float(text)
    except ValueError:
        forms.alert('Enter a valid number.', title=title)
        script.exit()


leakage_class = prompt_float(
    'Enter the SMACNA leakage class (for example 2, 3, 6, or 12).',
    'Leakage Class',
    '2',
)
test_pressure = prompt_float(
    'Enter the duct test pressure in inches of water column.',
    'Test Pressure (in. WC)',
    '2.0',
)
surface_area = prompt_float(
    'Enter the duct surface area in square feet.',
    'Surface Area (ft^2)',
    '47.12',
)
measured_leakage = prompt_float(
    'Enter the measured leakage in CFM.',
    'Measured Leakage (CFM)',
    '0.0',
)

try:
    results = leakage_summary(
        leakage_class,
        test_pressure,
        surface_area,
        measured_leakage,
    )
except ValueError as exc:
    forms.alert(str(exc), title='Duct Pressure')
    script.exit()

status_text = 'PASS' if results['passes'] else 'FAIL'

output.print_md('# Duct Pressure Calculator')
output.print_md(
    '- Leakage Class: `{:.3f}`  \n'
    '- Test Pressure: `{:.3f}` in. WC  \n'
    '- Surface Area: `{:.3f}` ft^2  \n'
    '- Measured Leakage: `{:.3f}` CFM'.format(
        leakage_class,
        test_pressure,
        surface_area,
        measured_leakage,
    )
)
output.print_md('---')
output.print_md(
    '### Allowable F\n'
    '`{:.3f} * ({:.3f} / 1.0)^0.65 / 100 = {:.6f} CFM/ft^2`'.format(
        leakage_class,
        test_pressure,
        results['allowable_f'],
    )
)
output.print_md(
    '### Max Allowable Leakage\n'
    '`{:.6f} * {:.3f} = {:.3f} CFM`'.format(
        results['allowable_f'],
        surface_area,
        results['allowable_leakage'],
    )
)
output.print_md(
    '### Measured F\n'
    '`{:.3f} / {:.3f} = {:.6f} CFM/ft^2`'.format(
        measured_leakage,
        surface_area,
        results['measured_f'],
    )
)
output.print_md(
    '## Result: `{}`  \n'
    'Measured leakage `{:.3f}` CFM {} allowable leakage `{:.3f}` CFM.'.format(
        status_text,
        measured_leakage,
        '<=' if results['passes'] else '>',
        results['allowable_leakage'],
    )
)

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false, reportCallIssue=false
# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

import os
import re

from Autodesk.Revit.DB import (
    Family,
    FilteredElementCollector,
)
from pyrevit import revit, script
from System.Windows.Forms import DialogResult, FolderBrowserDialog

# Button info
# ======================================================================
__title__ = 'Import Tag Families'
__doc__ = '''
Import all annotation/tag family
'''

# Variables
# ======================================================================

output = script.get_output()


def _find_families_dir(start_dir):
    """Search upward from this script for the repo's lib/families folder."""
    search_dir = start_dir
    while True:
        candidate = os.path.join(search_dir, 'lib', 'families')
        if os.path.isdir(candidate):
            return candidate

        parent = os.path.dirname(search_dir)
        if parent == search_dir:
            return None
        search_dir = parent


def _pick_families_dir():
    dialog = FolderBrowserDialog()
    dialog.Description = 'Select the folder containing annotation family files (.rfa).'
    dialog.ShowNewFolderButton = False

    if dialog.ShowDialog() == DialogResult.OK and dialog.SelectedPath:
        return dialog.SelectedPath

    return None


def _get_family_files(families_dir):
    family_files = []
    backup_pattern = re.compile(r'\d{4}\.rfa$', re.IGNORECASE)

    for root, _, files in os.walk(families_dir):
        for filename in files:
            if not filename.lower().endswith('.rfa'):
                continue
            if backup_pattern.search(filename):
                continue
            family_files.append(os.path.join(root, filename))

    family_files.sort(key=lambda p: os.path.basename(p).lower())
    return family_files


def _print_list(title, values):
    if not values:
        return
    output.print_md('### {}'.format(title))
    for value in values:
        output.print_md('- {}'.format(value))


def main():
    families_dir = _find_families_dir(os.path.dirname(__file__))
    if not families_dir:
        output.print_md(
            'Could not auto-locate lib/families. Pick a folder manually.')
        families_dir = _pick_families_dir()

    if not families_dir or not os.path.isdir(families_dir):
        output.print_md(
            'No valid families folder selected. Nothing was imported.')
        script.exit()

    family_files = _get_family_files(families_dir)
    if not family_files:
        output.print_md('No .rfa files found in: {}'.format(families_dir))
        script.exit()

    loaded = []
    skipped = []
    failed = []

    existing_family_names = {
        f.Name for f in FilteredElementCollector(revit.doc).OfClass(Family).ToElements()
    }

    try:
        with revit.Transaction('Import Tag Families'):
            for family_path in family_files:
                family_name = os.path.basename(family_path)
                family_name_no_ext = os.path.splitext(family_name)[0]

                if family_name_no_ext in existing_family_names:
                    skipped.append(family_name)
                    continue

                try:
                    loaded_ok = revit.doc.LoadFamily(family_path)
                    if loaded_ok:
                        loaded.append(family_name)
                        existing_family_names.add(family_name_no_ext)
                    else:
                        skipped.append(family_name)
                except Exception as load_err:
                    failed.append('{} ({})'.format(family_name, load_err))
    except Exception:
        output.print_md('Import transaction failed and was rolled back.')
        raise

    output.print_md('## Import Tag Families')
    output.print_md('- Source folder: {}'.format(families_dir))
    output.print_md('- Files found: {}'.format(len(family_files)))
    output.print_md('- Loaded (new): {}'.format(len(loaded)))
    output.print_md(
        '- Skipped (already in project or unchanged): {}'.format(len(skipped)))
    output.print_md('- Failed: {}'.format(len(failed)))

    _print_list('Loaded', loaded)
    _print_list('Skipped', skipped)
    _print_list('Failed', failed)


main()

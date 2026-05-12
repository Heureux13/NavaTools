# pyright: reportMissingImports=false, reportAttributeAccessIssue=false
# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

import os
import sys

from pyrevit import revit, script

# Button info
# ======================================================================
__title__ = 'Testing Populate Schedule Rooms'
__doc__ = '''
Import Excel or CSV data and populate an existing editable Revit schedule.
'''


def _find_lib_path(start_dir):
    search_dir = start_dir
    while True:
        candidate = os.path.join(search_dir, 'lib')
        if os.path.isdir(candidate):
            return candidate

        parent = os.path.dirname(search_dir)
        if parent == search_dir:
            return None
        search_dir = parent


def main():
    lib_path = _find_lib_path(os.path.dirname(__file__))
    if not lib_path:
        script.get_output().print_md('Could not locate lib folder from script path.')
        script.exit()
    assert lib_path is not None

    if lib_path not in sys.path:
        sys.path.insert(0, lib_path)

    from schedules.revit_schedules import RevitSchedules

    scheduler = RevitSchedules(
        doc=revit.doc,
        active_view=revit.active_view,
        output_obj=script.get_output(),
    )
    scheduler.run_populate_from_file_dialog()


main()

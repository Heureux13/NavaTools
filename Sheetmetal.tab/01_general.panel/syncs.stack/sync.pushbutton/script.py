# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from Autodesk.Revit.DB import SynchronizeWithCentralOptions, TransactWithCentralOptions
from pyrevit import revit, output, script

# Button display information
# =================================================
__title__ = "Sync"
__doc__ = """
Asks if you want to sync and save every hour
"""

# Variables
# ======================================================================================
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
doc = revit.doc
view = revit.active_view
output = script.get_output()

sync_options = SynchronizeWithCentralOptions()
transact_options = TransactWithCentralOptions()

try:
    doc.SynchronizeWithCentral(transact_options, sync_options)
    # output.print_md("# Sync & Save complete")
    # output.print_md("### I am proud of you")
except Exception as e:
    output.print_md("Sync & Save failed: {}".format(e))

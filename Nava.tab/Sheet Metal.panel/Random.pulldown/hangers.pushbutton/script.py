# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, script
from Autodesk.Revit.DB import *
from System.Collections.Generic import List

# Button info
# ===================================================
__title__ = "Select Unhosted Hangers"
__doc__ = """
Select all hangers that have no host.
"""

# Variables
# ==================================================
app = __revit__.Application             # type: Application
uidoc = __revit__.ActiveUIDocument      # type: UIDocument
doc = revit.doc                         # type: Document
output = script.get_output()

place holder

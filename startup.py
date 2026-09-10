# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

"""Runs automatically  at Revit startup to keep this extension up to date"""

import subprocess

NavaTools = 'MyTools'

try:
    subprocess.Popen(
        ['pyrevit', 'extension', 'update', NavaTools],
        creationFlags=subprocess.CREATE_NO_WINDOW
    )

except Exception:
    pass
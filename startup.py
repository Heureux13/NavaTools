# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

"""Runs automatically  at Revit startup to keep this extension up to date"""
import os
import subprocess
from pyrevit import script

logger = script.get_logger()
EXT_DIR = os.path.dirname(__file__)

try:
    p = subprocess.Popen(
        ['git', '-C', EXT_DIR, 'pull', '--ff-only'],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE
    )
    out, err = p.communicate()

    if p.returncode != 0:
        err_text = (err or b'').decode('utf-8', errors='ignore').strip()
        if not err_text:
            err_text = (out or b'').decode('utf-8', errors='ignore').strip()
        logger.warning('NavaTools auto-update failed (code %s): %s', p.returncode, err_text)

except Exception as ex:
    logger.warning('NavaTools auto-update exception: %s', ex)
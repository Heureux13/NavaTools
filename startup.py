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

    if out:
        logger.info(out.decode('utf-8', errors='ignore').strip())
    if err:
        logger.warning(err.decode('utf-8', errors='ignore').strip())
    if p.returncode != 0:
        logger.warning('Auto-update failed (exit %s).', p.returncode)
except Exception as ex:
    logger.warning('Auto-update error: %s', ex)
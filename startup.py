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

    out_text = out.decode('utf-8', errors='ignore').strip()
    err_text = err.decode('utf-8', errors='ignore').strip()

    if out_text:
        logger.info(out_text)
    if err_text:
        logger.warning(err_text)

    if p.returncode != 0:
        logger.warning('git pull failed with code %s', p.returncode)

except Exception as ex:
    logger.warning('startup update exception: %s', ex)

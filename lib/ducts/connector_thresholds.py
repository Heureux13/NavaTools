# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

from enum import Enum

# Define Constants
CONNECTOR_THRESHOLDS = {
    ("Straight", "TDC"): 56.00,
    ("Straight", "TDF"): 56.00,
    ("Straight", "Standing S&D"): 59.00,
    ("Straight", "Slip & Drive"): 59.00,
    ("Straight", "S&D"): 59.00,
    ("Tube", "AccuFlange"): 120.00,
    ("Tube", "GRC_Swage-Female"): 120.00,
    ("Spiral Duct", "Raw"): 120.00,
    ("Spiral Pipe", "Raw"): 120.00,
}

DEFAULT_SHORT_THRESHOLD_IN = 56.00


# Joint Size Class
# ====================================================
class JointSize(Enum):
    SHORT = "short"
    FULL = "full"
    LONG = "long"
    INVALID = "invalid"

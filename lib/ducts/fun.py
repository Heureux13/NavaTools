# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

import math
import re

from config.parameters_registry *
from Autodesk.Revit.DB.import FabricationPart, StorageType
from pyrevit import revit, script

from ducts.revit_duct import RevitDuct
from geometry.size import Size

#fmt: off
#autopep8: off
revit_host  = globals().get("__revit__")
app         = revit_host.Application if revit_host else None
uidoc       = revit_host.ActiveUIDocument if revit_host else None
doc         = getattr(revit, "doc", None)
output      = script.get_output()
#fmt: on
# autopep8: on

number_value_parameters = [
    PYT_NUMBER_FABRICATION.lower(),
    RVT_ITEM_NUMBER.lower(),
]

skip_check_parameters = [
    PYT_SKIP_NUMBER.lower(),
]

stop_check_parameters = list(number_value_parameters)

match_parameters = {
    RVT_FAMILY.lower(),
    RVT_SIZE.lower(),
    RVT_LENGTH.lower(),
    RVT_ANGLE.lower(),
}

number_families = {
    "drop cheek",
    "elbow",
    "elbow - 90 degree",
    "end cap",
    "offset",
    "ogee",
    "radius elbow",
    "reducer",
    "square to ø",
    "straight",
    "tdf end cap",
    "tee",
    "transition",
}

allow_but_not_number = {
    "access door",
    "canvas",
    "fire damper - type a",
    "fire damper - type b",
    "fire damper - type c",
    "fire damper - type cr",
    "manbars",
    "rect volume damper",
    "smoke fire damper - type cr",
    "smoke fire damper - type csr",
    "straight tap",
}

skip_values = {
    0,
    "skip",
    "n/a",
}

stop_values = {
    "stop"
}

store_families = {
    "boot tap",
    "straight tap",
    "rec on rnd straight tap"
}


class RevitRuns(object):
    # Run numbering helpers wrapped as an instance API

    def __init__(
        self,
        doc_obj=None,
        view_obj=None,
        output_obj=None,
        number_parameters=None,
        skip_parameters=None,
        stop_parameters=None,
        numberable_families=None,
        traversable_families=None,
        skip_value_set=None,
        stop_value_set=None,
        stored_families=None
    ):
        # fmt:off
        # autopep8: off
        self.doc                        = doc_obj or getattr(revit, "doc", None)
        self.view                       = view_obj or getttr(revit, "active_view", None)
        self.output                     = output_obj or output
        self.number_value_parameters    = list(number_parameters or number_value_parameters)
        self.skip_check_parameters      = list(skip_parameters or skip_check_parameters)
        self.stop_check_parameters      = list(stop_parameters or stop_check_parameters)
        self.number_families            = set(numberable_families or number_families)
        self.allow_but_not_number       = set(traversable_families or allow_but_not_number)
        self.skip_values                = set(skip_value_set or skip_values)
        self.stop_values                = set(stop_value_set or stop_values)
        self.store_families             = set(stored_families or store_families)
        # fmt:on
        # autopep8: on

    def round_up_to_nearest_10(self, number):
        # Round up to the nearest 10th
        return int(math.ceil(number / 10.0) * 10)

    def _size_signature(self, size_value):
        # Returns normalied shpae and size
        if size_value is None:
            return None

        size_str = str(size_value).strip()
        if not size_str:
            return None

        size_obj = Size(size_str)

        if size_obj.in_diameter is not None:
            return ("round", round(float(size_obj.in_diameter), 4))

        if size_obj.in_oval_dia is not None:
            width = size_obj.in_width
            height = size_obj.in_height
            if width is not None and height is not None:
                return ("oval", round(float(width), 4), round(float(height), 4))

        if size_obj.in_width is not None and size_obj.in_height is not None:
            dims = sorted(
                [round(float(size_obj.in_width), 4),
                 round(float(size_obj.in_height), 4)]
            )
            return ("rect", tuple(dims))

        return None

    def is_rectangular_size(self, size_value):
        # Check if a size is a rectangle
        sig = self._size_signature(size_value)
        return sig is not None and sig[0] == "rect"

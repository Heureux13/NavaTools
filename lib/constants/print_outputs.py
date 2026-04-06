# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""


def print_disclaimer(output):
    """Print standard help message about missing parameters."""
    output.print_md("---")
    output.print_md(
        "*If info is missing, Naviate parameters must be imported (may require enabling per family/element)*")
    output.print_md(
        "**Connectors**: *Turn all on*")
    output.print_md(
        "**Dimensions**: *Turn all on*")
    output.print_md(
        "**Fab Properties:** *Size, Weight, Diameter, Family, Left Extension, Right Extension, Elevations, \
            Centerline Length, Depth, Inner Radius, and SheetMetalArea*")

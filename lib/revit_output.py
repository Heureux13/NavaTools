# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""


def print_parameter_help(output):
    """Print standard help message about missing parameters."""
    output.print_md("---")
    output.print_md("If info is missing, Naviate parameters must be imported")
    output.print_md(
        "Connectors: Turn all on (may require enabling per family/element)")
    output.print_md(
        "Dimensions: Turn all on (may require enabling per family/element)")
    output.print_md(
        "Fab Properties: Size, Weight, Diameter, Family, Left Extension, Right Extension, and Elevations")


def print_selection_summary(output, element_ids, label="Total elements"):
    """Print summary of selected elements with linkified IDs."""
    output.print_md("# {}: {}, {}".format(
        label, len(element_ids), output.linkify(element_ids)
    ))

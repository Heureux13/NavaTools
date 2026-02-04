# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from revit_element import RevitElement
from revit_duct import RevitDuct
from revit_output import print_disclaimer
from pyrevit import revit, script
from Autodesk.Revit.DB import *

# Button info
# ===================================================
__title__ = "Hanger weight on all runs"
__doc__ = """
Assigns weight on all hangers and runs.
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view
output = script.get_output()

hanger_parameters = [
    '_hang_weight_supporting',
]

duct_parameters = [
    '_duct_weight_run',
]

# Main Code
# =================================================

# Regenerate document to clear cache and refresh parameters
with revit.Transaction("Regenerate Document"):
    doc.Regenerate()


def get_host_duct_from_hanger(hanger, ducts_cache):
    """Resolve the duct element for a hanger.

    Tries Primary Element parameter first, then falls back to bbox intersection
    against all ducts in the view.
    """
    ref_param = hanger.LookupParameter("Primary Element")
    if ref_param:
        ref_id_str = ref_param.AsString()
        try:
            ref_id = int(ref_id_str)
            elem = doc.GetElement(ElementId(ref_id))
            if elem:
                return RevitDuct(doc, view, elem)
        except Exception:
            pass

    # Fallback: bbox intersect to find a duct
    bbox = hanger.get_BoundingBox(None)
    if bbox:
        outline = Outline(bbox.Min, bbox.Max)
        bbox_filter = BoundingBoxIntersectsFilter(outline)
        # Use cached ducts to avoid repeated collectors
        for d in ducts_cache:
            d_bbox = d.element.get_BoundingBox(None)
            if not d_bbox:
                continue
            d_outline = Outline(d_bbox.Min, d_bbox.Max)
            if outline.Intersects(d_outline, 0):
                return d
    return None


# Collect all ducts and hangers in view
all_ducts = RevitDuct.all(doc, view)
all_hangers = FilteredElementCollector(doc, view.Id)\
    .OfCategory(BuiltInCategory.OST_FabricationHangers)\
    .WhereElementIsNotElementType()\
    .ToElements()

pending_hanger_ids = [h.Id for h in all_hangers]
processed_duct_ids = set()
processed_hanger_ids = set()
run_number = 0

while pending_hanger_ids:
    hanger_id = pending_hanger_ids.pop(0)
    hanger = doc.GetElement(hanger_id)
    if not hanger:
        continue

    host_duct = get_host_duct_from_hanger(hanger, all_ducts)
    if not host_duct:
        # Could not resolve host; skip
        continue

    if host_duct.id in processed_duct_ids:
        continue

    run_number += 1

    # Build run from the host duct
    run = RevitDuct.create_duct_run(host_duct, doc, view)
    RevitElement.select_many(uidoc, run)
    run_total_length = sum(d.length or 0 for d in run)
    run_total_weight = sum(d.weight or 0 for d in run)

    # Collect hangers intersecting any duct in this run
    run_hanger_ids = set()
    for duct in run:
        bbox = duct.element.get_BoundingBox(None)
        if not bbox:
            continue
        outline = Outline(bbox.Min, bbox.Max)
        bbox_filter = BoundingBoxIntersectsFilter(outline)
        intersecting = FilteredElementCollector(doc)\
            .OfCategory(BuiltInCategory.OST_FabricationHangers)\
            .WherePasses(bbox_filter)\
            .WhereElementIsNotElementType()\
            .ToElements()
        for h in intersecting:
            # Skip hangers already assigned to previous runs
            if h.Id not in processed_hanger_ids:
                run_hanger_ids.add(h.Id)

    # Convert IDs back to elements
    run_hangers = [doc.GetElement(hid) for hid in run_hanger_ids]

    # Remove processed hangers from pending list and mark as processed
    pending_hanger_ids = [
        hid for hid in pending_hanger_ids if hid not in run_hanger_ids]
    processed_hanger_ids.update(run_hanger_ids)

    if run_hangers:
        weight_per_hanger = run_total_weight / \
            len(run_hangers) if run_hangers else 0
        hanger_ids = [h.Id for h in run_hangers]
        RevitElement.select_many(uidoc, run_hangers)

        with revit.Transaction("Set Hanger Mark - Run {}".format(run_number)):
            for h in run_hangers:
                set_parameter = None
                for parameter_name in hanger_parameters:
                    p = h.LookupParameter(parameter_name)
                    if not p or p.IsReadOnly:
                        continue
                    set_parameter = p
                    break

                if set_parameter:
                    try:
                        if set_parameter.StorageType == StorageType.Double:
                            set_parameter.Set(weight_per_hanger)
                        elif set_parameter.StorageType == StorageType.String:
                            set_parameter.Set(str(round(weight_per_hanger, 2)))
                    except Exception:
                        pass

            # Set run weight on each duct in the run
            for d in run:
                set_parameter = None
                for parameter_name in duct_parameters:
                    p = d.element.LookupParameter(parameter_name)
                    if not p or p.IsReadOnly:
                        continue
                    set_parameter = p
                    break

                if set_parameter:
                    try:
                        if set_parameter.StorageType == StorageType.Double:
                            set_parameter.Set(round(run_total_weight, 2))
                        elif set_parameter.StorageType == StorageType.String:
                            set_parameter.Set(str(round(run_total_weight, 2)))
                    except Exception:
                        pass

        output.print_md("---")
        output.print_md("# Run {}".format(run_number))
        output.print_md("Hangers: {} | Support each: {:0.2f} lbs".format(
            output.linkify(hanger_ids), weight_per_hanger))

    else:
        output.print_md("---")
        output.print_md("# Run {}".format(run_number))
        output.print_md("Hangers: None found")

        with revit.Transaction("Set Duct Weight - Run {}".format(run_number)):
            for d in run:
                set_parameter = None
                for parameter_name in duct_parameters:
                    p = d.element.LookupParameter(parameter_name)
                    if not p or p.IsReadOnly:
                        continue
                    set_parameter = p
                    break

                if set_parameter:
                    try:
                        if set_parameter.StorageType == StorageType.Double:
                            set_parameter.Set(round(run_total_weight, 2))
                        elif set_parameter.StorageType == StorageType.String:
                            set_parameter.Set(str(round(run_total_weight, 2)))
                    except Exception:
                        pass

    # Totals
    duct_element_ids = [d.element.Id for d in run]
    total_length_ft = run_total_length / 12.0 if run_total_length else 0.0
    lbs_per_ft = (run_total_weight /
                  total_length_ft) if total_length_ft else 0.0
    length_str = "{:06.2f}".format(float(total_length_ft))
    weight_str = "{:6.2f}".format(float(run_total_weight))
    lbs_ft_str = "{:6.2f}".format(float(lbs_per_ft))
    duct_ids_str = str(output.linkify(duct_element_ids))
    output.print_md("Ducts: {} | Qty: {} | Length: {} ft | Weight: {} lbs | lbs/ft: {}".format(
        duct_ids_str,
        len(duct_element_ids),
        length_str,
        weight_str,
        lbs_ft_str
    ))

    # Mark ducts processed
    for d in run:
        processed_duct_ids.add(d.id)

# Process remaining ducts without hangers
output.print_md("---")
output.print_md("## Processing unassigned ducts")

remaining_ducts = [d for d in all_ducts if d.id not in processed_duct_ids]

while remaining_ducts:
    host_duct = remaining_ducts.pop(0)

    # Build run from this unassigned duct
    run = RevitDuct.create_duct_run(host_duct, doc, view)
    run_number += 1

    run_total_weight = sum(d.weight or 0 for d in run)
    run_total_length = sum(d.length or 0 for d in run)

    # Remove these ducts from remaining_ducts
    remaining_ducts = [
        d for d in remaining_ducts if d.id not in set(rd.id for rd in run)]

    if run_total_weight > 0:
        with revit.Transaction("Set Duct Weight - Run {} (no hangers)".format(run_number)):
            for d in run:
                set_parameter = None
                for parameter_name in duct_parameters:
                    p = d.element.LookupParameter(parameter_name)
                    if not p or p.IsReadOnly:
                        continue
                    set_parameter = p
                    break

                if set_parameter:
                    try:
                        if set_parameter.StorageType == StorageType.Double:
                            set_parameter.Set(round(run_total_weight, 2))
                        elif set_parameter.StorageType == StorageType.String:
                            set_parameter.Set(str(round(run_total_weight, 2)))
                    except Exception:
                        pass

    # Output
    duct_element_ids = [d.element.Id for d in run]
    total_length_ft = run_total_length / 12.0 if run_total_length else 0.0
    lbs_per_ft = (run_total_weight /
                  total_length_ft) if total_length_ft else 0.0
    length_str = "{:06.2f}".format(float(total_length_ft))
    weight_str = "{:6.2f}".format(float(run_total_weight))
    lbs_ft_str = "{:6.2f}".format(float(lbs_per_ft))
    duct_ids_str = str(output.linkify(duct_element_ids))

    output.print_md("---")
    output.print_md("# Run {}".format(run_number))
    output.print_md("Ducts: {} | Qty: {} | Length: {} ft | Weight: {} lbs | lbs/ft: {}".format(
        duct_ids_str,
        len(duct_element_ids),
        length_str,
        weight_str,
        lbs_ft_str
    ))

output.print_md("---")
output.print_md(
    "## Processing complete - {} runs processed".format(run_number))
print_disclaimer(output)

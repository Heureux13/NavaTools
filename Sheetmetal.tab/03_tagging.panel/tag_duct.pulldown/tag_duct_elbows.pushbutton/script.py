# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from System.Collections.Generic import List
from constants.print_outputs import print_disclaimer
from tagging.revit_tagging import RevitTagging
from revit.revit_element import RevitElement
from ducts.revit_duct import RevitDuct
from tagging.revit_tagging_fittings import Fittings
from pyrevit import DB, revit, script
from Autodesk.Revit.DB import ElementId, Transaction

# Button info
# ==================================================
__title__ = "Tag Elbows"
__doc__ = """
Tag all fitting with assosiated label
"""

# Revit context
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
output = script.get_output()
view = revit.active_view
tagger = RevitTagging(doc=doc, view=view)

fittings = Fittings(doc=doc, view=view, tagger=tagger)


def _fmt_length(value):
    """Format length safely for report output."""
    if isinstance(value, (int, float)):
        return "{:06.2f}".format(float(value))
    try:
        return "{:06.2f}".format(float(str(value).strip()))
    except Exception:
        return str(value)


# ======================================================================
# MAIN
# ======================================================================

# Change this tuple to target other groups, e.g. ("reducers",) or ("all",).
TARGET_FAMILY_GROUPS = ("elbows",)
duct_families = fittings.get_duct_families_for_groups(TARGET_FAMILY_GROUPS)

# Collect and filter ducts in the active view.
ducts = RevitDuct.all(doc, view)
if not ducts:
    output.print_md("# No ducts found in the current view")
    import sys
    sys.exit()

if fittings.missing_tag_labels:
    output.print_md("## Missing tag label(s); skipped where unavailable: {}".format(
        ", ".join(sorted(fittings.missing_tag_labels))))

dic_ducts = [d for d in ducts if fittings._norm(
    d.family) in duct_families]

# Tag in a single transaction.
t = Transaction(doc, "General Tagging")
t.Start()
try:
    needs_tagging = []
    already_tagged = []
    skipped_placement = []
    skipped_by_param = []
    skipped_no_tag_config = []
    auto_removed = []

    for d in dic_ducts:
        key = fittings._norm(d.family)
        tag_configs = duct_families.get(key)
        if not tag_configs:
            skipped_no_tag_config.append(d)
            continue

        fittings.update_write_parameter_from_hierarchy(d.element)

        removed_count = fittings.delete_skipped_tags_for_element(d, tag_configs)
        if removed_count:
            auto_removed.append((d, removed_count))

        if fittings.should_skip_by_param(d):
            skipped_by_param.append(d)
            continue

        existing_tag_fams = tagger.get_existing_tag_families(d.element)
        requested_tag_fams = {
            (tag.Family.Name or '').strip().lower()
            for tag, _ in tag_configs
            if tag is not None and getattr(tag, 'Family', None) is not None
        }
        has_matching_existing_tag = bool(existing_tag_fams & requested_tag_fams)

        tagged_this_element = False
        placement_failed_reason = None
        for tag, loc_param in tag_configs:
            if tag is None or fittings.should_skip_tag(d, tag):
                continue
            fam_name = (
                tag.Family.Name if tag and tag.Family else "").strip().lower()
            if not fam_name or fam_name in existing_tag_fams:
                continue

            # Tag placement: FabricationPart uses face reference; others use location.
            placed_tag = None
            if isinstance(d.element, DB.FabricationPart):
                face_ref, face_pt = tagger.get_face_facing_view(
                    d.element, prefer_point=None)
                if face_ref is not None and face_pt is not None:
                    placed_tag = tagger.place_tag(face_ref, tag, face_pt)
                else:
                    bbox = d.element.get_BoundingBox(view)
                    if bbox is None:
                        continue
                    placed_tag = tagger.place_tag(
                        d.element, tag, (bbox.Min + bbox.Max) / 2.0)
            else:
                loc = getattr(d.element, "Location", None)
                if not loc:
                    bbox = d.element.get_BoundingBox(view)
                    if bbox is None:
                        continue
                    placed_tag = tagger.place_tag(
                        d.element, tag, (bbox.Min + bbox.Max) / 2.0)
                elif hasattr(loc, "Point") and loc.Point is not None:
                    placed_tag = tagger.place_tag(d.element, tag, loc.Point)
                elif hasattr(loc, "Curve") and loc.Curve is not None:
                    placed_tag = tagger.place_tag(
                        d.element, tag, loc.Curve.Evaluate(loc_param, True))
                else:
                    continue

            if placed_tag is None:
                placement_failed_reason = (
                    tagger.last_place_tag_failure
                    or "No compatible tag could be placed"
                )
                continue

            existing_tag_fams.add(fam_name)
            tagged_this_element = True

        if tagged_this_element:
            needs_tagging.append(d)
        elif has_matching_existing_tag:
            already_tagged.append(d)
        else:
            skipped_placement.append(
                (d, placement_failed_reason or "No matching existing tag and no new tag was placed"))

    # Update selection.
    if needs_tagging:
        RevitElement.select_many(uidoc, needs_tagging)
    else:
        uidoc.Selection.SetElementIds(List[ElementId]())

    t.Commit()

except Exception as e:
    output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise

# Report.
output.print_md("# Tagged {} new fitting(s) | {} total fittings in view".format(
    len(needs_tagging), len(dic_ducts)))
output.print_md("---")

if needs_tagging:
    output.print_md("## Newly Tagged")
    for i, d in enumerate(needs_tagging, start=1):
        output.print_md("### No.{} | ID: {} | Fam: {} | Size: {} | Le: {} | Ex: {}".format(
            i, output.linkify(d.element.Id), d.family, d.size, _fmt_length(d.length), d.extension_bottom))
    output.print_md("---")

if already_tagged:
    output.print_md("## Already Tagged")
    for i, d in enumerate(already_tagged, start=1):
        output.print_md("### {} | Size: {} | Family: {} | Length: {} | ID: {}".format(
            i, d.size, d.family, _fmt_length(d.length), output.linkify(d.element.Id)))
    output.print_md("---")

if skipped_placement:
    output.print_md("## Skipped – Placement Failed")
    for i, item in enumerate(skipped_placement, start=1):
        d, reason = item
        output.print_md("### {} | Size: {} | Family: {} | Length: {} | ID: {} | Reason: {}".format(
            i, d.size, d.family, _fmt_length(d.length), output.linkify(d.element.Id), reason))
    output.print_md("---")

if skipped_by_param:
    output.print_md("## Skipped by Parameter")
    for i, d in enumerate(skipped_by_param, start=1):
        output.print_md("### {} | Size: {} | Family: {} | Length: {} | ID: {}".format(
            i, d.size, d.family, _fmt_length(d.length), output.linkify(d.element.Id)))
    output.print_md("---")

if auto_removed:
    output.print_md("## Auto Removed Invalid Tags")
    for i, item in enumerate(auto_removed, start=1):
        d, removed_count = item
        output.print_md("### {} | Removed: {} | Size: {} | Family: {} | ID: {}".format(
            i, removed_count, d.size, d.family, output.linkify(d.element.Id)))
    output.print_md("---")

if skipped_no_tag_config:
    output.print_md("## Skipped – Tag Family Not Loaded")
    for i, d in enumerate(skipped_no_tag_config, start=1):
        output.print_md("### {} | Family: {} | Size: {} | ID: {}".format(
            i, d.family, d.size, output.linkify(d.element.Id)))
    output.print_md("---")

output.print_md("# Newly tagged: {}, {}".format(
    len(needs_tagging), output.linkify([d.element.Id for d in needs_tagging])))
output.print_md("# Already tagged: {}, {}".format(
    len(already_tagged), output.linkify([d.element.Id for d in already_tagged])))
output.print_md("# Skipped (placement failed): {}, {}".format(
    len(skipped_placement), output.linkify([d.element.Id for d, _ in skipped_placement])))
output.print_md("# Skipped by parameter: {}, {}".format(
    len(skipped_by_param), output.linkify([d.element.Id for d in skipped_by_param])))
output.print_md("# Auto removed invalid tags: {}, {}".format(
    len(auto_removed), output.linkify([d.element.Id for d, _ in auto_removed])))
output.print_md("# Skipped (no tag family loaded): {}, {}".format(
    len(skipped_no_tag_config), output.linkify([d.element.Id for d in skipped_no_tag_config])))
output.print_md("# Total: {}, {}".format(
    len(dic_ducts), output.linkify([d.element.Id for d in dic_ducts])))

print_disclaimer(output)

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
from revit_output import print_disclaimer
from revit_tagging import RevitTagging
from revit_element import RevitElement
from revit_duct import RevitDuct
from revit_tagging_fittings import Fittings
from pyrevit import DB, revit, script
from Autodesk.Revit.DB import ElementId, Transaction

# Button info
# ==================================================
__title__ = "Tag Fittings"
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


# ======================================================================
# MAIN
# ======================================================================

duct_families = fittings.duct_families
RECT_DAMPER_SWITCH_FAMILIES = fittings.rect_damper_switch_families

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
    skipped_by_param = []

    for d in dic_ducts:
        key = fittings._norm(d.family)
        tag_configs = duct_families.get(key)
        if not tag_configs:
            continue

        if fittings.should_skip_by_param(d):
            skipped_by_param.append(d)
            continue

        existing_tag_fams = tagger.get_existing_tag_families(d.element)

        # Rect volume damper: choose MARK or TM tag and remove the losing one first.
        if key == fittings._norm('rect volume damper'):
            chosen_tag, chosen_family_name = fittings._rect_volume_damper_tag_choice(
                d)
            if chosen_tag is not None:
                fittings._delete_conflicting_tags_for_element(
                    d.element, chosen_family_name, RECT_DAMPER_SWITCH_FAMILIES)
                existing_tag_fams = tagger.get_existing_tag_families(d.element)
                tag_configs = [(chosen_tag, 0.5)]

        tagged_this_element = False
        for tag, loc_param in tag_configs:
            if tag is None or fittings.should_skip_tag(d, tag):
                continue
            fam_name = (
                tag.Family.Name if tag and tag.Family else "").strip().lower()
            if not fam_name or fam_name in existing_tag_fams:
                continue

            # Tag placement: FabricationPart uses face reference; others use location.
            if isinstance(d.element, DB.FabricationPart):
                face_ref, face_pt = tagger.get_face_facing_view(
                    d.element, prefer_point=None)
                if face_ref is not None and face_pt is not None:
                    tagger.place_tag(face_ref, tag, face_pt)
                else:
                    bbox = d.element.get_BoundingBox(view)
                    if bbox is None:
                        continue
                    tagger.place_tag(
                        d.element, tag, (bbox.Min + bbox.Max) / 2.0)
            else:
                loc = getattr(d.element, "Location", None)
                if not loc:
                    bbox = d.element.get_BoundingBox(view)
                    if bbox is None:
                        continue
                    tagger.place_tag(
                        d.element, tag, (bbox.Min + bbox.Max) / 2.0)
                elif hasattr(loc, "Point") and loc.Point is not None:
                    tagger.place_tag(d.element, tag, loc.Point)
                elif hasattr(loc, "Curve") and loc.Curve is not None:
                    tagger.place_tag(
                        d.element, tag, loc.Curve.Evaluate(loc_param, True))
                else:
                    continue

            existing_tag_fams.add(fam_name)
            tagged_this_element = True

        (needs_tagging if tagged_this_element else already_tagged).append(d)

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
        output.print_md("### No.{} | ID: {} | Fam: {} | Size: {} | Le: {:06.2f} | Ex: {}".format(
            i, output.linkify(d.element.Id), d.family, d.size, d.length, d.extension_bottom))
    output.print_md("---")

if already_tagged:
    output.print_md("## Already Tagged")
    for i, d in enumerate(already_tagged, start=1):
        output.print_md("### {} | Size: {} | Family: {} | Length: {:06.2f} | ID: {}".format(
            i, d.size, d.family, d.length, output.linkify(d.element.Id)))
    output.print_md("---")

if skipped_by_param:
    output.print_md("## Skipped by Parameter")
    for i, d in enumerate(skipped_by_param, start=1):
        output.print_md("### {} | Size: {} | Family: {} | Length: {:06.2f} | ID: {}".format(
            i, d.size, d.family, d.length, output.linkify(d.element.Id)))
    output.print_md("---")

output.print_md("# Newly tagged: {}, {}".format(
    len(needs_tagging), output.linkify([d.element.Id for d in needs_tagging])))
output.print_md("# Already tagged: {}, {}".format(
    len(already_tagged), output.linkify([d.element.Id for d in already_tagged])))
output.print_md("# Skipped by parameter: {}, {}".format(
    len(skipped_by_param), output.linkify([d.element.Id for d in skipped_by_param])))
output.print_md("# Total: {}, {}".format(
    len(dic_ducts), output.linkify([d.element.Id for d in dic_ducts])))

print_disclaimer(output)

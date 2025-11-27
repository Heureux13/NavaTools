# -*- coding: utf-8 -*-
# ==================================================
# Tag full-size straight joints connected to non-straight fittings.
# ==================================================

from pyrevit import revit, script, forms
from revit_duct import RevitDuct, JointSize
from revit_tagging import RevitTagging
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, IndependentTag

__title__ = "DO NOT PRESS"
__doc__ = """
******************************************************************
Tags full-size straight joints that connect to non-straight ducts.
******************************************************************
"""

# Environment
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()

def main():
    # Collect all fabrication ducts visible in the view
    ducts = RevitDuct.all(doc, view)
    if not ducts:
        forms.alert("No ducts found in view.", exitscript=True)

    # Families considered "non-straight"
    straight_families = {"straight", "spiral tube", "round duct", "tube"}
    non_straight_ducts = [
        d for d in ducts
        if d.family and d.family.strip().lower() not in straight_families
    ]
    output.print_md("**Found {} non-straight ducts**".format(len(non_straight_ducts)))

    # Collect unique straight neighbors connected to any non-straight duct
    connected_straights = []
    seen_ids = set()
    for d in non_straight_ducts:
        for connector_index in (0, 1, 2):
            connected = d.get_connected_elements(connector_index)
            if not connected:
                continue
            for elem in connected:
                eid = elem.Id.IntegerValue
                if eid in seen_ids:
                    continue
                wrap = RevitDuct(doc, view, elem)
                fam = (wrap.family or "").strip().lower()
                if fam in straight_families:
                    connected_straights.append(elem)
                    seen_ids.add(eid)

    output.print_md("**Found {} unique connected straight ducts**".format(len(connected_straights)))

    # Filter to full-size joints only
    full_joint_straights = []
    for elem in connected_straights:
        wrap = RevitDuct(doc, view, elem)
        if wrap.joint_size == JointSize.FULL:
            full_joint_straights.append(elem)

    output.print_md("**Found {} full-size joints to tag**".format(len(full_joint_straights)))

    if not full_joint_straights:
        output.print_md("Nothing to tag. ✓")
        return

    tagger = RevitTagging(doc, view)

    # Build a set of (element_id, tag_family_name) for existing tags
    already_tagged = set()
    existing_tags = FilteredElementCollector(doc, view.Id).OfClass(IndependentTag).ToElements()
    for tag in existing_tags:
        try:
            refs = tag.GetTaggedReferences()
            if not refs:
                continue
            ref = refs[0]
            tagged_elem_id = ref.ElementId.IntegerValue
            fam_name = tag.GetType().FamilyName if tag.GetType() else ""
            already_tagged.add((tagged_elem_id, fam_name))
        except Exception:
            continue

    output.print_md("**Found {} existing tag references**".format(len(already_tagged)))

    # Tag names to apply
    tag_names = ["0_bod", "0_size"]
    tagged_count = 0
    skipped_count = 0

    with Transaction(doc, "Tag Full-Size Straights") as t:
        t.Start()
        for elem in full_joint_straights:
            eid = elem.Id.IntegerValue
            for tag_name in tag_names:
                try:
                    tag_symbol = tagger.get_label(tag_name)
                except Exception:
                    output.print_md("Missing tag family containing '{}'".format(tag_name))
                    continue

                if (eid, tag_symbol.FamilyName) in already_tagged:
                    skipped_count += 1
                    continue

                # Try face for better placement
                face_ref, face_pt = tagger.get_face_facing_view(elem)
                try:
                    if face_ref is not None and face_pt is not None:
                        tagger.place_tag(face_ref, tag_symbol, face_pt)
                    else:
                        # Fallback: bounding box center
                        bbox = elem.get_BoundingBox(view)
                        if bbox:
                            center = (bbox.Min + bbox.Max) / 2.0
                            tagger.place_tag(elem, tag_symbol, center)
                        else:
                            output.print_md("No placement geometry for element {}".format(eid))
                            continue
                    tagged_count += 1
                    already_tagged.add((eid, tag_symbol.FamilyName))
                except Exception as e:
                    output.print_md("Failed tagging {}: {}".format(eid, e))
        t.Commit()

    output.print_md("**Tagged {} elements, skipped {} already tagged**".format(tagged_count, skipped_count))
    output.print_md("✓ Complete")

if __name__ == "__main__":
    main()
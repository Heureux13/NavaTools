# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import Transaction
from pyrevit import script, DB, forms, revit
from revit.revit_tagging import RevitTagging
from config.tag_config import DEFAULT_TAG_SLOT_CANDIDATES, SLOT_LENGTH

# Button info
# ======================================================================
__title__ = 'Testting Revit Tagging Class.'
__author__ = ''
__doc__ = '''
Sandbox for testing Revit Tagging Class.
'''

# Variables
# ======================================================================
app    = __revit__.Application # type: Application
uidoc  = __revit__.ActiveUIDocument # type: UIDocument
doc    = revit.doc # type: Document
view   = doc.ActiveView
output = script.get_output()
tagger = RevitTagging(doc=doc, view=view)

selected = list(revit.get_selection().elements)
if not selected:
    output.print_md("Select an element first")
    raise SystemExit
element = selected[0]

tag_candidate = DEFAULT_TAG_SLOT_CANDIDATES.get(SLOT_LENGTH, [])
tag_symbol = None

for tag, fam, typ, famtyp in tagger._tag_data:
    #output.print_md("{} | {} | {}".format(fam, typ, tag))
    pass

# Test get_tag_symbol_id_from_family_and_type
for family_name, type_name in tag_candidate:
    try:
        sid = tagger.get_tag_symbol_id_from_family_and_type(family_name, type_name)
        (output.print_md(
            "family/type -> symbol id: {} / {} -> {}"
        .format(family_name, type_name, sid)))
    except LookupError:
        output.print_md("family/type not found: {} / {}".format(family_name, type_name))

# Test build_tag_symbol_id_map
slot_map = tagger.build_tag_symbol_id_map()
output.print_md("slot LENGTH ids: {}".format(slot_map.get(SLOT_LENGTH, [])))

# Test get_tag_symbol_id_from_element
element_ids = tagger.get_tag_symbol_id_from_element(element)
output.print_md("selected element tag type ids: {}".format(element_ids))

for family_name, type_name in tag_candidate:
    try:
        tag_symbol = tagger._tag_symbol(family_name, type_name)
        output.print_md(
            "family name: {} | Type name: {} | Tag symbol: {}".format(
                family_name,
                type_name,
                tag_symbol.Id))
        break
    except LookupError:
        continue

if tag_symbol is None:
    output.print_md("No tag symbol found")
    raise SystemExit
t = Transaction(doc, "Testing Revit Tagging Class")
t.Start()
try:
    if not tagger.already_tagged(element, tag_symbol):
        tag = tagger.create_tag(element, tag_symbol, x_loc=0.5,rotate=True)
        if tag:
            output.print_md("successfully created tag")
    else:
        output.print_md("already tagged")

    t.Commit()
except:
    t.RollBack()
    raise
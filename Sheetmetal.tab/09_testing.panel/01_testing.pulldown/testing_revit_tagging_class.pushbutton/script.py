# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import Transaction
from pyrevit import script, revit
from revit.revit_tagging_new import RevitTagging
from config.tag_config import DEFAULT_TAG_SLOT_CANDIDATES, SLOT_LENGTH

# Button info
# ======================================================================
__title__ = 'sandbox for new tagging'
__author__ = ''
__doc__ = """
Sandbox for testing Revit Tagging Class.
"""

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

# Test _tag_symbol
for family_name, type_name in tag_candidate:
    try:
        tag_symbol = tagger._tag_symbol(family_name, type_name)
        symbol_id  = tag_symbol.Id
        output.print_md(
            "family/type -> symbol id: {} / {} -> {}"
            .format(family_name,
                    type_name,
                    symbol_id))
    except LookupError:
        output.print_md("family/type not found: {} / {}".format(family_name, type_name))

# Test get_tag_symbols_from_element
symbols = tagger.get_tag_symbols_from_element(element)
for s in symbols:
    output.print_md("get_tag_symbols_from_element is working and returns: {}".format(s.Id))


# Test build_tag_symbol_id_map
slot_map = tagger.build_tag_symbol_id_map(slot_map=DEFAULT_TAG_SLOT_CANDIDATES)
output.print_md("build_tag_symbol_id_map is working slot LENGTH ids: {}"
                .format(slot_map.get(SLOT_LENGTH, [])))

# Mini test
mini_map = {SLOT_LENGTH: DEFAULT_TAG_SLOT_CANDIDATES.get(SLOT_LENGTH, [])}
slot_map = tagger.build_tag_symbol_id_map(slot_map=mini_map)
output.print_md("slot LENGTH ids: {}".format(slot_map.get(SLOT_LENGTH, [])))

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
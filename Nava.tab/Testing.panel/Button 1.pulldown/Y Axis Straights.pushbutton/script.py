"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""


from pyrevit import revit
from duct_tagger import DuctTagger   # import the class from your file

# Create an instance of the tagger for the current document and view
duct = DuctTagger(revit.doc, revit.active_view)

# Call the method you want
duct.tag_horizontal()
duct.tag_vertical()
# duct.tag_angled_up()
duct.tag_angled_down()

print("Franks Taggabalooze is complete, my leige.")

# script.py
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

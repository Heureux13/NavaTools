# -*- coding: utf-8 -*-
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""


__title__   = "Taggapaluza"
__doc__     = """
Version = 00.00.01
Date    = 2025-10-27
________________________________________________________________
Description:

This is the placeholder for a .pushbutton in a Stack/Pulldown
You can use it to start your pyRevit Add-In

________________________________________________________________
How-To:

1. [Hold ALT + CLICK] on the button to open its source folder.
You will be able to override this placeholder.

2. Automate Your Boring Work ;)

________________________________________________________________
TODO:
[FEATURE] - Describe Your ToDo Tasks Here
________________________________________________________________
Last Updates:
- [15.06.2024] v1.0 Change Description

________________________________________________________________
Author: Jose Nava
"""

# Imports
#==================================================
from Autodesk.Revit.DB import *

#.NET Imports
import clr
clr.AddReference('System')
from System.Collections.Generic import List


# Variables
#==================================================
app    = __revit__.Application
uidoc  = __revit__.ActiveUIDocument
doc    = __revit__.ActiveUIDocument.Document #type:Document


# Main Code
#==================================================




#🤖 Automate Your Boring Work Here
print("Hello there, Button 1! Looks like this button is ready for your custom script.")




#==================================================
#🚫 DELETE BELOW

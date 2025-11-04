# -*- coding: utf-8 -*-
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""


__title__   = "Elbow (Tee)"
__doc__     = """
Version = 00.00.01
Date    = 2025-10-27
________________________________________________________________
Description:

Eventually should select all elbow tees in the view and tag them.

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

print("""Finally somebody let me out of my cage!
        Now time for me is nothing 'cause I'm counting no age
        Nah, I coudn't be there, now you shouldn't be scared
        I'm good at repairs and I'm under each snare
        Intangible, bet you didn't think, so i command you to
        Panoramic view, look, I'll make it all manageable
        Pick and choose, sit and lose, all you different crews
        Chicks and dudes, who you think is really kickin' tunes?
        Picture you getting down in a picture tube
        Like you lit a fuse, you think it's fictional?
        Mystical? Mabye, spiritual hero
        Who appears in you to clearly your view when you're too crazy?
        Lifeless, to know the definnition for what life is
        Pricess to you because I put you in the hype shit
        You like it? Gun smokin', righteous with one toke
        Get psychic among those possess you with one go
        I ain't happy, I'm feeling glad
        I got sunshine in a bag""" )




#==================================================
#🚫 DELETE BELOW

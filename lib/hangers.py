# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

HANGER_DATA = {
    "Round Strap Hanger": {
        "shape": "round",
        "plumb": 1,
    },
    "Strut Hanger 1.625 Rnd": {
        "shape": "round",
        "plumb": 2,

    },
    "Half Strap Hanger": {
        "shape": "square",
        "plumb": 2,
    },

    "Strut Hanger 1.625 Rec": {
        "shape": "square",
        "plumb": 2,
    },

    "Double Strut Hanger - 1.625": {
        "shape": "square",
        "plumb": 2,
    },
}

class Hanger:
    def __init__(self, family_name):
        if family_name not in HANGER_DATA:
            raise ValueError('Hanger type {} not found'.format(family_name))
        
        attrs = HANGER_DATA[family_name]
        for key, value in attrs.items():
            setattr(self, key, value)
            
        self.family_name = family_name
        
    def __repr__(self):
        return f"<Hanger {self.family_name}: {self.__dict__}>"
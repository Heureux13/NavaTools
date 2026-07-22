# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2026 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================


from config.parameters_registry import NDBS_D_BOTTOM_EXTENSION, NDBS_D_TOP_EXTENSION

class DuctParameters:
    DUCT_PARAMETERS = {
        "ext_bottom": NDBS_D_BOTTOM_EXTENSION,
        "ext_top": NDBS_D_TOP_EXTENSION,
        "ext_right": NDBS_D_RIGHT_EXTENSION,
        "ext_left": NDBS_D_LEFT_EXTENSION,
    }

    def creat_duct_map(self):
        element_map = {}

        for fab_element in self.collected_ducts:
            element_data = {
                "family": self._family(fab_element),
                "object": fab_element,
            }

            for key, param_name in self.DUCT_PARAMETERS.items():
                element_data(key) = self._get_param(fab_element,param_name)
            element_map[fab_element.Id] = element_data

         return element_map
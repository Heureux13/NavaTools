# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

# Edit any parameter directly here.
DUCT_CATEGORIES = (
    'OST_DuctTerminal',
    'OST_FabricationDuctwork',
    'OST_MechanicalEquipment',
)

DUCT_PIPE_CATEGORIES = (
    'OST_MechanicalEquipment'
    'OST_FabricationDuctwork'
)

PIPE_CATEGORIES = (
    'OST_FabricationPipework'
)



ROOM_CATEGORIES = ('OST_Rooms',)

# fmt: off
# autopep8: off
parameter_bindings = {
    '_UMI_BBM_Author'            : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Ceiling'           : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_CFMEA'             : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_CFMSA'             : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Class'             : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Color'             : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Comments'          : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Damper'            : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Device'            : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Duty'              : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Face'              : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Fan'               : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_GPM'               : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Hand'              : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_HP'                : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_K'                 : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_KW'                : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Label'             : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Layer'             : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Make'              : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Material'          : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Model'             : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Mount'             : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Neck'              : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Notes'             : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Number'            : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_PageLabel'         : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Paint'             : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Phase'             : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Qty'               : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Section'           : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Size'              : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Sleeve'            : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Slot'              : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Space'             : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Status'            : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Subclass'          : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Subject'           : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_System'            : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Trade'             : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Type'              : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_Unit'              : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_VAV'               : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_BBM_VPH'               : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_JDG_OffsetLeft'        : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_JDG_OffsetRight'       : {'group': 'Data',    'categories': DUCT_CATEGORIES},
    '_UMI_PYT_AspectRatio'       : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_CFM'               : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_HeightPad'         : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_Label'             : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_Note0'             : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_Note1'             : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_Note2'             : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_Note3'             : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_Note4'             : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_Number'            : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_NumberFabrication' : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_NumberRun'         : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_NumberOrder'       : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_NumberSleeve'      : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_OffsetBottom'      : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_OffsetCenterH'     : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_OffsetCenterV'     : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_OffsetLeft'        : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_OffsetRight'       : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_OffsetTop'         : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_OffsetValue'       : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_SkipNumber'        : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_SkipTag'           : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_Sleeve'            : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_SleeveOpening'     : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_SleeveValue'       : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_WeightRun'         : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_WeightSupport'     : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_WeightPerFoot'     : {'group': 'General', 'categories': DUCT_CATEGORIES},
    '_UMI_PYT_CeilingHeight'     : {'group': 'Data',    'categories': ROOM_CATEGORIES},
    '_UMI_PYT_CeilingHeight0'    : {'group': 'Data',    'categories': ROOM_CATEGORIES},
    '_UMI_PYT_CeilingHeight1'    : {'group': 'Data',    'categories': ROOM_CATEGORIES},
    '_UMI_PYT_CeilingHeight2'    : {'group': 'Data',    'categories': ROOM_CATEGORIES},
    '_UMI_PYT_CeilingHeight3'    : {'group': 'Data',    'categories': ROOM_CATEGORIES},
    '_UMI_PYT_CeilingHeight4'    : {'group': 'Data',    'categories': ROOM_CATEGORIES},
    '_UMI_PYT_CeilingType'       : {'group': 'Data',    'categories': ROOM_CATEGORIES},
    '_UMI_PYT_CeilingType0'      : {'group': 'Data',    'categories': ROOM_CATEGORIES},
    '_UMI_PYT_CeilingType1'      : {'group': 'Data',    'categories': ROOM_CATEGORIES},
    '_UMI_PYT_CeilingType2'      : {'group': 'Data',    'categories': ROOM_CATEGORIES},
    '_UMI_PYT_CeilingType3'      : {'group': 'Data',    'categories': ROOM_CATEGORIES},
    '_UMI_PYT_CeilingType4'      : {'group': 'Data',    'categories': ROOM_CATEGORIES},
    '_UMI_PYT_RoomNumber'        : {'group': 'Data',    'categories': ROOM_CATEGORIES},
    '_UMI_PYT_RoomName'          : {'group': 'Data',    'categories': ROOM_CATEGORIES},
}
# fmt: on
# autopep8: on

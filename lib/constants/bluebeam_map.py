# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

"""
Bluebeam to Revit Column Mapping

Maps Bluebeam CSV column names to:
  - Revit schedule column header names (normalized)
  - Aliases for alternate column name variations
  - Expected parameter storage type: 'String', 'Integer', or 'Double'
"""


# Column name mapping: Revit schedule header -> (aliases, storage_type)
# Order matches CSV column order: Author,Class,Subclass,Subject,Layer,Check,Color,#,OG Label,Label,...
COLUMN_MAP = {
    'Author': {
        'aliases': [BBM_AUTHOR],
        'storage_type': 'String',
        'description': 'Author/creator of the item'
    },
    'Class': {
        'aliases': [BBM_CLASS],
        'storage_type': 'String',
        'description': 'Classification category'
    },
    'Subclass': {
        'aliases': [BBM_SUBCLASS],
        'storage_type': 'String',
        'description': 'Subclassification'
    },
    'Subject': {
        'aliases': [BBM_SUBJECT],
        'storage_type': 'String',
        'description': 'Subject/category (Schedule, Equipment, Unit, Fan, etc.)'
    },
    'Layer': {
        'aliases': [BBM_LAYER],
        'storage_type': 'String',
        'description': 'Layer assignment'
    },
    'Check': {
        'aliases': [BBM_CHECK],
        'storage_type': 'String',
        'description': 'Checked/unchecked status'
    },
    'Color': {
        'aliases': [BBM_COLOR],
        'storage_type': 'String',
        'description': 'Color assignment/code'
    },
    '#': {
        'aliases': [BBM_NUMBER],
        'storage_type': 'Integer',
        'description': 'Index/row number'
    },
    'OG Label': {
        'aliases': [BBM_OG_LABEL],
        'storage_type': 'String',
        'description': 'Original/legacy label'
    },
    'Label': {
        'aliases': [BBM_LABEL],
        'storage_type': 'String',
        'description': 'Primary key matching schedule elements to CSV rows'
    },
    'Page Label': {
        'aliases': [BBM_PAGE_LABEL],
        'storage_type': 'String',
        'description': 'Sheet/page reference label'
    },
    'Space': {
        'aliases': [BBM_SPACE],
        'storage_type': 'String',
        'description': 'Associated space/room number'
    },
    'Qty': {
        'aliases': [BBM_QTY],
        'storage_type': 'Integer',
        'description': 'Quantity'
    },
    'Make': {
        'aliases': [BBM_MAKE],
        'storage_type': 'String',
        'description': 'Equipment manufacturer'
    },
    'Model': {
        'aliases': [BBM_MODEL],
        'storage_type': 'String',
        'description': 'Equipment model/catalog number'
    },
    'Size': {
        'aliases': [BBM_SIZE],
        'storage_type': 'String',
        'description': 'Equipment/duct size'
    },
    'Neck': {
        'aliases': [BBM_NECK],
        'storage_type': 'String',
        'description': 'Neck size/dimension'
    },
    'Face': {
        'aliases': [BBM_FACE],
        'storage_type': 'String',
        'description': 'Face area/size'
    },
    'Mount': {
        'aliases': [BBM_MOUNT],
        'storage_type': 'String',
        'description': 'Mounting type (e.g. Ceiling, Floor, Vertical)'
    },
    'Ceiling': {
        'aliases': [BBM_CEILING],
        'storage_type': 'String',
        'description': 'Ceiling type'
    },
    'Type': {
        'aliases': [BBM_TYPE],
        'storage_type': 'String',
        'description': 'Equipment type (e.g. Direct, Belt, Inline)'
    },
    'Damper': {
        'aliases': [BBM_DAMPER],
        'storage_type': 'String',
        'description': 'Damper type (Control, Motorized, Electric, etc.)'
    },
    'Slot': {
        'aliases': [BBM_SLOT],
        'storage_type': 'String',
        'description': 'Slot or connection information'
    },
    'Hand': {
        'aliases': [BBM_HAND],
        'storage_type': 'String',
        'description': 'Handedness/orientation (Lh, Rh, etc.)'
    },
    'V - Ph': {
        'aliases': [BBM_VPH],
        'storage_type': 'String',
        'description': 'Voltage and phase (e.g. 120/1, 480/3)'
    },
    'Duty': {
        'aliases': [BBM_DUTY],
        'storage_type': 'String',
        'description': 'Operating duty (EA, SA, TA, Relief, etc.)'
    },
    'SA CFM': {
        'aliases': [BBM_SA_CFM],
        'storage_type': 'Double',
        'description': 'Supply air cubic feet per minute'
    },
    'EA CFM': {
        'aliases': [BBM_EA_CFM],
        'storage_type': 'Double',
        'description': 'Exhaust air cubic feet per minute'
    },
    'CFM': {
        'aliases': [BBM_CFM],
        'storage_type': 'Double',
        'description': 'Cubic feet per minute (airflow)'
    },
    'GPM': {
        'aliases': [BBM_GPM],
        'storage_type': 'Double',
        'description': 'Gallons per minute (water/fluid flow)'
    },
    'HP': {
        'aliases': [BBM_HP],
        'storage_type': 'Double',
        'description': 'Horsepower'
    },
    'Kw': {
        'aliases': [BBM_KW],
        'storage_type': 'Double',
        'description': 'Kilowatts (electrical power)'
    },
    'Sleeve': {
        'aliases': [BBM_SLEEVE],
        'storage_type': 'Double',
        'description': 'Sleeve/penetration information'
    },
    'K': {
        'aliases': [BBM_K],
        'storage_type': 'Double',
        'description': 'Insulation or utility flag'
    },
    'Material': {
        'aliases': [BBM_MATERIAL],
        'storage_type': 'String',
        'description': 'Material composition'
    },
    'Paint': {
        'aliases': [BBM_PAINT],
        'storage_type': 'String',
        'description': 'Paint/finish specification'
    },
    'Notes': {
        'aliases': [BBM_NOTES],
        'storage_type': 'String',
        'description': 'General notes and comments'
    },
    'Lock': {
        'aliases': [BBM_LOCK],
        'storage_type': 'String',
        'description': 'Lock status (Locked/Unlocked)'
    },
    'Status': {
        'aliases': [BBM_STATUS],
        'storage_type': 'String',
        'description': 'Status designation'
    },
    'Phase': {
        'aliases': [BBM_PHASE],
        'storage_type': 'String',
        'description': 'Project phase'
    },
    'Section': {
        'aliases': [BBM_SECTION],
        'storage_type': 'String',
        'description': 'Section designation'
    },
    'System': {
        'aliases': [BBM_SYSTEM],
        'storage_type': 'String',
        'description': 'HVAC system assignment'
    },
    'Unit': {
        'aliases': [BBM_UNIT],
        'storage_type': 'String',
        'description': 'Unit/component identifier'
    },
    'Fan': {
        'aliases': [BBM_FAN],
        'storage_type': 'String',
        'description': 'Fan designation'
    },
    'Device': {
        'aliases': [BBM_DEVICE],
        'storage_type': 'String',
        'description': 'Device type/designation'
    },
    'VAV': {
        'aliases': [BBM_VAV],
        'storage_type': 'String',
        'description': 'VAV box assignment'
    },
    'Page Index': {
        'aliases': [BBM_PAGE_INDEX],
        'storage_type': 'Integer',
        'description': 'Page index/number in document'
    },
    'Comments': {
        'aliases': [BBM_COMMENTS],
        'storage_type': 'String',
        'description': 'Additional comments (separate from Notes)'
    },
    'Trade': {
        'aliases': [BBM_TRADE],
        'storage_type': 'String',
        'description': 'Trade/discipline designation'
    },
}

# Legacy alias dict for backward compatibility with schedule column name mapping
SOURCE_HEADER_ALIASES = {
    'cfm': ['cfm', 'sa cfm', 'ra cfm'],
    'v/ph': ['v - ph'],
}

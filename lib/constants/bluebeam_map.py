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

# Explicit alias constants (stable for IronPython import and linting)
BBM_AUTHOR = '_author'
BBM_CLASS = '_class'
BBM_SUBCLASS = '_subclass'
BBM_SUBJECT = '_subject'
BBM_LAYER = '_layer'
BBM_CHECK = '_check'
BBM_COLOR = '_color'
BBM_NUMBER = '_number'
BBM_OG_LABEL = '_og_label'
BBM_LABEL = '_label'
BBM_PAGE_LABEL = '_page_label'
BBM_SPACE = '_space'
BBM_QTY = '_qty'
BBM_MAKE = '_make'
BBM_MODEL = '_model'
BBM_SIZE = '_size'
BBM_NECK = '_neck'
BBM_FACE = '_face'
BBM_MOUNT = '_mount'
BBM_CEILING = '_ceiling'
BBM_TYPE = '_type'
BBM_DAMPER = '_damper'
BBM_SLOT = '_slot'
BBM_HAND = '_hand'
BBM_VPH = '_v_ph'
BBM_DUTY = '_duty'
BBM_SA_CFM = '_sa_cfm'
BBM_EA_CFM = '_ea_cfm'
BBM_CFM = '_cfm'
BBM_GPM = '_gpm'
BBM_HP = '_hp'
BBM_KW = '_kw'
BBM_SLEEVE = '_sleeve'
BBM_K = '_k'
BBM_MATERIAL = '_material'
BBM_PAINT = '_paint'
BBM_NOTES = '_notes'
BBM_LOCK = '_lock'
BBM_STATUS = '_status'
BBM_PHASE = '_phase'
BBM_SECTION = '_section'
BBM_SYSTEM = '_system'
BBM_UNIT = '_unit'
BBM_FAN = '_fan'
BBM_DEVICE = '_device'
BBM_VAV = '_vav'
BBM_PAGE_INDEX = '_page_index'
BBM_COMMENTS = '_comments'
BBM_TRADE = '_trade'


def _bbm_alias_for(column_name):
    """Create default Revit parameter alias from Bluebeam header."""
    text = column_name.strip().lower()
    # keep only alnum and underscores, convert separators to underscore
    sanitized = []
    for ch in text:
        if ch.isalnum():
            sanitized.append(ch)
        else:
            sanitized.append('_')
    alias = ''.join(sanitized)
    while '__' in alias:
        alias = alias.replace('__', '_')
    alias = alias.strip('_')
    return '_{}'.format(alias) if alias else '_value'


_BLUEBEAM_COLUMNS = [
    'Author', 'Class', 'Subclass', 'Subject', 'Layer', 'Check', 'Color', '#',
    'OG Label', 'Label', 'Page Label', 'Space', 'Qty', 'Make', 'Model', 'Size',
    'Neck', 'Face', 'Mount', 'Ceiling', 'Type', 'Damper', 'Slot', 'Hand',
    'V - Ph', 'Duty', 'SA CFM', 'EA CFM', 'CFM', 'GPM', 'HP', 'Kw', 'Sleeve',
    'K', 'Material', 'Paint', 'Notes', 'Lock', 'Status', 'Phase', 'Section',
    'System', 'Unit', 'Fan', 'Device', 'VAV', 'Page Index', 'Comments', 'Trade'
]


for _col in _BLUEBEAM_COLUMNS:
    const_name = 'BBM_' + ''.join(
        ch if ch.isalnum() else '_' for ch in _col.upper()
    ).strip('_')
    while '__' in const_name:
        const_name = const_name.replace('__', '_')
    if const_name not in globals():
        globals()[const_name] = _bbm_alias_for(_col)

# Backward-compat constant names used by COLUMN_MAP.
# Auto-generation does not produce these exact legacy tokens.
if 'BBM_NUMBER' not in globals():
    BBM_NUMBER = _bbm_alias_for('#')
if 'BBM_VPH' not in globals():
    BBM_VPH = _bbm_alias_for('V - Ph')


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
    'v/ph': ['v - ph'],
}

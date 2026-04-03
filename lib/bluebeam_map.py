# -*- coding: utf-8 -*-

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
        'aliases': ['_author'],
        'storage_type': 'String',
        'description': 'Author/creator of the item'
    },
    'Class': {
        'aliases': ['_class'],
        'storage_type': 'String',
        'description': 'Classification category'
    },
    'Subclass': {
        'aliases': ['_subclass'],
        'storage_type': 'String',
        'description': 'Subclassification'
    },
    'Subject': {
        'aliases': ['_subject'],
        'storage_type': 'String',
        'description': 'Subject/category (Schedule, Equipment, Unit, Fan, etc.)'
    },
    'Layer': {
        'aliases': ['_layer'],
        'storage_type': 'String',
        'description': 'Layer assignment'
    },
    'Check': {
        'aliases': ['_check'],
        'storage_type': 'String',
        'description': 'Checked/unchecked status'
    },
    'Color': {
        'aliases': ['_color'],
        'storage_type': 'String',
        'description': 'Color assignment/code'
    },
    '#': {
        'aliases': ['_#'],
        'storage_type': 'Integer',
        'description': 'Index/row number'
    },
    'OG Label': {
        'aliases': ['_og_label'],
        'storage_type': 'String',
        'description': 'Original/legacy label'
    },
    'Label': {
        'aliases': ['_label'],
        'storage_type': 'String',
        'description': 'Primary key matching schedule elements to CSV rows'
    },
    'Page Label': {
        'aliases': ['_page_label'],
        'storage_type': 'String',
        'description': 'Sheet/page reference label'
    },
    'Space': {
        'aliases': ['_space'],
        'storage_type': 'String',
        'description': 'Associated space/room number'
    },
    'Qty': {
        'aliases': ['_qty'],
        'storage_type': 'Integer',
        'description': 'Quantity'
    },
    'Make': {
        'aliases': ['_make'],
        'storage_type': 'String',
        'description': 'Equipment manufacturer'
    },
    'Model': {
        'aliases': ['_model'],
        'storage_type': 'String',
        'description': 'Equipment model/catalog number'
    },
    'Size': {
        'aliases': ['_size'],
        'storage_type': 'String',
        'description': 'Equipment/duct size'
    },
    'Neck': {
        'aliases': ['_neck'],
        'storage_type': 'String',
        'description': 'Neck size/dimension'
    },
    'Face': {
        'aliases': ['_face'],
        'storage_type': 'String',
        'description': 'Face area/size'
    },
    'Mount': {
        'aliases': ['_mount'],
        'storage_type': 'String',
        'description': 'Mounting type (e.g. Ceiling, Floor, Vertical)'
    },
    'Ceiling': {
        'aliases': ['_ceiling'],
        'storage_type': 'String',
        'description': 'Ceiling type'
    },
    'Type': {
        'aliases': ['_type'],
        'storage_type': 'String',
        'description': 'Equipment type (e.g. Direct, Belt, Inline)'
    },
    'Damper': {
        'aliases': ['_damper'],
        'storage_type': 'String',
        'description': 'Damper type (Control, Motorized, Electric, etc.)'
    },
    'Slot': {
        'aliases': ['_slot'],
        'storage_type': 'String',
        'description': 'Slot or connection information'
    },
    'Hand': {
        'aliases': ['_hand'],
        'storage_type': 'String',
        'description': 'Handedness/orientation (Lh, Rh, etc.)'
    },
    'V - Ph': {
        'aliases': ['_v_ph'],
        'storage_type': 'String',
        'description': 'Voltage and phase (e.g. 120/1, 480/3)'
    },
    'Duty': {
        'aliases': ['_duty'],
        'storage_type': 'String',
        'description': 'Operating duty (EA, SA, TA, Relief, etc.)'
    },
    'SA CFM': {
        'aliases': ['_sa_cfm'],
        'storage_type': 'Double',
        'description': 'Supply air cubic feet per minute'
    },
    'EA CFM': {
        'aliases': ['_ea_cfm'],
        'storage_type': 'Double',
        'description': 'Exhaust air cubic feet per minute'
    },
    'CFM': {
        'aliases': ['_cfm'],
        'storage_type': 'Double',
        'description': 'Cubic feet per minute (airflow)'
    },
    'GPM': {
        'aliases': ['_gpm'],
        'storage_type': 'Double',
        'description': 'Gallons per minute (water/fluid flow)'
    },
    'HP': {
        'aliases': ['_hp'],
        'storage_type': 'Double',
        'description': 'Horsepower'
    },
    'Kw': {
        'aliases': ['_kw'],
        'storage_type': 'Double',
        'description': 'Kilowatts (electrical power)'
    },
    'Sleeve': {
        'aliases': ['_sleeve'],
        'storage_type': 'Double',
        'description': 'Sleeve/penetration information'
    },
    'K': {
        'aliases': ['_k'],
        'storage_type': 'Double',
        'description': 'Insulation or utility flag'
    },
    'Material': {
        'aliases': ['_material'],
        'storage_type': 'String',
        'description': 'Material composition'
    },
    'Paint': {
        'aliases': ['_paint'],
        'storage_type': 'String',
        'description': 'Paint/finish specification'
    },
    'Notes': {
        'aliases': ['_notes'],
        'storage_type': 'String',
        'description': 'General notes and comments'
    },
    'Lock': {
        'aliases': ['_lock'],
        'storage_type': 'String',
        'description': 'Lock status (Locked/Unlocked)'
    },
    'Phase': {
        'aliases': ['_phase'],
        'storage_type': 'String',
        'description': 'Project phase'
    },
    'Section': {
        'aliases': ['_section'],
        'storage_type': 'String',
        'description': 'Section designation'
    },
    'System': {
        'aliases': ['_system'],
        'storage_type': 'String',
        'description': 'HVAC system assignment'
    },
    'Unit': {
        'aliases': ['_unit'],
        'storage_type': 'String',
        'description': 'Unit/component identifier'
    },
    'Fan': {
        'aliases': ['_fan'],
        'storage_type': 'String',
        'description': 'Fan designation'
    },
    'Device': {
        'aliases': ['_device'],
        'storage_type': 'String',
        'description': 'Device type/designation'
    },
    'VAV': {
        'aliases': ['_vav'],
        'storage_type': 'String',
        'description': 'VAV box assignment'
    },
    'Page Index': {
        'aliases': ['_page_index'],
        'storage_type': 'Integer',
        'description': 'Page index/number in document'
    },
    'Comments': {
        'aliases': ['_comments'],
        'storage_type': 'String',
        'description': 'Additional comments (separate from Notes)'
    },
    'Trade': {
        'aliases': ['_trade'],
        'storage_type': 'String',
        'description': 'Trade/discipline designation'
    },
}

# Legacy alias dict for backward compatibility with schedule column name mapping
SOURCE_HEADER_ALIASES = {
    'cfm': ['cfm', 'sa cfm', 'ra cfm'],
    'v/ph': ['v - ph'],
}

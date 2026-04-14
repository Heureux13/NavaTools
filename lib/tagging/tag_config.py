# -*- coding: utf-8 -*-

from config.parameters_registry import *

"""Shared tag slot configuration used by fitting and joint tagging scripts."""

# fmt: off
# autopep8: off
SLOT_ACCESS_PANEL           = 'ACCESS_PANEL'
SLOT_BOD                    = 'BOD'
SLOT_BOD_LEFT               = 'BOD_LEFT'
SLOT_BOD_RIGHT              = 'BOD_RIGHT'
SLOT_CANVAS                 = 'CANVAS'
SLOT_DAMPER_CONTROL         = 'DAMPER_CONTROL'
SLOT_DAMPER_VOLUME          = 'DAMPER_VOLUME'
SLOT_DEGREE                 = 'DEGREE'
SLOT_ENDCAP_SD              = 'ENDCAP_SD'
SLOT_ENDCAP_TDF             = 'ENDCAP_TDF'
SLOT_EQUIPMENT_PAD          = 'EQUIPMENT_PAD'
SLOT_EXT_BOT                = 'EXT_IN'
SLOT_EXT_LEFT               = 'EXT_LEFT'
SLOT_EXT_RIGHT              = 'EXT_RIGHT'
SLOT_EXT_TOP                = 'EXT_OUT'
SLOT_FANS                   = 'FANS'
SLOT_DAMPER_FIRE            = 'FIRE_DAMPER'
SLOT_GRD                    = 'GRD'
SLOT_GRD_CFM                = 'GRD_CFM'
SLOT_LENGTH                 = 'LENGTH'
SLOT_LENGTH_LEFT            = 'LENGTH_LEFT'
SLOT_LENGTH_RIGHT           = 'LENGTH_RIGHT'
SLOT_LOUVER                 = 'LOUVER'
SLOT_MAN_BARS               = 'MAN_BARS'
SLOT_MARK                   = 'MARK'
SLOT_NOTE                   = 'NOTE'
SLOT_NUMBER_BLUEBEAM        = 'NUMBER_BLUEBEAM'
SLOT_NUMBER_FABRICATION     = 'NUMBER_FABRICATION'
SLOT_NUMBER_SLEEVE          = 'NUMBER_SLEEVE'
SLOT_OFFSET                 = 'OFFSET'
SLOT_SIZE                   = 'SIZE'
SLOT_SIZE_LEFT              = 'SIZE_LEFT'
SLOT_SIZE_RIGHT             = 'SIZE_RIGHT'
SLOT_TAP                    = 'TAP'
SLOT_TRANSITION             = 'TRANSITION'
SLOT_UNIT                   = 'UNIT'
SLOT_VAV                    = 'VAV'
SLOT_WEIGHT                 = 'WEIGHT'
# fmt: on
# autopep8: on

SLOT_ALL = (
    SLOT_ACCESS_PANEL,
    SLOT_BOD,
    SLOT_BOD_LEFT,
    SLOT_BOD_RIGHT,
    SLOT_CANVAS,
    SLOT_DEGREE,
    SLOT_DAMPER_CONTROL,
    SLOT_DAMPER_FIRE,
    SLOT_DAMPER_VOLUME,
    SLOT_ENDCAP_SD,
    SLOT_ENDCAP_TDF,
    SLOT_EQUIPMENT_PAD,
    SLOT_EXT_BOT,
    SLOT_EXT_LEFT,
    SLOT_EXT_RIGHT,
    SLOT_EXT_TOP,
    SLOT_FANS,
    SLOT_GRD,
    SLOT_GRD_CFM,
    SLOT_LENGTH,
    SLOT_LENGTH_LEFT,
    SLOT_LENGTH_RIGHT,
    SLOT_LOUVER,
    SLOT_MAN_BARS,
    SLOT_MARK,
    SLOT_NOTE,
    SLOT_NUMBER_BLUEBEAM,
    SLOT_NUMBER_FABRICATION,
    SLOT_NUMBER_SLEEVE,
    SLOT_OFFSET,
    SLOT_TAP,
    SLOT_TRANSITION,
    SLOT_SIZE,
    SLOT_SIZE_LEFT,
    SLOT_SIZE_RIGHT,
    SLOT_UNIT,
    SLOT_VAV,
    SLOT_WEIGHT,
)


DEFAULT_TAG_SLOT_CANDIDATES = {
    SLOT_ACCESS_PANEL: [
        ('_Tag.DT_AccessPanel', 'Center'),
    ],
    SLOT_BOD: [
        ('_Tag.DT_BOD', 'Center'),
    ],
    SLOT_BOD_LEFT: [
        ('_Tag.DT_BOD', 'Left'),
    ],
    SLOT_BOD_RIGHT: [
        ('_Tag.DT_BOD', 'Right'),
    ],
    SLOT_CANVAS: [
        ('_Tag.DT_Canvas', 'Center'),
    ],
    SLOT_DEGREE: [
        ('_Tag.DT_Degree', 'Center'),
    ],
    SLOT_ENDCAP_SD: [
        ('_Tag.DT_EndcapSD', 'Center'),
    ],
    SLOT_ENDCAP_TDF: [
        ('_Tag.DT_EndcapTDF', 'Center'),
    ],
    SLOT_DAMPER_CONTROL: [
        ('_Tag.DV_DamperControl', 'Center'),
    ],
    SLOT_DAMPER_FIRE: [
        ('_Tag.DV_DamperFire', 'Center'),
    ],
    SLOT_DAMPER_VOLUME: [
        ('_Tag.DV_DamperVolume', 'Center'),
    ],
    SLOT_EQUIPMENT_PAD: [
        ('_Tag.EQ_EquipmentPad', 'Black'),
    ],
    SLOT_EXT_BOT: [
        ('_Tag.DT_ExtensionBottom', 'Center'),
    ],
    SLOT_EXT_LEFT: [
        ('_Tag.DT_ExtensionLeft', 'Center'),
    ],
    SLOT_EXT_RIGHT: [
        ('_Tag.DT_ExtensionRight', 'Center'),
    ],
    SLOT_EXT_TOP: [
        ('_Tag.DT_ExtensionTop', 'Center'),
    ],
    SLOT_FANS: [
        ('_Tag.EQ_Fans', 'Black'),
    ],
    SLOT_GRD: [
        ('_Tag.DV_GRD', 'w/ flow'),
    ],
    SLOT_GRD_CFM: [
        ('_Tag.DV_GRD', 'w/o flow'),
    ],
    SLOT_LENGTH: [
        ('_Tag.DT_Length', 'Center'),
    ],
    SLOT_LENGTH_LEFT: [
        ('_Tag.DT_Length', 'Left'),
    ],
    SLOT_LENGTH_RIGHT: [
        ('_Tag.DT_Length', 'Right'),
    ],
    SLOT_LOUVER: [
        ('_Tag.DV_Louver', 'Black'),
    ],
    SLOT_MAN_BARS: [
        ('_Tag.DV_ManBars', 'Center'),
    ],
    SLOT_MARK: [
        ('_Tag.DT_Mark', 'Black'),
    ],
    SLOT_NOTE: [
        ('_Tag.DT_Note', '1 note'),
    ],
    SLOT_NUMBER_BLUEBEAM: [
        ('_Tag.DT_NumberBluebeam', 'Center'),
    ],
    SLOT_NUMBER_FABRICATION: [
        ('_Tag.DT_NumberFabrication', 'Center'),
    ],
    SLOT_NUMBER_SLEEVE: [
        ('_Tag.DT_NumberSleeve', 'Center'),
    ],
    SLOT_OFFSET: [
        ('_Tag.DT_Offset', 'Center'),
    ],
    SLOT_TAP: [
        ('_Tag.DT_Tap', 'Center'),
    ],
    SLOT_TRANSITION: [
        ('_Tag.DT_Transition', 'Center'),
    ],
    SLOT_SIZE: [
        ('_Tag.DT_Size', 'Center'),
    ],
    SLOT_SIZE_LEFT: [
        ('_Tag.DT_Size', 'Left'),
    ],
    SLOT_SIZE_RIGHT: [
        ('_Tag.DT_Size', 'Right'),
    ],
    SLOT_UNIT: [
        ('_Tag.EQ_Mark', 'Black'),
    ],
    SLOT_VAV: [
        ('_Tag.EQ_VAV', 'Black'),
    ],
    SLOT_WEIGHT: [
        ('_Tag.DT_Weight', 'Defaut'),
    ],
}

DEFAULT_TAG_SKIP_PARAMETERS = {
    PYT_SKIP_TAG: ['skip'],
}

DEFAULT_NUMBER_SKIP_PARAMETERS = {
    PYT_SKIP_NUMBER: ['skip'],
}

DEFAULT_PARAMETER_HIERARCHY = [
    RVT_TYPE_MARK,
    RVT_MARK,
    PYT_LABEL,
]

WRITE_PARAMETER = BBM_LABEL

STRAIGHT_JOINT_FAMILIES = {
    'round duct',
    'spiral duct',
    'spiral tube',
    'straight',
    'tube',
}

DEFAULT_JOINT_TAG_SLOTS = [
    SLOT_BOD,
    SLOT_LENGTH,
    SLOT_SIZE,
]

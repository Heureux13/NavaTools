# -*- coding: utf-8 -*-

from config.parameters_registry import *

"""Shared tag slot configuration used by fitting and joint tagging scripts."""

# fmt: off
# autopep8: off
SLOT_BOD                    = 'BOD'
SLOT_BOD_CENTER             = 'BOD_CENTER'
SLOT_BOD_LEFT               = 'BOD_LEFT'
SLOT_BOD_RIGHT              = 'BOD_RIGHT'
SLOT_DEGREE                 = 'DEGREE'
SLOT_EXT_IN                 = 'EXT_IN'
SLOT_EXT_LEFT               = 'EXT_LEFT'
SLOT_EXT_OUT                = 'EXT_OUT'
SLOT_EXT_RIGHT              = 'EXT_RIGHT'
SLOT_LENGTH                 = 'LENGTH'
SLOT_LENGTH_CENTER          = 'LENGTH_CENTER'
SLOT_LENGTH_LEFT            = 'LENGTH_LEFT'
SLOT_LENGTH_RIGHT           = 'LENGTH_RIGHT'
SLOT_MARK                   = 'MARK'
SLOT_MARK_NOTE              = 'MARK_NOTE'
SLOT_SIZE                   = 'SIZE'
SLOT_SIZE_CENTER            = 'SIZE_CENTER'
SLOT_SIZE_LEFT              = 'SIZE_LEFT'
SLOT_SIZE_RIGHT             = 'SIZE_RIGHT'
SLOT_TRAN                   = 'TRAN'
SLOT_TYPE_MARK              = 'TYPE_MARK'
SLOT_WEIGHT                 = 'WEIGHT'
# fmt: on
# autopep8: on

SLOT_ALL = (
    SLOT_BOD,
    SLOT_BOD_CENTER,
    SLOT_BOD_LEFT,
    SLOT_BOD_RIGHT,
    SLOT_DEGREE,
    SLOT_EXT_IN,
    SLOT_EXT_LEFT,
    SLOT_EXT_OUT,
    SLOT_EXT_RIGHT,
    SLOT_LENGTH,
    SLOT_LENGTH_CENTER,
    SLOT_LENGTH_LEFT,
    SLOT_LENGTH_RIGHT,
    SLOT_MARK,
    SLOT_MARK_NOTE,
    SLOT_SIZE,
    SLOT_SIZE_CENTER,
    SLOT_SIZE_LEFT,
    SLOT_SIZE_RIGHT,
    SLOT_TRAN,
    SLOT_TYPE_MARK,
    SLOT_WEIGHT
)


DEFAULT_TAG_SLOT_CANDIDATES = {
    SLOT_BOD: [
        "-FabDuct_BOD_Tag",
        "_umi_BOD",
        '_Tag.DT_BOD'
    ],
    SLOT_BOD_CENTER: [
        "_umi_BOD_CENTER",
        '_Tag.DT_BOD.Center'
    ],
    SLOT_BOD_LEFT: [
        "_umi_BOD_LEFT",
        '_Tag.DT_BOD.Left'
    ],
    SLOT_BOD_RIGHT: [
        "_umi_BOD_RIGHT",
        '_Tag.DT_BOD.Right'
    ],
    SLOT_DEGREE: [
        "-FabDuct_DEGREE_MV_Tag",
        "_umi_ANGLE",
        '_Tag.DT_Degree'
    ],
    SLOT_EXT_IN: [
        "-FabDuct_EXT IN_MV_Tag",
        "_umi_EXTENSION_BOTTOM",
        '_Tag.DT_Ext.Bottom'
    ],
    SLOT_EXT_LEFT: [
        "-FabDuct_EXT LEFT_MV_Tag",
        "_umi_EXTENSION_LEFT",
        '_Tag.DT_Ext.Left'
    ],
    SLOT_EXT_OUT: [
        "-FabDuct_EXT OUT_MV_Tag",
        "_umi_EXTENSION_TOP",
        '_Tag.DT_Ext.Top'
    ],
    SLOT_EXT_RIGHT: [
        "-FabDuct_EXT RIGHT_MV_Tag",
        "_umi_EXTENSION_RIGHT",
        '_Tag.DT_Ext.Right'
    ],
    SLOT_LENGTH: [
        "-FabDuct_LENGTH_Tag",
        "_umi_LENGTH",
        '_Tag.DT_Length'
    ],
    SLOT_LENGTH_CENTER: [
        "-FabDuct_LENGTH_Tag",
        "_umi_LENGTH_CENTER",
        '_Tag.DT_Length.Center'
    ],
    SLOT_LENGTH_LEFT: [
        "-FabDuct_LENGTH_Tag",
        "_umi_LENGTH_LEFT",
        '_Tag.DT_Length.Left'
    ],
    SLOT_LENGTH_RIGHT: [
        "-FabDuct_LENGTH_Tag",
        "_umi_LENGTH_RIGHT",
        '_Tag.DT_Length.Right'
    ],
    SLOT_MARK: [
        "-FabDuct_MARK_Tag",
        "_umi_MARK",
        '_Tag.DT_Mark'
    ],
    SLOT_MARK_NOTE: [
        "_umi_MARK_note",
        '_Tag.DT_Mark.Note'
    ],
    SLOT_SIZE: [
        "-FabDuct_SIZE_Tag",
        "_umi_SIZE",
        '_Tag.DT_Size'
    ],
    SLOT_TRAN: [
        "-FabDuct_TRAN_MV_Tag",
        "_umi_OFFSET",
        '_Tag.DT_Offset.Value'
    ],
    SLOT_TYPE_MARK: [
        "-FabDuct_TM_MV_Tag",
        "_umi_TYPE_MARK",
        '_Tag.DT_Type.Mark'
    ],
    SLOT_WEIGHT: [
        "-FabDuct_WEIGHT_Tag",
        "_umi_WEIGHT",
        '_Tag.DT_Weight'
    ],
}

DEFAULT_SKIP_PARAMETERS = {
    PYT_SKIP_NUMBER: ['skip'],
    PYT_SKIP_TAG: ['skip'],
}

DEFAULT_PARAMETER_HIERARCHY = [
    RVT_TYPE_MARK,
    RVT_MARK,
    BBM_LABEL,
]

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

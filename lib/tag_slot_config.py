# -*- coding: utf-8 -*-

"""Shared tag slot configuration used by fitting and joint tagging scripts."""

SLOT_BOD = 'BOD'
SLOT_DEGREE = 'DEGREE'
SLOT_EXT_IN = 'EXT_IN'
SLOT_EXT_LEFT = 'EXT_LEFT'
SLOT_EXT_OUT = 'EXT_OUT'
SLOT_EXT_RIGHT = 'EXT_RIGHT'
SLOT_LENGTH = 'LENGTH'
SLOT_MARK = 'MARK'
SLOT_SIZE = 'SIZE'
SLOT_TRAN = 'TRAN'
SLOT_MARK_TYPE = 'TYPE_MARK'
SLOT_WEIGHT = 'WEIGHT'


DEFAULT_TAG_SLOT_CANDIDATES = {
    SLOT_BOD: [
        "-FabDuct_BOD_Tag",
        "_umi_BOD"
    ],
    SLOT_DEGREE: [
        "-FabDuct_DEGREE_MV_Tag",
        "_umi_ANGLE"
    ],
    SLOT_EXT_IN: [
        "-FabDuct_EXT IN_MV_Tag",
        "_umi_EXTENSION_BOTTOM"
    ],
    SLOT_EXT_LEFT: [
        "-FabDuct_EXT LEFT_MV_Tag",
        "_umi_EXTENSION_LEFT"
    ],
    SLOT_EXT_OUT: [
        "-FabDuct_EXT OUT_MV_Tag",
        "_umi_EXTENSION_TOP"
    ],
    SLOT_EXT_RIGHT: [
        "-FabDuct_EXT RIGHT_MV_Tag",
        "_umi_EXTENSION_RIGHT"
    ],
    SLOT_LENGTH: [
        "-FabDuct_LENGTH_Tag",
        "_umi_LENGTH"
    ],
    SLOT_MARK: [
        "-FabDuct_MARK_Tag",
        "_umi_MARK"
    ],
    SLOT_SIZE: [
        "-FabDuct_SIZE_Tag",
        "_umi_SIZE"
    ],
    SLOT_TRAN: [
        "-FabDuct_TRAN_MV_Tag",
        "_umi_OFFSET"
    ],
    SLOT_MARK_TYPE: [
        "-FabDuct_TYPE_MARK_MV_Tag",
        "_umi_TYPE_MARK"
    ],
    SLOT_WEIGHT: [
        "-FabDuct_WEIGHT_Tag",
        "_umi_WEIGHT"
    ],
}

DEFAULT_SKIP_PARAMETERS = {
    '_duct_tag':        ['skip', 'skip n/a'],
    '_duct_tag_offset': ['skip', 'skip n/a'],
    'mark':             ['skip', 'skip n/a'],
}

DEFAULT_PARAMETER_HIERARCHY = [
    'mark',
    'type mark',
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

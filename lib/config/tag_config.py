# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from config.parameters_registry import *
from pyrevit import revit
import importlib

"""Shared tag slot configuration used by fitting and joint tagging scripts."""

USER_CONFIG_MODULES = {
    'goolsby': 'config.tag_config_goolsby',
}


def _current_user_lower():
    try:
        return (revit.doc.Application.Username or "").strip().lower()
    except Exception:
        return ""


def _load_user_candidate():
    username = _current_user_lower()
    module_name = USER_CONFIG_MODULES.get(username)
    if not module_name:
        return None
    try:
        mod = importlib.import_module(module_name)
        return getattr(mod, "DEFAULT_TAG_SLOT_CANDIDATES", None)
    except Exception:
        return None




# fmt: off
# autopep8: off

# Duct
SLOT_CANVAS                 = 'CANVAS'
SLOT_ENDCAP_SD              = 'ENDCAP_SD'
SLOT_ENDCAP_TDF             = 'ENDCAP_TDF'
SLOT_OFFSET                 = 'OFFSET'
SLOT_REDUCER                = 'REDUCER'
SLOT_TAP                    = 'TAP'
SLOT_TRANSITION             = 'TRANSITION'
SLOT_STACK                  = 'STACK'

# Devices
SLOT_ACCESS_PANEL           = 'ACCESS_PANEL'
SLOT_DAMPER_CONTROL         = 'DAMPER_CONTROL'
SLOT_DAMPER_FIRE            = 'FIRE_DAMPER'
SLOT_DAMPER_VOLUME          = 'DAMPER_VOLUME'
SLOT_GRD                    = 'GRD'
SLOT_GRD_CFM                = 'GRD_CFM'
SLOT_LOUVER                 = 'LOUVER'
SLOT_LOUVER_NOTE            = 'LOUVER_NOTE'
SLOT_MAN_BARS               = 'MAN_BARS'

# Equipment
SLOT_CONDENSER              = "CONDENSER"
SLOT_CONDENSER_NOTE         = "CONDENSER_NOTE"
SLOT_EQUIPMENT_PAD          = 'EQUIPMENT_PAD'
SLOT_FAN                    = "FAN"
SLOT_FAN_NOTE               = "FAN_NOTE"
SLOT_HEAT_PUMP              = "HEAT_PUMP"
SLOT_HEAT_PUMP_NOTE         = "HEAT_PUMP_NOTE"
SLOT_HEATER                 = "HEATER"
SLOT_HEATER_NOTE            = "HEATER_NOTE"
SLOT_HOOD                   = "HOOD"
SLOT_HOOD_NOTE              = "HOOD_NOTE"
SLOT_HUMIDIFIER             = "HUMIDIFIER"
SLOT_HUMIDIFIER_NOTE        = "HUMIDIFIER_NOTE"
SLOT_SPLIT                  = "SPLIT"
SLOT_SPLIT_NOTE             = "SPLIT_NOTE"
SLOT_UNIT                   = 'UNIT'
SLOT_UNIT_NOTE              = 'UNIT_NOTE'
SLOT_VALVE                  = 'VALVE'
SLOT_VALVE_NOTE             = 'VALVE_NOTE'
SLOT_VRF                    = "VRF"
SLOT_VRF_NOTE               = "VRF_NOTE"

# Parameters
SLOT_BOD                    = 'BOD'
SLOT_BOD_LEFT               = 'BOD_LEFT'
SLOT_BOD_RIGHT              = 'BOD_RIGHT'
SLOT_DEGREE                 = 'DEGREE'
SLOT_EXT_BOT                = 'EXT_IN'
SLOT_EXT_LEFT               = 'EXT_LEFT'
SLOT_EXT_RIGHT              = 'EXT_RIGHT'
SLOT_EXT_TOP                = 'EXT_OUT'
SLOT_LENGTH                 = 'LENGTH'
SLOT_LENGTH_LEFT            = 'LENGTH_LEFT'
SLOT_LENGTH_RIGHT           = 'LENGTH_RIGHT'
SLOT_NOTE                   = 'NOTE'
SLOT_NUMBER_BLUEBEAM        = 'NUMBER_BLUEBEAM'
SLOT_NUMBER_FABRICATION     = 'NUMBER_FABRICATION'
SLOT_NUMBER_SLEEVE          = 'NUMBER_SLEEVE'
SLOT_SIZE                   = 'SIZE'
SLOT_SIZE_LEFT              = 'SIZE_LEFT'
SLOT_SIZE_RIGHT             = 'SIZE_RIGHT'
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
    SLOT_FAN,
    SLOT_GRD,
    SLOT_GRD_CFM,
    SLOT_LENGTH,
    SLOT_LENGTH_LEFT,
    SLOT_LENGTH_RIGHT,
    SLOT_LOUVER,
    SLOT_LOUVER_NOTE,
    SLOT_MAN_BARS,
    SLOT_CONDENSER,
    SLOT_HEAT_PUMP,
    SLOT_HEATER,
    SLOT_HOOD,
    SLOT_HUMIDIFIER,
    SLOT_NOTE,
    SLOT_NUMBER_BLUEBEAM,
    SLOT_NUMBER_FABRICATION,
    SLOT_NUMBER_SLEEVE,
    SLOT_OFFSET,
    SLOT_REDUCER,
    SLOT_SPLIT,
    SLOT_TAP,
    SLOT_TRANSITION,
    SLOT_STACK,
    SLOT_SIZE,
    SLOT_SIZE_LEFT,
    SLOT_SIZE_RIGHT,
    SLOT_UNIT,
    SLOT_VALVE,
    SLOT_VRF,
    SLOT_WEIGHT,
    SLOT_CONDENSER_NOTE,
    SLOT_FAN_NOTE,
    SLOT_HEAT_PUMP_NOTE,
    SLOT_HEATER_NOTE,
    SLOT_HOOD_NOTE,
    SLOT_HUMIDIFIER_NOTE,
    SLOT_SPLIT_NOTE,
    SLOT_UNIT_NOTE,
    SLOT_VALVE_NOTE,
    SLOT_VRF_NOTE,
)


DEFAULT_TAG_SLOT_CANDIDATES = {
    SLOT_ACCESS_PANEL: [
        ('_Tag.DCT_AccessPanel', 'Center'),
        ('_Tag.DCT_AccessPanel', 'Default'),
    ],
    SLOT_BOD: [
        ('_Tag.DCT_BOD', 'Center'),
        ('_Tag.DCT_BOD', 'Default'),
    ],
    SLOT_BOD_LEFT: [
        ('_Tag.DCT_BOD', 'Left'),
    ],
    SLOT_BOD_RIGHT: [
        ('_Tag.DCT_BOD', 'Right'),
    ],
    SLOT_CANVAS: [
        ('_Tag.DCT_Canvas', 'Center'),
        ('_Tag.DCT_Canvas', 'Default'),
    ],
    SLOT_CONDENSER: [
        ('_Tag.EQP_Condenser', 'Black'),
    ],
    SLOT_DAMPER_CONTROL: [
        ('_Tag.DEV_DamperControl', 'Center'),
        ('_Tag.DEV_DamperControl', 'Default'),
    ],
    SLOT_DAMPER_FIRE: [
        ('_Tag.DEV_DamperFire', 'Center'),
        ('_Tag.DEV_DamperLife', 'Default'),
    ],
    SLOT_DAMPER_VOLUME: [
        ('_Tag.DEV_DamperVolume', 'Center'),
        ('_Tag.DEV_DamperVolume', 'Default'),
    ],
    SLOT_DEGREE: [
        ('_Tag.DCT_Degree', 'Center'),
        ('_Tag.DCT_Degree', 'Default'),
    ],
    SLOT_ENDCAP_SD: [
        ('_Tag.DCT_Endcap', 'S&D'),
    ],
    SLOT_ENDCAP_TDF: [
        ('_Tag.DCT_Endcap', 'TDF'),
    ],
    SLOT_EQUIPMENT_PAD: [
        ('_Tag.EQP_EquipmentPad', 'Black'),
    ],
    SLOT_EXT_BOT: [
        ('_Tag.DCT_ElbowExtension', 'Bottom'),
        ('_Tag.DCT_Extensions', 'Bottom'),
    ],
    SLOT_EXT_LEFT: [
        ('_Tag.DCT_ElbowExtension', 'Left'),
        ('_Tag.DCT_Extensions', 'Left'),
    ],
    SLOT_EXT_RIGHT: [
        ('_Tag.DCT_ElbowExtension', 'Right'),
        ('_Tag.DCT_Extensions', 'Right'),
    ],
    SLOT_EXT_TOP: [
        ('_Tag.DCT_ElbowExtension', 'Top'),
        ('_Tag.DCT_Extensions', 'Top'),
    ],
    SLOT_FAN: [
        ('_Tag.EQP_Fan', 'Black'),
    ],
    SLOT_FAN_NOTE: [
        ('_Tag.EQP_Fan', 'Note Default'),
    ],
    SLOT_GRD: [
        ('_Tag.DEV_GRD', 'CFM'),
        ('_Tag.DEV_GRD', 'Label'),
    ],
    SLOT_GRD_CFM: [
        ('_Tag.DEV_GRD', 'Center'),
        ('_Tag.DEV_GRD', 'CFM'),
    ],
    SLOT_HEAT_PUMP: [
        ('_Tag.EQP_HeatPump', 'Black'),
    ],
    SLOT_HEAT_PUMP_NOTE: [
        ('_Tag.EQP_HeatPump', 'Note Default'),
    ],
    SLOT_HEATER: [
        ('_Tag.EQP_Heater', 'Black'),
    ],
    SLOT_HEATER_NOTE: [
        ('_Tag.EQP_Heater', 'Note Default'),
    ],
    SLOT_HOOD: [
        ('_Tag.EQP_Hood', 'Black'),
    ],
    SLOT_HOOD_NOTE: [
        ('_Tag.EQP_Hood', 'Note Default'),
    ],
    SLOT_HUMIDIFIER: [
        ('_Tag.EQP_Humidifier', 'Black'),
    ],
    SLOT_HUMIDIFIER_NOTE: [
        ('_Tag.EQP_Humidifier', 'Note Default'),
    ],
    SLOT_LENGTH: [
        ('_Tag.DCT_Length', 'Center'),
        ('_Tag.DCT_Length', 'Default'),
    ],
    SLOT_LENGTH_LEFT: [
        ('_Tag.DCT_Length', 'Left'),
    ],
    SLOT_LENGTH_RIGHT: [
        ('_Tag.DCT_Length', 'Right'),
    ],
    SLOT_LOUVER: [
        ('_Tag.DEV_Louvers', 'Black'),
    ],
    SLOT_LOUVER_NOTE: [
        ('_Tag.DEV_Louvers', 'Note Default'),
    ],
    SLOT_MAN_BARS: [
        ('_Tag.DEV_ManBars', 'Center'),
        ('_Tag.DEV_ManBars', 'Default'),
    ],
    SLOT_NOTE: [
        ('_Tag.DCT_Note', '1 note'),
        ('_Tag.DCT_Note', 'Notes'),
    ],
    SLOT_NUMBER_BLUEBEAM: [
        ('_Tag.DCT_NumberBluebeam', 'Center'),
    ],
    SLOT_NUMBER_FABRICATION: [
        ('_Tag.DCT_NumberDuct', 'Small'),
    ],
    SLOT_NUMBER_SLEEVE: [
        ('_Tag.DCT_NumberSleeve', 'Green'),
        ('_Tag.DCT_NumberSleeve', 'Duct Penetration'),
    ],
    SLOT_OFFSET: [
        ('_Tag.DCT_Offset', 'Center'),
        ('_Tag.DCT_Offset', 'RND 1.00'),
    ],
    SLOT_REDUCER: [
        ('_Tag.DCT_Reducer', 'Default'),
    ],
    SLOT_SIZE: [
        ('_Tag.DCT_Size', 'Center'),
        ('_Tag.DCT_Size', 'Default'),
    ],
    SLOT_SIZE_LEFT: [
        ('_Tag.DCT_Size', 'Left'),
    ],
    SLOT_SIZE_RIGHT: [
        ('_Tag.DCT_Size', 'Right'),
    ],
    SLOT_STACK: [
        ('_Tag.DCT_Stack', '3_Center'),
    ],
    SLOT_SPLIT: [
        ('_Tag.EQP_Split', 'Black'),
    ],
    SLOT_SPLIT_NOTE: [
        ('_Tag.EQP_Split', 'Note Default'),
    ],
    SLOT_TAP: [
        ('_Tag.DCT_Tap', 'Center'),
        ('_Tag.DCT_Tap', 'Default'),
    ],
    SLOT_TRANSITION: [
        ('_Tag.DCT_Transition', 'Center'),
        ('_Tag.DCT_Transition', 'RND 1.00'),
    ],
    SLOT_UNIT: [
        ('_Tag.EQP_Unit', 'Black'),
    ],
    SLOT_UNIT_NOTE: [
        ('_Tag.EQP_Unit', 'Note Default'),
    ],
    SLOT_VALVE: [
        ('_Tag.EQP_Valve', 'Black'),
    ],
    SLOT_VALVE_NOTE: [
        ('_Tag.EQP_Valve', 'Note Default'),
    ],
    SLOT_VRF: [
        ('_Tag.EQP_VRF', 'Black'),
    ],
    SLOT_VRF_NOTE: [
        ('_Tag.EQP_VRF', 'BlackNote'),
    ],
    SLOT_WEIGHT: [
        ('_Tag.DCT_Weight', 'Default'),
    ],
}

# fmt: off
# autopep8: off

# Square families
FML_SQR_90_ELBOW                = '90° Elbow'
FML_SQR_ADJ_ELBOW               = 'Adjustable Elbow'
FML_SQR_BIRD_SCREEN             = 'Bird Screen'
FML_SQR_BOOT_SADDLE             = 'Boot Saddle'
FML_SQR_BOOT_TAP                = 'Boot Tap'
FML_SQR_BOOT_TAP_W_DAMPER       = 'Boot Tap - wDamper'
FML_SQR_CANVAS                  = 'Canvas'
FML_SQR_CONICAL_TAP             = 'Conical Tap'
FML_SQR_CONICAL_TAP_W_DAMPER    = 'Conical Tap - wDamper'
FML_SQR_DROP_CHEEK              = 'Drop Cheek'
FML_SQR_ENDCAP                  = 'End Cap'
FML_SQR_MAN_BAR                 = 'MAN BAR'
FML_SQR_OFFSET                  = 'Offset'
FML_SQR_OGEE                    = 'Ogee'
FML_SQR_PANTS                   = 'Pants'
FML_SQR_RADIUS_CLUSTER          = 'Radius Cluster'
FML_SQR_RADIUS_ELBOW            = 'Radius Elbow'
FML_SQR_ROOF_CURB               = 'Roof Curb'
FML_SQR_STRAIGHT                = 'Straight'
FML_SQR_TDF_ENDCAP              = 'TDF End Cap'
FML_SQR_TEE                     = 'Tee'
FML_SQR_TO_RND                  = 'Square to Ø'
FML_SQR_TRANSITION              = 'Transition'

# Round families
FML_RND_3_WAY                   = 'Three Way'
FML_RND_3_WAY_BRANCH            = '3 Way Branch'
FML_RND_BOOT_SADDLE             = 'Boot Saddle Tap'
FML_RND_BOX_SADDLE              = 'Box Saddle'
FML_RND_CANVAS                  = 'Ø Canvas'
FML_RND_CHINA_CAP               = 'China Cap'
FML_RND_CONICAL_TEE             = 'Conical Tee'
FML_RND_COUPLING                = 'Coupling'
FML_RND_COUPLING_FITTING        = 'Coupling - Fitting'
FML_RND_CROSS_TYPE_2            = 'Cross (Type 2)'
FML_RND_DAMPER                  = 'Round Damper'
FML_RND_DAMPER_VOLUME           = '8inch Long Coupler wDamper'
FML_RND_ELBOW_90_SR_STAMPED     = 'Elbow 90 SR - stamped'
FML_RND_ENDCAP                  = 'Ø End Cap'
FML_RND_GORED_ELBOW             = 'Gored Elbow'
FML_RND_OFFSET                  = 'Ø Offset'
FML_RND_REDNECK_REDUCER         = 'Redneck Reducer'
FML_RND_REDUCER                 = 'Reducer'
FML_RND_REDUCING_TEE            = 'Reducing Tee'
FML_RND_ROOF_JACK               = 'Roof Jack'
FML_RND_SADDLE                  = 'Saddle Tap'
FML_RND_SADDLE_HANGER           = 'Saddle Hanger'
FML_RND_STRAIGHT                = 'Spiral Duct'
FML_RND_WYE                     = 'WYE'

# Oval families
FML_OVL_COUPLING                = 'Oval Coupling'
FML_OVL_CONICAL_TAP_HARD        = 'Oval Conical Tap Hard'
FML_OVL_ELBOW_EASY              = 'Oval Elbow Easy'
FML_OVL_ELBOW_HARD              = 'Oval Elbow Hard'
FML_OVL_ENDCAP                  = 'Oval Cap End'
FML_OVL_STRAIGHT                = 'Oval Pipe (LP)'
FML_OVL_REDUCER                 = 'Oval Reducer'
FML_OVL_TO_ROUND                = 'Oval to Round'

# fmt: on
# autopep8: on


all_duct_families = {
    # Square
    FML_SQR_90_ELBOW: 'Square',
    FML_SQR_ADJ_ELBOW: 'Square',
    FML_SQR_BIRD_SCREEN: 'Square',
    FML_SQR_BOOT_SADDLE: 'Square',
    FML_SQR_BOOT_TAP: 'Square',
    FML_SQR_BOOT_TAP_W_DAMPER: 'Square',
    FML_SQR_CANVAS: 'Square',
    FML_SQR_CONICAL_TAP: 'Square',
    FML_SQR_CONICAL_TAP_W_DAMPER: 'Square',
    FML_SQR_DROP_CHEEK: 'Square',
    FML_SQR_ENDCAP: 'Square',
    FML_SQR_MAN_BAR: 'Square',
    FML_SQR_OGEE: 'Square',
    FML_SQR_OFFSET: 'Square',
    FML_SQR_PANTS: 'Square',
    FML_SQR_RADIUS_CLUSTER: 'Square',
    FML_SQR_RADIUS_ELBOW: 'Square',
    FML_SQR_ROOF_CURB: 'Square',
    FML_SQR_TO_RND: 'Square',
    FML_SQR_STRAIGHT: 'Square',
    FML_SQR_TDF_ENDCAP: 'Square',
    FML_SQR_TEE: 'Square',
    FML_SQR_TRANSITION: 'Square',
    # Round
    FML_RND_3_WAY: 'Round',
    FML_RND_3_WAY_BRANCH: 'Round',
    FML_RND_BOOT_SADDLE: 'Round',
    FML_RND_BOX_SADDLE: 'Round',
    FML_RND_CANVAS: 'Round',
    FML_RND_CHINA_CAP: 'Round',
    FML_RND_CONICAL_TEE: 'Round',
    FML_RND_DAMPER_VOLUME: 'Round',
    FML_RND_COUPLING: 'Round',
    FML_RND_COUPLING_FITTING: 'Round',
    FML_RND_CROSS_TYPE_2: 'Round',
    FML_RND_DAMPER: 'Round',
    FML_RND_ELBOW_90_SR_STAMPED: 'Round',
    FML_RND_ENDCAP: 'Round',
    FML_RND_GORED_ELBOW: 'Round',
    FML_RND_OFFSET: 'Round',
    FML_RND_REDNECK_REDUCER: 'Round',
    FML_RND_REDUCER: 'Round',
    FML_RND_REDUCING_TEE: 'Round',
    FML_RND_ROOF_JACK: 'Round',
    FML_RND_SADDLE: 'Round',
    FML_RND_SADDLE_HANGER: 'Round',
    FML_RND_STRAIGHT: 'Round',
    FML_RND_WYE: 'Round',
    # Oval
    FML_OVL_ENDCAP: 'Oval',
    FML_OVL_CONICAL_TAP_HARD: 'Oval',
    FML_OVL_COUPLING: 'Oval',
    FML_OVL_ELBOW_EASY: 'Oval',
    FML_OVL_ELBOW_HARD: 'Oval',
    FML_OVL_STRAIGHT: 'Oval',
    FML_OVL_REDUCER: 'Oval',
    FML_OVL_TO_ROUND: 'Oval',
}

DUCT_FAMILY_TAG_SLOTS = {
    # Square
    FML_SQR_90_ELBOW:            [SLOT_EXT_TOP, SLOT_EXT_BOT, SLOT_DEGREE],
    FML_SQR_ADJ_ELBOW:           [SLOT_EXT_TOP, SLOT_EXT_BOT, SLOT_DEGREE],
    FML_SQR_RADIUS_ELBOW:        [SLOT_DEGREE],
    FML_SQR_TEE:                 [SLOT_EXT_LEFT, SLOT_EXT_RIGHT, SLOT_EXT_BOT],
    FML_SQR_OFFSET:              [SLOT_OFFSET],
    FML_SQR_TRANSITION:          [SLOT_TRANSITION],
    FML_SQR_ENDCAP:              [SLOT_ENDCAP_SD],
    FML_SQR_TDF_ENDCAP:          [SLOT_ENDCAP_TDF],
    FML_SQR_CANVAS:              [SLOT_CANVAS],
    FML_SQR_BOOT_TAP:            [SLOT_TAP],
    FML_SQR_BOOT_TAP_W_DAMPER:   [SLOT_DAMPER_VOLUME],
    FML_SQR_CONICAL_TAP:         [SLOT_TAP],
    FML_SQR_CONICAL_TAP_W_DAMPER:[SLOT_DAMPER_VOLUME],
    FML_SQR_MAN_BAR:             [SLOT_MAN_BARS],

    # Round
    FML_RND_ELBOW_90_SR_STAMPED: [SLOT_DEGREE],
    FML_RND_GORED_ELBOW:         [SLOT_DEGREE],
    FML_RND_OFFSET:              [SLOT_OFFSET],
    FML_RND_REDUCER:             [SLOT_OFFSET],
    FML_RND_REDUCING_TEE:        [SLOT_EXT_LEFT, SLOT_EXT_RIGHT],
    FML_RND_ENDCAP:              [SLOT_ENDCAP_SD],
    FML_RND_DAMPER_VOLUME:       [SLOT_DAMPER_VOLUME],
    FML_RND_DAMPER:              [SLOT_DAMPER_FIRE],
    FML_RND_CANVAS:              [SLOT_CANVAS],
    FML_RND_BOOT_SADDLE:         [SLOT_TAP],

    # Oval
    FML_OVL_ELBOW_EASY:          [SLOT_DEGREE],
    FML_OVL_ELBOW_HARD:          [SLOT_DEGREE],
    FML_OVL_REDUCER:             [SLOT_OFFSET],
    FML_OVL_ENDCAP:              [SLOT_ENDCAP_SD],
    FML_OVL_CONICAL_TAP_HARD:    [SLOT_TAP],
}


_user_candidates = _load_user_candidate()
if isinstance(_user_candidates, dict):
    DEFAULT_TAG_SLOT_CANDIDATES = _user_candidates

DEFAULT_TAG_SKIP_PARAMETERS = {
    PYT_SKIP_TAG: ['skip'],
}

DEFAULT_NUMBER_SKIP_PARAMETERS = {
    PYT_SKIP_NUMBER: ['skip'],
}

DEFAULT_PARAMETER_HIERARCHY = [
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

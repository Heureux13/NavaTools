# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""
from config.tag_config import *


# fmt: off
# autopep8: off

def clean(v):
    return v.strip().lower()

# Square families
SQR_BIRD_SCREEN             = clean('Bird Screen')
SQR_BOOT_TAP                = clean('Boot Tap')
SQR_BOOT_TAP_W_DAMPER       = clean('Boot Tap - wDamper')
SQR_CANVAS                  = clean('Canvas')
SQR_CONICAL_TAP             = clean('Conical Tap')
SQR_CONICAL_TAP_W_DAMPER    = clean('Conical Tap - wDamper')
SQR_DAMPER                  = clean('Rect Volume Damper')
SQR_ELBOW_90                = clean('90° Elbow')
SQR_ELBOW_ADJ               = clean('Adjustable Elbow')
SQR_ELBOW_TEE               = clean('Tee')
SQR_ENDCAP_SND              = clean('End Cap')
SQR_ENDCAP_TDF              = clean('TDF End Cap')
SQR_MAN_BAR                 = clean('MAN BAR')
SQR_OFFSET                  = clean('Offset')
SQR_OGEE                    = clean('Ogee')
SQR_PANTS                   = clean('Pants')
SQR_RADIUS_CLUSTER          = clean('Radius Cluster')
SQR_RADIUS_DROP             = clean('Drop Cheek')
SQR_RADIUS_ELBOW            = clean('Radius Elbow')
SQR_STRAIGHT                = clean('Straight')
SQR_TO_RND                  = clean('Square to Ø')
SQR_TRANSITION              = clean('Transition')

# Round families
RND_3_WAY                   = clean('Three Way')
RND_3_WAY_BRANCH            = clean('3 Way Branch')
RND_BOOT_SADDLE_TAP         = clean('Boot Saddle Tap')
RND_BOOT_SADDLE             = clean('Boot Saddle')
RND_BOX_SADDLE              = clean('Box Saddle')
RND_CANVAS                  = clean('Ø Canvas')
RND_CHINA_CAP               = clean('China Cap')
RND_CONICAL_TEE             = clean('Conical Tee')
RND_COUPLING_MM             = clean('Coupling')
RND_COUPLING_FF             = clean('Coupling - Fitting')
RND_CROSS_TYPE_2            = clean('Cross (Type 2)')
RND_DAMPER                  = clean('Round Volume Damper')
RND_DAMPER_VOLUME           = clean('8inch Long Coupler wDamper')
RND_ELBOW_GORED             = clean('Gored Elbow')
RND_ELBOW_STAMPED           = clean('Elbow 90 SR - stamped')
RND_ENDCAP                  = clean('Ø End Cap')
RND_OFFSET                  = clean('Ø Offset')
RND_REDUCER_REDNECK         = clean('Redneck Reducer')
RND_REDUCER                 = clean('Reducer')
RND_REDUCING_TEE            = clean('Reducing Tee')
RND_ROOF_JACK               = clean('Roof Jack')
RND_SADDLE                  = clean('Saddle Tap')
RND_SADDLE_HANGER           = clean('Saddle Hanger')
RND_STRAIGHT                = clean('Spiral Duct')
RND_WYE                     = clean('WYE')

# Oval families
OVL_COUPLING                = clean('Oval Coupling')
OVL_CONICAL_TAP             = clean('Oval Conical Tap Hard')
OVL_ELBOW_EASY              = clean('Oval Elbow Easy')
OVL_ELBOW_HARD              = clean('Oval Elbow Hard')
OVL_ENDCAP                  = clean('Oval Cap End')
OVL_STRAIGHT                = clean('Oval Pipe (LP)')
OVL_REDUCER                 = clean('Oval Reducer')
OVL_TO_ROUND                = clean('Oval to Round')

# Devices families
DEV_DAMPER_FIRE_A           = clean('Fire Damper - Type A')
DEV_DAMPER_FIRE_B           = clean('Fire Damper - Type B')
DEV_DAMPER_FIRE_C           = clean('Fire Damper - Type C')
DEV_DAMPER_FIRE_CR          = clean('Fire Damper - Type CR')
DEV_DAMPER_FIRE_SMOKE_A     = clean('Smoke Fire Damper - Type A')
DEV_DAMPER_FIRE_SMOKE_B     = clean('Smoke yFire Damper - Type B')
DEV_ACCESS_DOOR             = clean('Access Door')
DEV_ACCESS_DOOR_RND         = clean('Rnd Access Door')

# Hangers
HNG_SQR_STRAP               = clean('Half Strap Hanger')
HNG_SQR_STRUT               = clean('Rectangular Strut Hanger')
HNG_SQR_DOUBLE_STRUT        = clean('Double Strut Hanger')
HNG_RND_STRAP               = clean('Round Strap Hanger')
HNG_RND_STRUT               = clean('Round Strut Hanger')

# Sleeves
SLV_WALL                    = clean('Rectangular Wall Sleeve')
SLV_FLOOR                   = clean('Rectangular Floor Sleeve')

# Other
HOUSE_PAD                   = clean('Housekeeping Pad')
ROOF_CURB                   = clean('Roof Curb')
# fmt: on
# autopep8: on

DUCT_FAMILY_TAG_SLOTS = {
    # Square
    SQR_ELBOW_90:            [SLOT_EXT_TOP, SLOT_EXT_BOT, SLOT_DEGREE],
    SQR_ELBOW_ADJ:           [SLOT_EXT_TOP, SLOT_EXT_BOT, SLOT_DEGREE],
    SQR_ELBOW_TEE:           [SLOT_EXT_LEFT, SLOT_EXT_RIGHT, SLOT_EXT_BOT],
    SQR_RADIUS_ELBOW:        [SLOT_DEGREE],
    SQR_OFFSET:              [SLOT_OFFSET],
    SQR_TRANSITION:          [SLOT_TRANSITION],
    SQR_ENDCAP_SND:          [SLOT_ENDCAP_SD],
    SQR_ENDCAP_TDF:          [SLOT_ENDCAP_TDF],
    SQR_CANVAS:              [SLOT_CANVAS],
    SQR_BOOT_TAP:            [SLOT_TAP],
    SQR_BOOT_TAP_W_DAMPER:   [SLOT_DAMPER_VOLUME],
    SQR_CONICAL_TAP:         [SLOT_TAP],
    SQR_CONICAL_TAP_W_DAMPER:[SLOT_DAMPER_VOLUME],
    SQR_MAN_BAR:             [SLOT_MAN_BARS],
    # Round
    RND_ELBOW_STAMPED:       [SLOT_DEGREE],
    RND_ELBOW_GORED:         [SLOT_DEGREE],
    RND_OFFSET:              [SLOT_OFFSET],
    RND_REDUCER:             [SLOT_OFFSET],
    RND_REDUCING_TEE:        [SLOT_EXT_LEFT, SLOT_EXT_RIGHT],
    RND_ENDCAP:              [SLOT_ENDCAP_SD],
    RND_DAMPER_VOLUME:       [SLOT_DAMPER_VOLUME],
    RND_DAMPER:              [SLOT_DAMPER_FIRE],
    RND_CANVAS:              [SLOT_CANVAS],
    RND_BOOT_SADDLE:         [SLOT_TAP],

    # Oval
    OVL_ELBOW_EASY:          [SLOT_DEGREE],
    OVL_ELBOW_HARD:          [SLOT_DEGREE],
    OVL_REDUCER:             [SLOT_OFFSET],
    OVL_ENDCAP:              [SLOT_ENDCAP_SD],
    OVL_CONICAL_TAP:         [SLOT_TAP],

    # Dampers
    DEV_DAMPER_FIRE_A:       [SLOT_DAMPER_FIRE],
    DEV_DAMPER_FIRE_B:       [SLOT_DAMPER_FIRE],
    DEV_DAMPER_FIRE_C:       [SLOT_DAMPER_FIRE],
    DEV_DAMPER_FIRE_CR:      [SLOT_DAMPER_FIRE],
    DEV_DAMPER_FIRE_SMOKE_A: [SLOT_DAMPER_FIRE],
    DEV_DAMPER_FIRE_SMOKE_B: [SLOT_DAMPER_FIRE],

    # Other
    DEV_ACCESS_DOOR:         [SLOT_ACCESS_PANEL],
    DEV_ACCESS_DOOR_RND:     [SLOT_ACCESS_PANEL],

    # Hangers
    HNG_SQR_STRAP:           [],
    HNG_SQR_STRUT:           [],
    HNG_SQR_DOUBLE_STRUT:    [],
    HNG_RND_STRAP:           [],
    HNG_RND_STRUT:           [],
}


STRAIGHT_FAMILIES = [
    RND_STRAIGHT,
    OVL_STRAIGHT,
    SQR_STRAIGHT,
 ]

ELBOW_FAMILIES = [
    SQR_ELBOW_90,
    SQR_ELBOW_ADJ,
    SQR_ELBOW_TEE,
    SQR_RADIUS_CLUSTER,
    SQR_RADIUS_ELBOW,
    RND_CONICAL_TEE,
    RND_ELBOW_STAMPED,
    RND_ELBOW_GORED,
    RND_REDUCING_TEE,
    OVL_ELBOW_EASY,
    OVL_ELBOW_HARD,
]

TAP_FAMILIES = [
    SQR_BOOT_TAP,
    SQR_BOOT_TAP_W_DAMPER,
    SQR_CONICAL_TAP,
    SQR_CONICAL_TAP_W_DAMPER,
    RND_BOOT_SADDLE,
]

DAMPER_FAMILIES = [
    DEV_DAMPER_FIRE_A,
    DEV_DAMPER_FIRE_B,
    DEV_DAMPER_FIRE_C,
    DEV_DAMPER_FIRE_CR,
    DEV_DAMPER_FIRE_SMOKE_A,
    DEV_DAMPER_FIRE_SMOKE_B,
    DEV_ACCESS_DOOR
]

FITTING_FAMILIES = [
    # Square
    SQR_ELBOW_90,
    SQR_ELBOW_ADJ,
    SQR_ELBOW_TEE,
    SQR_BIRD_SCREEN,
    SQR_BOOT_TAP,
    SQR_BOOT_TAP_W_DAMPER,
    SQR_CANVAS,
    SQR_CONICAL_TAP,
    SQR_CONICAL_TAP_W_DAMPER,
    SQR_RADIUS_DROP,
    SQR_ENDCAP_SND,
    SQR_OFFSET,
    SQR_OGEE,
    SQR_PANTS,
    SQR_RADIUS_CLUSTER,
    SQR_RADIUS_ELBOW,
    SQR_ENDCAP_TDF,
    SQR_TO_RND,
    SQR_TRANSITION,
    # Round
    RND_BOOT_SADDLE,
    RND_3_WAY,
    RND_3_WAY_BRANCH,
    RND_BOX_SADDLE,
    RND_CANVAS,
    RND_CHINA_CAP,
    RND_CONICAL_TEE,
    RND_CROSS_TYPE_2,
    RND_ELBOW_STAMPED,
    RND_ENDCAP,
    RND_ELBOW_GORED,
    RND_OFFSET,
    RND_REDUCER_REDNECK,
    RND_REDUCER,
    RND_REDUCING_TEE,
    RND_ROOF_JACK,
    RND_SADDLE,
    RND_SADDLE_HANGER,
    RND_WYE,
    # Oval
    OVL_CONICAL_TAP,
    OVL_ELBOW_EASY,
    OVL_ELBOW_HARD,
    OVL_ENDCAP,
    OVL_REDUCER,
    OVL_TO_ROUND,
]

HANGER_FAMILIES = [
    HNG_SQR_STRAP,
    HNG_SQR_STRUT,
    HNG_SQR_DOUBLE_STRUT,
    HNG_RND_STRAP,
    HNG_RND_STRUT,
]

ACCESS_DOOR_FAMILIES = [
    DEV_ACCESS_DOOR,
    DEV_ACCESS_DOOR_RND,
]

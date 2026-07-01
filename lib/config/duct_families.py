# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""
# fmt: off
# autopep8: off

# Square families
FML_SQR_STRAIGHT                = 'Straight'
FML_SQR_90_ELBOW                = '90° Elbow'
FML_SQR_ADJUSTABLE_ELBOW        = 'Adjustable Elbow'
FML_SQR_BIRD_SCREEN             = 'Bird Screen'
FML_SQR_BOOT_SADDLE             = 'Boot Saddle'
FML_SQR_BOOT_TAP                = 'Boot Tap'
FML_SQR_CANVAS                  = 'Canvas'
FML_SQR_DROP_CHEEK              = 'Drop Cheek'
FML_SQR_END_CAP                 = 'End Cap'
FML_SQR_TDF_END_CAP             = 'TDF End Cap'
FML_SQR_MAN_BAR                 = 'MAN BAR'
FML_SQR_OFFSET                  = 'Offset'
FML_SQR_OGEE                    = 'Ogee'
FML_SQR_PANTS                   = 'Pants'
FML_SQR_RADIUS_CLUSTER          = 'Radius Cluster'
FML_SQR_RADIUS_ELBOW            = 'Radius Elbow'
FML_SQR_ROOF_CURB               = 'Roof Curb'
FML_SQR_TO_RND                  = 'Square to Ø'
FML_SQR_TRANSITION              = 'Transition'
FML_SQR_TEE                     = 'Tee'
FML_SQR_BOOT_TAP_W_DAMPER       = 'Boot tap - wDamper'
FML_SQR_CONICAL_TAP             = 'Conical Tap'
FML_SQR_CONICAL_TAP_W_DAMPER    = 'Conical Tap - wDamper'

# Round families
FML_RND_SPIRAL_DUCT             = 'Spiral Duct'
FML_RND_ELBOW_90_SR_STAMPED     = 'Elbow 90 SR - stamped'
FML_RND_GORED_ELBOW             = 'Gored Elbow'
FML_RND_COUPLER_8IN_W_DAMPER    = '8inch Long coupler wDamper'
FML_RND_BOOT_SADDLE_LOWER       = 'boot saddle'
FML_RND_REDUCER                 = 'Reducer'
FML_RND_DAMPER                  = 'Round Damper'
FML_RND_3_WAY_BRANCH            = '3 Way Branch'
FML_RND_BOX_SADDLE              = 'Box Saddle'
FML_RND_CHINA_CAP               = 'China Cap'
FML_RND_CONICAL_TEE             = 'Conical Tee'
FML_RND_COUPLING                = 'Coupling'
FML_RND_CROSS_TYPE_2            = 'Cross (Type 2)'
FML_RND_ROOF_JACK               = 'Roof Jack'
FML_RND_ROUND_PIPE              = 'Round Pipe'
FML_RND_WYE                     = 'WYE'
FML_RND_COUPLING_FITTING        = 'Coupling - Fitting'

# Oval families
FML_OVL_PIPE_LP                 = 'Oval Pipe (LP)'
FML_OVL_COUPLING                = 'Oval Coupling'
FML_OVL_REDUCER                 = 'Oval Reducer'
FML_OVL_ELBOW_HARD              = 'Oval Elbow Hard'
FML_OVL_ELBOW_EASY              = 'Oval Elbow Easy'
FML_OVL_CONICAL_TAP_HARD        = 'Oval Conical Tap Hard'
FML_OVL_TO_ROUND                = 'Oval to Round'
FML_OVL_END_CAP                 = 'Oval Cap End'

# fmt: on
# autopep8: on


duct_families = {
    FML_SQR_STRAIGHT: 'Square',
    FML_SQR_90_ELBOW: 'Square',
    FML_SQR_ADJUSTABLE_ELBOW: 'Square',
    FML_SQR_BIRD_SCREEN: 'Square',
    FML_SQR_BOOT_SADDLE: 'Square',
    FML_SQR_BOOT_TAP: 'Square',
    FML_SQR_CANVAS: 'Round',
    FML_SQR_DROP_CHEEK: 'Square',
    FML_SQR_END_CAP: 'Round',
    FML_SQR_TDF_END_CAP: 'Square',
    FML_SQR_MAN_BAR: 'Square',
    FML_SQR_OFFSET: 'Round',
    FML_SQR_OGEE: 'Square',
    FML_SQR_PANTS: 'Square',
    FML_SQR_RADIUS_CLUSTER: 'Square',
    FML_SQR_RADIUS_ELBOW: 'Square',
    FML_SQR_ROOF_CURB: 'Square',
    FML_SQR_SQUARE_TO_DIA: 'Square',
    FML_SQR_TRANSITION: 'Square',
    FML_SQR_TEE: 'Square',
    FML_SQR_BOOT_TAP_W_DAMPER: 'Square',
    FML_SQR_CONICAL_TAP: 'Square',
    FML_SQR_CONICAL_TAP_W_DAMPER: 'Square',
    FML_RND_SPIRAL_DUCT: 'Round',
    FML_RND_ELBOW_90_SR_STAMPED: 'Round',
    FML_RND_GORED_ELBOW: 'Round',
    FML_RND_COUPLER_8IN_W_DAMPER: 'Round',
    FML_RND_BOOT_SADDLE_LOWER: 'Round',
    FML_RND_REDUCER: 'Round',
    FML_RND_DAMPER: 'Round',
    FML_RND_3_WAY_BRANCH: 'Round',
    FML_RND_BOX_SADDLE: 'Round',
    FML_RND_CHINA_CAP: 'Round',
    FML_RND_CONICAL_TEE: 'Round',
    FML_RND_COUPLING: 'Round',
    FML_RND_CROSS_TYPE_2: 'Round',
    FML_RND_ROOF_JACK: 'Round',
    FML_RND_ROUND_PIPE: 'Round',
    FML_RND_WYE: 'Round',
    FML_RND_COUPLING_FITTING: 'Round',
    FML_OVL_PIPE_LP: 'Oval',
    FML_OVL_COUPLING: 'Oval',
    FML_OVL_REDUCER: 'Oval',
    FML_OVL_ELBOW_HARD: 'Oval',
    FML_OVL_ELBOW_EASY: 'Oval',
    FML_OVL_CONICAL_TAP_HARD: 'Oval',
    FML_OVL_TO_ROUND: 'Oval',
    FML_OVL_CAP_END: 'Oval',
}

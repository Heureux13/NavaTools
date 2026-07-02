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
FML_SQR_90_ELBOW                = '90° Elbow'
FML_SQR_ADJUSTABLE_ELBOW        = 'Adjustable Elbow'
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
FML_RND_SPIRAL_DUCT             = 'Spiral Duct'
FML_RND_WYE                     = 'WYE'

# Oval families
FML_OVL_COUPLING                = 'Oval Coupling'
FML_OVL_CONICAL_TAP_HARD        = 'Oval Conical Tap Hard'
FML_OVL_ELBOW_EASY              = 'Oval Elbow Easy'
FML_OVL_ELBOW_HARD              = 'Oval Elbow Hard'
FML_OVL_ENDCAP                  = 'Oval Cap End'
FML_OVL_PIPE_LP                 = 'Oval Pipe (LP)'
FML_OVL_REDUCER                 = 'Oval Reducer'
FML_OVL_TO_ROUND                = 'Oval to Round'

# fmt: on
# autopep8: on


all_duct_families = {
    # Square
    FML_SQR_90_ELBOW: 'Square',
    FML_SQR_ADJUSTABLE_ELBOW: 'Square',
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
    FML_RND_SPIRAL_DUCT: 'Round',
    FML_RND_WYE: 'Round',
    # Oval
    FML_OVL_ENDCAP: 'Oval',
    FML_OVL_CONICAL_TAP_HARD: 'Oval',
    FML_OVL_COUPLING: 'Oval',
    FML_OVL_ELBOW_EASY: 'Oval',
    FML_OVL_ELBOW_HARD: 'Oval',
    FML_OVL_PIPE_LP: 'Oval',
    FML_OVL_REDUCER: 'Oval',
    FML_OVL_TO_ROUND: 'Oval',
}

fittings_families = {
    FML_SQR_90_ELBOW
}

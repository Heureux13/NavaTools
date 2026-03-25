# Revit Parameters

## Ductwork
| Parameter                 | Type      | location              |
|---------------------------|-----------|-----------------------|
| _duct_#                   | text      | segments and fittings |
| _aspect_ratio             | text      | model properties      |
| _duct_label               | text      | segments and fittings |
| _offset_bottom            | number    | model properties      |
| _offset_center_h          | number    | model properties      |
| _offset_center_v          | number    | model properties      |
| _offset_left              | number    | model properties      |
| _offset_right             | number    | model properties      |
| _offset_top               | number    | model properties      |
| _duct_run_weight          | number    | segments and fittings |
| _offset               | text      | segments and fittings |

## Equipment
| Parameter                 | Type      | location              |
|---------------------------|-----------|-----------------------|
| _equi_damper              | text      | segments and fittings |
| _equi_filter_1            | text      | model properties      |
| _equi_filter_2            | text      | model properties      |
| _equi_filter_3            | text      | model properties      |
| _equi_filter_4            | text      | model properties      |
| _equi_filter_5            | text      | model properties      |
| _equi_filter_6            | text      | model properties      |
| _equi_flue                | text      | segments and fittings |
| _equi_handing             | text      | segments and fittings |
| _equi_make                | text      | segments and fittings |
| _equi_model               | text      | segments and fittings |
| _equi_mount               | text      | segments and fittings |
| _equi_open_ea             | text      | model properties      |
| _equi_open_oa             | text      | model properties      |
| _equi_open_ra             | text      | model properties      |
| _equi_open_sa             | text      | model properties      |
| _equi_type                | text      | segments and fittings |
| _equi_v_ph                | text      | segments and fittings |

## Hangers
| Parameter                 | Type      | location              |
|---------------------------|-----------|-----------------------|
| _hang_weight_support      | number    | segments and fittings |

## Air Terminals
| Parameter                 | Type      | location              |
|---------------------------|-----------|-----------------------|
| _air_#                    | text      | segments and fittings |
| _air_cfm                  | text      | segments and fittings |
| _air_color                | text      | segments and fittings |
| _air_damper               | text      | segments and fittings |
| _air_make                 | text      | segments and fittings |
| _air_material             | text      | segments and fittings |
| _air_model                | text      | segments and fittings |
| _air_mount                | text      | segments and fittings |
| _air_size                 | text      | segments and fittings |
| _air_type                 | text      | segments and fittings |

## Accessories
| Parameter                 | Type      | location              |
|---------------------------|-----------|-----------------------|
| _acc_actuator             | text      | segments and fittings |
| _acc_make                 | text      | segments and fittings |
| _acc_model                | text      | segments and fittings |
| _acc_mount                | text      | segments and fittings |
| _acc_position             | text      | segments and fittings |
| _acc_sleeve               | length    | segments and fittings |
| _acc_type                 | text      | segments and fittings |
| _acc_v_ph                 | text      | segments and fittings |
| _acc_value_a              | length    | segments and fittings |
| _acc_value_k              | length    | segments and fittings |

## Total List
| Parameter             | MEP Fab Ductwork | Equipment | Hangers | Air Terminals | location              |
|-----------------------|------------------|-----------|---------|---------------|-----------------------|
| _#                    | x                |           |         | x             | segments and fittings |
| _actuator             | x                |           |         |               | segments and fittings |
| _aspect_ratio         | x                |           |         |               | model properties      |
| _cfm                  |                  |           |         | x             | segments and fittings |
| _color                |                  |           |         | x             | segments and fittings |
| _damper               |                  | x         |         | x             | segments and fittings |
| _filter_1             |                  | x         |         |               | model properties      |
| _filter_2             |                  | x         |         |               | model properties      |
| _filter_3             |                  | x         |         |               | model properties      |
| _filter_4             |                  | x         |         |               | model properties      |
| _filter_5             |                  | x         |         |               | model properties      |
| _filter_6             |                  | x         |         |               | model properties      |
| _flue                 |                  | x         |         |               | segments and fittings |
| _handing              |                  | x         |         |               | segments and fittings |
| _label                | x                | x         | x       | x             | segments and fittings |
| _make                 | x                | x         |         | x             | segments and fittings |
| _material             |                  |           |         | x             | segments and fittings |
| _model                | x                | x         |         | x             | segments and fittings |
| _mount                | x                | x         |         | x             | segments and fittings |
| _offset_bottom        | x                |           |         |               | model properties      |
| _offset_center_h      | x                |           |         |               | model properties      |
| _offset_center_v      | x                |           |         |               | model properties      |
| _offset_left          | x                |           |         |               | model properties      |
| _offset_right         | x                |           |         |               | model properties      |
| _offset_top           | x                |           |         |               | model properties      |
| _open_ea              |                  | x         |         |               | model properties      |
| _open_oa              |                  | x         |         |               | model properties      |
| _open_ra              |                  | x         |         |               | model properties      |
| _open_sa              |                  | x         |         |               | model properties      |
| _position             | x                |           |         |               | segments and fittings |
| _run_weight           | x                |           |         |               | segments and fittings |
| _size                 |                  |           |         | x             | segments and fittings |
| _skip                 | x                | x         | x       | x             | segments and fittings |
| _sleeve               | x                |           |         |               | segments and fittings |
| _offset               | x                |           |         |               | segments and fittings |
| _type                 | x                | x         |         | x             | segments and fittings |
| _v_ph                 | x                | x         |         |               | segments and fittings |
| _value_a              | x                |           |         |               | segments and fittings |
| _value_k              | x                |           |         |               | segments and fittings |
| _weight_support       |                  |           | x       |               | segments and fittings |
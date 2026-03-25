# Revit Parameters

## Ductwork
| Parameter                 | Type      | location              |
|---------------------------|-----------|-----------------------|
| _duct_#                   | text      | segments and fittings |
| _aspect_ratio             | text      | construction      |
| _duct_label               | text      | segments and fittings |
| _offset_bottom            | number    | construction      |
| _offset_center_h          | number    | construction      |
| _offset_center_v          | number    | construction      |
| _offset_left              | number    | construction      |
| _offset_right             | number    | construction      |
| _offset_top               | number    | construction      |
| _duct_run_weight          | number    | segments and fittings |
| _offset               | text      | segments and fittings |

## Equipment
| Parameter                 | Type      | location              |
|---------------------------|-----------|-----------------------|
| _equi_damper              | text      | segments and fittings |
| _equi_filter_1            | text      | construction      |
| _equi_filter_2            | text      | construction      |
| _equi_filter_3            | text      | construction      |
| _equi_filter_4            | text      | construction      |
| _equi_filter_5            | text      | construction      |
| _equi_filter_6            | text      | construction      |
| _equi_flue                | text      | segments and fittings |
| _equi_handing             | text      | segments and fittings |
| _equi_make                | text      | segments and fittings |
| _equi_model               | text      | segments and fittings |
| _equi_mount               | text      | segments and fittings |
| _equi_open_ea             | text      | construction      |
| _equi_open_oa             | text      | construction      |
| _equi_open_ra             | text      | construction      |
| _equi_open_sa             | text      | construction      |
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
| _aspect_ratio         | x                |           |         |               | construction      |
| _cfm                  |                  |           |         | x             | segments and fittings |
| _color                |                  |           |         | x             | segments and fittings |
| _damper               |                  | x         |         | x             | segments and fittings |
| _filter_1             |                  | x         |         |               | construction      |
| _filter_2             |                  | x         |         |               | construction      |
| _filter_3             |                  | x         |         |               | construction      |
| _filter_4             |                  | x         |         |               | construction      |
| _filter_5             |                  | x         |         |               | construction      |
| _filter_6             |                  | x         |         |               | construction      |
| _flue                 |                  | x         |         |               | segments and fittings |
| _handing              |                  | x         |         |               | segments and fittings |
| _label                | x                | x         | x       | x             | segments and fittings |
| _make                 | x                | x         |         | x             | segments and fittings |
| _material             |                  |           |         | x             | segments and fittings |
| _model                | x                | x         |         | x             | segments and fittings |
| _mount                | x                | x         |         | x             | segments and fittings |
| _offset_bottom        | x                |           |         |               | construction      |
| _offset_center_h      | x                |           |         |               | construction      |
| _offset_center_v      | x                |           |         |               | construction      |
| _offset_left          | x                |           |         |               | construction      |
| _offset_right         | x                |           |         |               | construction      |
| _offset_top           | x                |           |         |               | construction      |
| _open_ea              |                  | x         |         |               | construction      |
| _open_oa              |                  | x         |         |               | construction      |
| _open_ra              |                  | x         |         |               | construction      |
| _open_sa              |                  | x         |         |               | construction      |
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
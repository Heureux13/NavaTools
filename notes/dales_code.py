# Dynamo Python Node
# Revit Fabrication Duct Slip/Drive Connector Corrector
# Compatible target: Revit 2022 through Revit 2027
# Written with IronPython2-safe syntax so it can also run in CPython3 where the same API calls are available.
#
# IN[0] = Fabrication parts to check
# IN[1] = SandD connector name string, default "SandD"
# IN[2] = DandS connector name string, default "DandS"
# IN[3] = Dry run true/false, default True
# IN[4] = Square tolerance in feet, default 0.0001
# IN[5] = Process square connectors true/false, default False
# IN[6] = Optional max connector ID scan fallback, default 2000
#
# Rule:
# SandD puts Slip on the wider side.
# DandS puts Slip on the narrower side.
#
# Desired condition:
# Slip on top/bottom.
# Drives on vertical sides.
#
# For rectangular connectors:
# - If the physical top/bottom side is wider, use SandD.
# - If the physical top/bottom side is narrower, use DandS.
#
# For square connectors:
# - Width and height are equal, so wider/narrower cannot decide.
# - Fabrication fittings place the Slip on the side labeled Width.
# - If the Width axis is physically top/bottom, use SandD.
# - If the Width axis is physically vertical, use DandS.
#
# This script determines top/bottom orientation from connector coordinate systems,
# not from duct Width/Depth labels alone.

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from RevitServices.Transactions import TransactionManager
from RevitServices.Persistence import DocumentManager
import clr

clr.AddReference("RevitServices")

clr.AddReference("RevitAPIUI")

clr.AddReference("RevitAPI")

doc = DocumentManager.Instance.CurrentDBDocument

# ============================================================
# INPUTS
# ============================================================

try:
    raw_parts = UnwrapElement(IN[0])
except BaseException:
    raw_parts = IN[0]

try:
    SANDD_NAME = str(IN[1]).strip()
    if not SANDD_NAME:
        SANDD_NAME = "S&D"
except BaseException:
    SANDD_NAME = "S&D"

try:
    DANDS_NAME = str(IN[2]).strip()
    if not DANDS_NAME:
        DANDS_NAME = "D&S"
except BaseException:
    DANDS_NAME = "D&S"

try:
    DRY_RUN = bool(IN[3])
except BaseException:
    DRY_RUN = True

try:
    SQUARE_TOLERANCE = float(IN[4])
except BaseException:
    SQUARE_TOLERANCE = 0.0001

try:
    PROCESS_SQUARE = bool(IN[5])
except BaseException:
    PROCESS_SQUARE = False

try:
    MAX_CONNECTOR_ID_SCAN = int(IN[6])
    if MAX_CONNECTOR_ID_SCAN < 1:
        MAX_CONNECTOR_ID_SCAN = 2000
except BaseException:
    MAX_CONNECTOR_ID_SCAN = 2000


# ============================================================
# BASIC HELPERS
# ============================================================

def ensure_list(x):
    if x is None:
        return []
    if isinstance(x, list):
        return x
    return [x]


def flatten_list(value):
    result = []

    if value is None:
        return result

    if isinstance(value, list):
        for item in value:
            result.extend(flatten_list(item))
    else:
        result.append(value)

    return result


def eid(elem):
    try:
        return elem.Id.IntegerValue
    except BaseException:
        try:
            return int(elem.Id.Value)
        except BaseException:
            try:
                return str(elem.Id)
            except BaseException:
                return None


def safe_round(value, digits=6):
    try:
        return round(float(value), digits)
    except BaseException:
        return value


def normalize_name(value):
    try:
        text = str(value).strip().lower()
        text = text.replace('"', '')
        text = text.replace("'", "")
        text = text.replace(u"\u201c", "")
        text = text.replace(u"\u201d", "")
        text = text.replace(u"\u2018", "")
        text = text.replace(u"\u2019", "")
        text = text.replace(u"\u00a0", " ")
        while "  " in text:
            text = text.replace("  ", " ")
        return text.strip()
    except BaseException:
        return ""


def get_param_value(elem, name):
    try:
        p = elem.LookupParameter(name)
    except BaseException:
        return None

    if p is None:
        return None

    try:
        if p.StorageType == StorageType.String:
            return p.AsString()
        elif p.StorageType == StorageType.Double:
            return p.AsDouble()
        elif p.StorageType == StorageType.Integer:
            return p.AsInteger()
        else:
            return p.AsValueString()
    except BaseException:
        return None


def get_item_custom_id(elem):
    try:
        return str(elem.ItemCustomId)
    except BaseException:
        return None


def get_part_name(elem):
    try:
        return elem.Name
    except BaseException:
        return None


def get_connectors(elem):
    connectors = []

    try:
        cm = elem.ConnectorManager
        if cm:
            for c in cm.Connectors:
                connectors.append(c)
            return connectors
    except BaseException:
        pass

    try:
        mep = elem.MEPModel
        if mep:
            cm = mep.ConnectorManager
            if cm:
                for c in cm.Connectors:
                    connectors.append(c)
                return connectors
    except BaseException:
        pass

    return connectors


def abs_dot_z(vec):
    try:
        return abs(vec.Normalize().DotProduct(XYZ.BasisZ))
    except BaseException:
        return None


def get_connector_width_height(conn, fallback_width, fallback_depth):
    width = None
    height = None

    try:
        width = conn.Width
    except BaseException:
        width = None

    try:
        height = conn.Height
    except BaseException:
        height = None

    if width is None:
        width = fallback_width

    if height is None:
        height = fallback_depth

    return width, height


def get_fab_info(conn):
    try:
        info = conn.GetFabricationConnectorInfo()
        if info and info.IsValid():
            return info
    except BaseException:
        pass

    return None


def get_body_connector_id(conn):
    info = get_fab_info(conn)

    if info is None:
        return None

    try:
        return info.BodyConnectorId
    except BaseException:
        return None


def get_fabrication_index(conn):
    info = get_fab_info(conn)

    if info is None:
        return None

    try:
        return info.FabricationIndex
    except BaseException:
        return None


def connected_owner_ids(part, conn):
    ids = []

    try:
        refs = conn.AllRefs
    except BaseException:
        refs = []

    for ref in refs:
        try:
            owner = ref.Owner
            if owner and owner.Id != part.Id:
                ids.append(eid(owner))
        except BaseException:
            pass

    return ids


# ============================================================
# FABRICATION CONNECTOR NAME LOOKUP
# Revit 2022 through Revit 2027 compatibility layer
# ============================================================

def get_fabrication_config(doc):
    try:
        return FabricationConfiguration.GetFabricationConfiguration(doc)
    except BaseException:
        return None


def try_get_connector_name(config, connector_id):
    if config is None:
        return None

    try:
        name = config.GetFabricationConnectorName(int(connector_id))
        if name:
            return str(name).strip()
    except BaseException:
        pass

    return None


def get_connector_definition_ids_from_config(config):
    ids = []

    if config is None:
        return ids

    try:
        raw_ids = config.GetAllFabricationConnectorDefinitions(
            ConnectorDomainType.Undefined,
            ConnectorProfileType.Invalid
        )

        for cid in raw_ids:
            try:
                ids.append(int(cid))
            except BaseException:
                pass

        return ids
    except BaseException:
        pass

    try:
        raw_ids = config.GetAllFabricationConnectorDefinitions()

        for cid in raw_ids:
            try:
                ids.append(int(cid))
            except BaseException:
                pass

        return ids
    except BaseException:
        pass

    return ids


def build_connector_name_lookup(doc, max_scan_id):
    connector_lookup = {}
    id_to_name = {}
    duplicate_connector_names = {}
    lookup_report = []
    errors = []

    config = get_fabrication_config(doc)

    if config is None:
        errors.append("Could not get FabricationConfiguration from the current document")
        return connector_lookup, id_to_name, duplicate_connector_names, lookup_report, errors

    connector_ids = get_connector_definition_ids_from_config(config)

    if not connector_ids:
        errors.append("GetAllFabricationConnectorDefinitions did not return ids. Falling back to connector id scan.")

        try:
            scan_limit = int(max_scan_id)
        except BaseException:
            scan_limit = 2000

        if scan_limit < 1:
            scan_limit = 2000

        for cid in range(0, scan_limit + 1):
            name = try_get_connector_name(config, cid)

            if name:
                connector_ids.append(int(cid))

    seen_ids = set()

    for cid in connector_ids:
        try:
            cid_int = int(cid)
        except BaseException:
            continue

        if cid_int in seen_ids:
            continue

        seen_ids.add(cid_int)

        name = try_get_connector_name(config, cid_int)

        if not name:
            continue

        name_text = str(name).strip()

        if not name_text:
            continue

        key = normalize_name(name_text)

        id_to_name[cid_int] = name_text

        if key in connector_lookup:
            if key not in duplicate_connector_names:
                duplicate_connector_names[key] = [connector_lookup[key]]
            duplicate_connector_names[key].append(cid_int)
        else:
            connector_lookup[key] = cid_int

        lookup_report.append({
            "ConnectorId": cid_int,
            "ConnectorName": name_text
        })

    if len(connector_lookup) == 0:
        errors.append("No connector names were found from the active fabrication configuration")

    return connector_lookup, id_to_name, duplicate_connector_names, lookup_report, errors


def get_connector_id_by_name(connector_lookup, connector_name):
    key = normalize_name(connector_name)

    if key in connector_lookup:
        return connector_lookup[key]

    return None


def get_name_by_connector_id(id_to_name, connector_id):
    try:
        cid = int(connector_id)
    except BaseException:
        return None

    if cid in id_to_name:
        return id_to_name[cid]

    return None


def resolve_connector_id_from_report(lookup_report, connector_name):
    target_key = normalize_name(connector_name)

    for row in lookup_report:
        try:
            row_name = row["ConnectorName"]
            row_id = row["ConnectorId"]
        except BaseException:
            continue

        if normalize_name(row_name) == target_key:
            try:
                return int(row_id)
            except BaseException:
                return row_id

    return None


connector_lookup, connector_id_to_name, duplicate_connector_names, connector_lookup_full_report, connector_lookup_errors = build_connector_name_lookup(
    doc, MAX_CONNECTOR_ID_SCAN)

sandd_key = normalize_name(SANDD_NAME)
dands_key = normalize_name(DANDS_NAME)

SANDD_ID = resolve_connector_id_from_report(connector_lookup_full_report, SANDD_NAME)
DANDS_ID = resolve_connector_id_from_report(connector_lookup_full_report, DANDS_NAME)

if SANDD_ID is None:
    try:
        SANDD_ID = connector_lookup.get(sandd_key)
    except BaseException:
        SANDD_ID = None

if DANDS_ID is None:
    try:
        DANDS_ID = connector_lookup.get(dands_key)
    except BaseException:
        DANDS_ID = None

connector_lookup_report = {
    "SandD Name Input": SANDD_NAME,
    "DandS Name Input": DANDS_NAME,
    "SandD Key": sandd_key,
    "DandS Key": dands_key,
    "SandD Exists In Dictionary": sandd_key in connector_lookup,
    "DandS Exists In Dictionary": dands_key in connector_lookup,
    "Resolved SandD BodyConnectorId": SANDD_ID,
    "Resolved DandS BodyConnectorId": DANDS_ID,
    "Resolver Method": "Full Connector Lookup Report exact normalized name match, with dictionary fallback",
    "Connector Lookup Count": len(connector_lookup),
    "Duplicate Connector Names": duplicate_connector_names,
    "Lookup Errors": connector_lookup_errors,
    "Max Connector ID Scan": MAX_CONNECTOR_ID_SCAN,
    "Available Connector Names": sorted(connector_lookup.keys()),
    "Full Connector Lookup Report": connector_lookup_full_report
}

STOP_SCRIPT = False
STOP_REASONS = []

if SANDD_ID is None:
    STOP_SCRIPT = True
    STOP_REASONS.append("Could not resolve SandD connector name: " + str(SANDD_NAME))

if DANDS_ID is None:
    STOP_SCRIPT = True
    STOP_REASONS.append("Could not resolve DandS connector name: " + str(DANDS_NAME))


# ============================================================
# CONNECTION STORAGE HELPERS
# ============================================================

def is_external_connector(part, other_conn):
    try:
        if other_conn is None:
            return False
        if other_conn.Owner is None:
            return False
        if other_conn.Owner.Id == part.Id:
            return False
        return True
    except BaseException:
        return False


def connector_key(conn):
    try:
        owner_id = eid(conn.Owner)
    except BaseException:
        owner_id = "UnknownOwner"

    try:
        origin = conn.Origin
        ox = round(origin.X, 6)
        oy = round(origin.Y, 6)
        oz = round(origin.Z, 6)
    except BaseException:
        ox = oy = oz = 0

    try:
        fab_index = get_fabrication_index(conn)
    except BaseException:
        fab_index = None

    return str(owner_id) + "|" + str(fab_index) + "|" + str(ox) + "," + str(oy) + "," + str(oz)


def pair_key(c1, c2):
    k1 = connector_key(c1)
    k2 = connector_key(c2)

    if k1 <= k2:
        return k1 + " <-> " + k2

    return k2 + " <-> " + k1


def collect_connection_pairs_for_part(part):
    pairs = []
    seen = set()

    connectors = get_connectors(part)

    for c in connectors:
        try:
            if not c.IsConnected:
                continue
        except BaseException:
            continue

        try:
            refs = c.AllRefs
        except BaseException:
            refs = []

        for ref_conn in refs:
            if not is_external_connector(part, ref_conn):
                continue

            key = pair_key(c, ref_conn)

            if key in seen:
                continue

            seen.add(key)

            pairs.append({
                "PairKey": key,
                "PartId": eid(part),
                "PartConnector": c,
                "ExternalConnector": ref_conn,
                "ExternalOwnerId": eid(ref_conn.Owner)
            })

    return pairs


def disconnect_pair(pair):
    c1 = pair["PartConnector"]
    c2 = pair["ExternalConnector"]

    try:
        if c1.IsConnectedTo(c2):
            c1.DisconnectFrom(c2)
            return True, "Disconnected"

        return True, "Already disconnected"
    except Exception as ex:
        return False, str(ex)


def reconnect_pair(pair):
    c1 = pair["PartConnector"]
    c2 = pair["ExternalConnector"]

    try:
        if not c1.IsConnectedTo(c2):
            c1.ConnectTo(c2)
            return True, "Reconnected"

        return True, "Already connected"
    except Exception as ex:
        return False, str(ex)


# ============================================================
# ORIENTATION AND PROPOSAL LOGIC
# ============================================================

def blank_proposal(reason, bx_z=None, by_z=None, bz_z=None, is_square=None):
    return {
        "ProposedName": "Review",
        "ProposedBodyConnectorId": None,
        "Reason": reason,
        "TopBottomAxis": None,
        "VerticalSideAxis": None,
        "TopBottomDimension": None,
        "VerticalSideDimension": None,
        "BasisX_Dot_GlobalZ": bx_z,
        "BasisY_Dot_GlobalZ": by_z,
        "BasisZ_Dot_GlobalZ": bz_z,
        "IsSquare": is_square
    }


def proposal_result(
        proposed_name,
        proposed_id,
        reason,
        top_bottom_axis,
        vertical_axis,
        top_bottom_dim,
        vertical_dim,
        bx_z,
        by_z,
        bz_z,
        is_square
):
    return {
        "ProposedName": proposed_name,
        "ProposedBodyConnectorId": proposed_id,
        "Reason": reason,
        "TopBottomAxis": top_bottom_axis,
        "VerticalSideAxis": vertical_axis,
        "TopBottomDimension": top_bottom_dim,
        "VerticalSideDimension": vertical_dim,
        "BasisX_Dot_GlobalZ": bx_z,
        "BasisY_Dot_GlobalZ": by_z,
        "BasisZ_Dot_GlobalZ": bz_z,
        "IsSquare": is_square
    }


def propose_connector(conn, width, height):
    try:
        cs = conn.CoordinateSystem
        bx_z = abs_dot_z(cs.BasisX)
        by_z = abs_dot_z(cs.BasisY)
        bz_z = abs_dot_z(cs.BasisZ)
    except BaseException:
        return blank_proposal("Could not read connector coordinate system")

    if width is None or height is None:
        return blank_proposal(
            "Connector width or height unavailable",
            bx_z,
            by_z,
            bz_z
        )

    try:
        width = float(width)
        height = float(height)
    except BaseException:
        return blank_proposal(
            "Connector width or height could not be converted to float",
            bx_z,
            by_z,
            bz_z
        )

    is_square = abs(width - height) <= SQUARE_TOLERANCE

    if by_z is not None and bx_z is not None and by_z >= bx_z:
        top_bottom_axis = "X / Width"
        vertical_axis = "Y / Height"
        top_bottom_dim = width
        vertical_dim = height

    elif bx_z is not None and by_z is not None and bx_z > by_z:
        top_bottom_axis = "Y / Height"
        vertical_axis = "X / Width"
        top_bottom_dim = height
        vertical_dim = width

    else:
        return blank_proposal(
            "Could not determine vertical connector axis",
            bx_z,
            by_z,
            bz_z,
            is_square
        )

    if is_square:

        if not PROCESS_SQUARE:
            return proposal_result(
                "Skip Square",
                None,
                "Square connector skipped by safety setting",
                top_bottom_axis,
                vertical_axis,
                top_bottom_dim,
                vertical_dim,
                bx_z,
                by_z,
                bz_z,
                is_square
            )

        if top_bottom_axis == "X / Width":
            return proposal_result(
                SANDD_NAME,
                SANDD_ID,
                "Square connector: Width axis is top/bottom, so SandD keeps Slip on top/bottom",
                top_bottom_axis,
                vertical_axis,
                top_bottom_dim,
                vertical_dim,
                bx_z,
                by_z,
                bz_z,
                is_square
            )

        if vertical_axis == "X / Width":
            return proposal_result(
                DANDS_NAME,
                DANDS_ID,
                "Square connector: Width axis is vertical, so DandS moves Slip to top/bottom",
                top_bottom_axis,
                vertical_axis,
                top_bottom_dim,
                vertical_dim,
                bx_z,
                by_z,
                bz_z,
                is_square
            )

        return proposal_result(
            "Review",
            None,
            "Square connector: could not determine Width axis orientation",
            top_bottom_axis,
            vertical_axis,
            top_bottom_dim,
            vertical_dim,
            bx_z,
            by_z,
            bz_z,
            is_square
        )

    if top_bottom_dim > vertical_dim:
        return proposal_result(
            SANDD_NAME,
            SANDD_ID,
            "Top/bottom side is wider, so SandD puts Slip on top/bottom",
            top_bottom_axis,
            vertical_axis,
            top_bottom_dim,
            vertical_dim,
            bx_z,
            by_z,
            bz_z,
            is_square
        )

    if top_bottom_dim < vertical_dim:
        return proposal_result(
            DANDS_NAME,
            DANDS_ID,
            "Top/bottom side is narrower, so DandS puts Slip on top/bottom",
            top_bottom_axis,
            vertical_axis,
            top_bottom_dim,
            vertical_dim,
            bx_z,
            by_z,
            bz_z,
            is_square
        )

    return proposal_result(
        "Review",
        None,
        "Connector dimensions are equal but square logic did not handle the condition",
        top_bottom_axis,
        vertical_axis,
        top_bottom_dim,
        vertical_dim,
        bx_z,
        by_z,
        bz_z,
        is_square
    )


def set_body_connector_id(conn, new_id):
    info = get_fab_info(conn)

    if info is None:
        return False, "No valid FabricationConnectorInfo"

    try:
        if not info.IsValid():
            return False, "FabricationConnectorInfo is not valid"
    except BaseException:
        pass

    if new_id is None:
        return False, "No proposed connector id"

    try:
        current = info.BodyConnectorId
    except BaseException:
        current = None

    try:
        if current == int(new_id):
            return True, "Already correct"
    except BaseException:
        pass

    try:
        info.BodyConnectorId = int(new_id)
    except Exception as ex:
        return False, str(ex)

    try:
        verify = info.BodyConnectorId
        if int(verify) == int(new_id):
            return True, "Updated"
        else:
            return False, "Set attempted but value did not verify. Current value after set: " + str(verify)
    except Exception as ex:
        return True, "Updated, but verification failed: " + str(ex)


# ============================================================
# MAIN PROCESSING
# ============================================================

parts = flatten_list(raw_parts)

if STOP_SCRIPT:

    OUT = {
        "Mode": "Stopped",
        "Reason": "; ".join(STOP_REASONS),
        "Connector Lookup Report": connector_lookup_report,
        "Input Part Count": len(parts)
    }

else:

    reports = []
    targets_by_part_id = {}

    for part in parts:

        if part is None:
            continue

        part_id = eid(part)

        main_width = get_param_value(part, "Main Primary Width")
        main_depth = get_param_value(part, "Main Primary Depth")
        size = get_param_value(part, "Size")
        item_custom_id = get_item_custom_id(part)
        part_name = get_part_name(part)

        connectors = get_connectors(part)

        part_report = {
            "PartId": part_id,
            "PartName": part_name,
            "ItemCustomId": item_custom_id,
            "Size": size,
            "Main Primary Width": safe_round(main_width),
            "Main Primary Depth": safe_round(main_depth),
            "Connector Count": len(connectors),
            "Connector Reports": []
        }

        for index, conn in enumerate(connectors):

            width, height = get_connector_width_height(conn, main_width, main_depth)

            current_id = get_body_connector_id(conn)
            current_name = get_name_by_connector_id(connector_id_to_name, current_id)
            fab_index = get_fabrication_index(conn)

            proposal = propose_connector(conn, width, height)
            proposed_id = proposal["ProposedBodyConnectorId"]
            proposed_name = proposal["ProposedName"]

            needs_change = False
            skip_reason = ""

            try:
                current_id_int = int(current_id)

                if current_id_int not in [int(SANDD_ID), int(DANDS_ID)]:
                    needs_change = False
                    skip_reason = "Skipped: current connector is not one of the two target connector names"

                elif proposed_id is not None:
                    needs_change = current_id_int != int(proposed_id)

                    if needs_change:
                        skip_reason = "Will update"
                    else:
                        skip_reason = "Already correct"

                else:
                    needs_change = False
                    skip_reason = "Skipped: no proposed connector id"

            except Exception as ex:
                needs_change = False
                skip_reason = "Skipped because of exception: " + str(ex)

            try:
                is_connected = conn.IsConnected
            except BaseException:
                is_connected = None

            conn_report = {
                "ConnectorIndex": index,
                "FabricationIndex": fab_index,
                "CurrentBodyConnectorId": current_id,
                "CurrentConnectorName": current_name,
                "SandD Name Input": SANDD_NAME,
                "DandS Name Input": DANDS_NAME,
                "Resolved SandD BodyConnectorId": SANDD_ID,
                "Resolved DandS BodyConnectorId": DANDS_ID,
                "Proposed": proposed_name,
                "ProposedBodyConnectorId": proposed_id,
                "NeedsChange": needs_change,
                "SkipReason": skip_reason,
                "IsConnected": is_connected,
                "ConnectedOwnerIds": connected_owner_ids(part, conn),
                "Connector Width": safe_round(width),
                "Connector Height": safe_round(height),
                "TopBottomAxis": proposal["TopBottomAxis"],
                "VerticalSideAxis": proposal["VerticalSideAxis"],
                "TopBottomDimension": safe_round(proposal["TopBottomDimension"]),
                "VerticalSideDimension": safe_round(proposal["VerticalSideDimension"]),
                "BasisX_Dot_GlobalZ": safe_round(proposal["BasisX_Dot_GlobalZ"]),
                "BasisY_Dot_GlobalZ": safe_round(proposal["BasisY_Dot_GlobalZ"]),
                "BasisZ_Dot_GlobalZ": safe_round(proposal["BasisZ_Dot_GlobalZ"]),
                "IsSquare": proposal["IsSquare"],
                "Reason": proposal["Reason"]
            }

            part_report["Connector Reports"].append(conn_report)

            if needs_change:
                if part_id not in targets_by_part_id:
                    targets_by_part_id[part_id] = {
                        "Part": part,
                        "ConnectorChanges": []
                    }

                targets_by_part_id[part_id]["ConnectorChanges"].append({
                    "Connector": conn,
                    "ConnectorIndex": index,
                    "FabricationIndex": fab_index,
                    "CurrentBodyConnectorId": current_id,
                    "CurrentConnectorName": current_name,
                    "ProposedBodyConnectorId": proposed_id,
                    "ProposedName": proposed_name,
                    "Reason": proposal["Reason"]
                })

        reports.append(part_report)

    if DRY_RUN:

        OUT = {
            "Mode": "Dry Run",
            "Connector Lookup Report": connector_lookup_report,
            "SandD Name": SANDD_NAME,
            "DandS Name": DANDS_NAME,
            "SandD BodyConnectorId": SANDD_ID,
            "DandS BodyConnectorId": DANDS_ID,
            "Square Tolerance": SQUARE_TOLERANCE,
            "Process Square Connectors": PROCESS_SQUARE,
            "Input Part Count": len(parts),
            "Parts With Proposed Changes": len(targets_by_part_id),
            "Reports": reports
        }

    else:

        all_connection_pairs = []
        seen_pair_keys = set()
        connection_report = []

        for part_id, data in targets_by_part_id.items():

            part = data["Part"]
            pairs = collect_connection_pairs_for_part(part)
            stored_count = 0

            for pair in pairs:
                key = pair["PairKey"]

                if key in seen_pair_keys:
                    continue

                seen_pair_keys.add(key)
                all_connection_pairs.append(pair)
                stored_count += 1

            connection_report.append({
                "PartId": part_id,
                "StoredConnections": stored_count
            })

        disconnect_success = []
        disconnect_failed = []

        TransactionManager.Instance.EnsureInTransaction(doc)

        for pair in all_connection_pairs:
            ok, msg = disconnect_pair(pair)

            row = {
                "PartId": pair["PartId"],
                "ExternalOwnerId": pair["ExternalOwnerId"],
                "Message": msg
            }

            if ok:
                disconnect_success.append(row)
            else:
                disconnect_failed.append(row)

        TransactionManager.Instance.TransactionTaskDone()

        try:
            doc.Regenerate()
        except BaseException:
            pass

        update_success = []
        update_failed = []

        TransactionManager.Instance.EnsureInTransaction(doc)

        for part_id, data in targets_by_part_id.items():

            for change in data["ConnectorChanges"]:

                conn = change["Connector"]
                new_id = change["ProposedBodyConnectorId"]

                ok, msg = set_body_connector_id(conn, new_id)

                row = {
                    "PartId": part_id,
                    "ConnectorIndex": change["ConnectorIndex"],
                    "FabricationIndex": change["FabricationIndex"],
                    "FromBodyConnectorId": change["CurrentBodyConnectorId"],
                    "FromName": change["CurrentConnectorName"],
                    "ToBodyConnectorId": new_id,
                    "ToName": change["ProposedName"],
                    "Message": msg
                }

                if ok:
                    update_success.append(row)
                else:
                    update_failed.append(row)

        TransactionManager.Instance.TransactionTaskDone()

        try:
            doc.Regenerate()
        except BaseException:
            pass

        reconnect_success = []
        reconnect_failed = []

        TransactionManager.Instance.EnsureInTransaction(doc)

        for pair in all_connection_pairs:
            ok, msg = reconnect_pair(pair)

            row = {
                "PartId": pair["PartId"],
                "ExternalOwnerId": pair["ExternalOwnerId"],
                "Message": msg
            }

            if ok:
                reconnect_success.append(row)
            else:
                reconnect_failed.append(row)

        TransactionManager.Instance.TransactionTaskDone()

        try:
            doc.Regenerate()
        except BaseException:
            pass

        # ============================================================
        # COMPLETION POPUP
        # ============================================================

        sandd_to_dands = 0
        dands_to_sandd = 0
        other_connector_changes = 0

        for row in update_success:
            try:
                from_id = int(row["FromBodyConnectorId"])
                to_id = int(row["ToBodyConnectorId"])

                if from_id == int(SANDD_ID) and to_id == int(DANDS_ID):
                    sandd_to_dands += 1

                elif from_id == int(DANDS_ID) and to_id == int(SANDD_ID):
                    dands_to_sandd += 1

                else:
                    other_connector_changes += 1

            except BaseException:
                other_connector_changes += 1

        message = (
            "Slip/Drive Connector Correction Complete\n\n"
            + "Parts Selected: {}\n".format(len(parts))
            + "Parts Modified: {}\n".format(len(targets_by_part_id))
            + "Connectors Updated: {}\n".format(len(update_success))
            + "{} -> {}: {}\n".format(
                SANDD_NAME,
                DANDS_NAME,
                sandd_to_dands
            )
            + "{} -> {}: {}\n".format(
                DANDS_NAME,
                SANDD_NAME,
                dands_to_sandd
            )
            + "Other/Unclassified Changes: {}\n".format(
                other_connector_changes
            )
            + "Update Failures: {}\n".format(len(update_failed))
            + "Reconnect Failures: {}".format(len(reconnect_failed))
        )

        try:
            title = "SUCCESS"

            if (
                    len(update_failed) > 0
                    or len(reconnect_failed) > 0
                    or other_connector_changes > 0
            ):
                title = "WARNING"

            TaskDialog.Show(title, message)

        except BaseException:
            pass

        OUT = {
            "Mode": "Modify",
            "Connector Lookup Report": connector_lookup_report,
            "SandD Name": SANDD_NAME,
            "DandS Name": DANDS_NAME,
            "SandD BodyConnectorId": SANDD_ID,
            "DandS BodyConnectorId": DANDS_ID,
            "Square Tolerance": SQUARE_TOLERANCE,
            "Process Square Connectors": PROCESS_SQUARE,
            "Input Part Count": len(parts),
            "Parts With Proposed Changes": len(targets_by_part_id),
            "Stored Connection Count": len(all_connection_pairs),
            "Disconnected Count": len(disconnect_success),
            "Disconnect Failed Count": len(disconnect_failed),
            "Updated Connector Count": len(update_success),
            "Update Failed Count": len(update_failed),
            "Reconnected Count": len(reconnect_success),
            "Reconnect Failed Count": len(reconnect_failed),
            "Connection Report": connection_report,
            "Update Success": update_success,
            "Update Failures": update_failed,
            "Disconnect Failures": disconnect_failed,
            "Reconnect Failures": reconnect_failed,
            "Full Reports": reports
        }

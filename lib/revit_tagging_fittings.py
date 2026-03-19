# -*- coding: utf-8 -*-
import re
from revit_element import RevitElement
from revit_xyz import RevitXYZ
from Autodesk.Revit.DB import FilteredElementCollector, IndependentTag


class Fittings:
    """Tag fitting management for fabricated duct elements.

    Class attributes define the config — override them per project.
    Instantiate with the Revit context: Fittings(doc, view, tagger).
    """

    # ------------------------------------------------------------------
    # CONFIG — edit per project
    # ------------------------------------------------------------------

    TAG_SLOT_CANDIDATES = {
        'TM':        ["-FabDuct_TM_MV_Tag", "_umi_duct_ITEM_NUMBER"],
        'SIZE_FIX':  ["-FabDuct_SIZE_FIX_Tag", "_umi_duct_ITEM_NUMBER"],
        'EXT_IN':    ["-FabDuct_EXT IN_MV_Tag", "_umi_duct_ITEM_NUMBER"],
        'EXT_OUT':   ["-FabDuct_EXT OUT_MV_Tag", "_umi_duct_ITEM_NUMBER"],
        'EXT_LEFT':  ["-FabDuct_EXT LEFT_MV_Tag", "_umi_duct_ITEM_NUMBER"],
        'EXT_RIGHT': ["-FabDuct_EXT RIGHT_MV_Tag", "_umi_duct_ITEM_NUMBER"],
        'DEGREE':    ["-FabDuct_DEGREE_MV_Tag", "_umi_duct_ITEM_NUMBER"],
        'TRAN':      ["-FabDuct_TRAN_MV_Tag", "_umi_offset"],
        'MARK':      ["-FabDuct_MARK_Tag", "_umi_duct_ITEM_NUMBER"],
    }

    SLOT_TYPE_MARK = 'TM'
    SLOT_SIZE_FIX = 'SIZE_FIX'
    SLOT_EXT_IN = 'EXT_IN'
    SLOT_EXT_OUT = 'EXT_OUT'
    SLOT_EXT_LEFT = 'EXT_LEFT'
    SLOT_EXT_RIGHT = 'EXT_RIGHT'
    SLOT_DEGREE = 'DEGREE'
    SLOT_TRAN = 'TRAN'
    SLOT_MARK = 'MARK'

    skip_parameters = {
        'mark':             ['skip', 'skip n/a'],
        '_duct_tag_offset': ['skip', 'skip n/a'],
        '_duct_tag':        ['skip', 'skip n/a'],
    }

    parameter_hierarchy_to_check = ['mark', 'type mark']

    elbow_throat_allowances = {'tdf': 6, 's&d': 4}

    elbow_extension_tags = {
        '-FabDuct_EXT IN_MV_Tag',
        '-FabDuct_EXT OUT_MV_Tag',
        '-FabDuct_EXT LEFT_MV_Tag',
        '-FabDuct_EXT RIGHT_MV_Tag',
    }

    elbow_families = {
        'elbow',
        'elbow 90 degree',
        'tee',
        'elbow 90 sr - stamped',
    }

    square_elbow_families = {
        'elbow',
        'elbow 90 degree',
    }

    family_to_angle_skip = {
        'radius elbow',
        'gored elbow',
    }

    # ------------------------------------------------------------------
    # Init
    # ------------------------------------------------------------------

    def __init__(self, doc, view, tagger):
        self.doc = doc
        self.view = view
        self.tagger = tagger

        self.missing_tag_labels = set()
        self._slot_resolution_cache = {}

        # Pre-normalize rule sets once.
        self._norm_elbow_fam = {self._norm(x) for x in self.elbow_families}
        self._norm_square_elbow_fam = {self._norm(
            x) for x in self.square_elbow_families}
        self._norm_angle_skip_fam = {self._norm(
            x) for x in self.family_to_angle_skip}
        self._norm_ext_tags = {t.strip().lower()
                               for t in self.elbow_extension_tags}

        # Resolve all slots and build the family map once.
        for slot in self.TAG_SLOT_CANDIDATES:
            self._resolve_slot(slot)
        self.duct_families = self._build_duct_families()
        self.rect_damper_switch_families = self._build_rect_damper_switch_families()

    # ------------------------------------------------------------------
    # Tag resolution
    # ------------------------------------------------------------------

    @staticmethod
    def _norm(name):
        if not name:
            return ""
        return re.sub(r"[^a-z0-9]+", " ", str(name).strip().lower()).strip()

    def _resolve_slot(self, slot_name):
        """Return (tag, label_lower) for the first candidate found; cache the result."""
        key = str(slot_name or '').strip().upper()
        if key in self._slot_resolution_cache:
            return self._slot_resolution_cache[key]
        result = (None, None)
        seen = set()
        attempted_names = []
        for name in (self.TAG_SLOT_CANDIDATES.get(key) or []):
            if not name:
                continue
            k = str(name).strip()
            if not k or k.lower() in seen:
                continue
            seen.add(k.lower())
            attempted_names.append(k)
            try:
                tag = self.tagger.get_label(k)
            except LookupError:
                tag = None
            if tag is not None:
                result = (tag, k.strip().lower())
                break

        # Only report missing labels if no candidate in this slot could be resolved.
        if result[0] is None:
            self.missing_tag_labels.update(attempted_names)

        self._slot_resolution_cache[key] = result
        return result

    def _tag_cfg(self, *slot_names):
        """Build [(tag, 0.5), ...] from slot names, skipping unresolved slots."""
        cfg = []
        for slot_name in slot_names:
            tag, _ = self._resolve_slot(slot_name)
            if tag is not None:
                cfg.append((tag, 0.5))
        return cfg

    # ------------------------------------------------------------------
    # Family -> tag map
    # ------------------------------------------------------------------

    def _build_duct_families(self):
        """Return normalized family-name -> [(tag, position)] map."""
        s = self
        family_cfg = {
            "8inch long coupler wdamper":     s._tag_cfg(s.SLOT_TYPE_MARK),
            "conical tap - wdamper":          s._tag_cfg(s.SLOT_TYPE_MARK),
            # tag choice resolved at runtime
            "rect volume damper":             s._tag_cfg(s.SLOT_TYPE_MARK),
            "boot tap - wdamper":             s._tag_cfg(s.SLOT_TYPE_MARK),
            "access panel":                   s._tag_cfg(s.SLOT_TYPE_MARK),
            "cap":                            s._tag_cfg(s.SLOT_TYPE_MARK),
            "canvas":                         s._tag_cfg(s.SLOT_TYPE_MARK),
            "end cap":                        s._tag_cfg(s.SLOT_TYPE_MARK),
            "tdf end cap":                    s._tag_cfg(s.SLOT_TYPE_MARK),
            "drop cheek":                     s._tag_cfg(s.SLOT_SIZE_FIX),
            "radius bend":                    s._tag_cfg(s.SLOT_SIZE_FIX),
            "square bend":                    s._tag_cfg(s.SLOT_SIZE_FIX),
            "tap":                            s._tag_cfg(s.SLOT_SIZE_FIX),
            "elbow":                          s._tag_cfg(s.SLOT_EXT_IN, s.SLOT_EXT_OUT, s.SLOT_DEGREE),
            "elbow 90 degree":                s._tag_cfg(s.SLOT_EXT_IN, s.SLOT_EXT_OUT, s.SLOT_DEGREE),
            "gored elbow":                    s._tag_cfg(s.SLOT_DEGREE),
            "radius elbow":                   s._tag_cfg(s.SLOT_DEGREE),
            "tee":                            s._tag_cfg(s.SLOT_EXT_IN, s.SLOT_EXT_LEFT, s.SLOT_EXT_RIGHT),
            "mitred offset":                  s._tag_cfg(s.SLOT_TRAN),
            "cid330 - (radius 2-way offset)": s._tag_cfg(s.SLOT_TRAN),
            "offset":                         s._tag_cfg(s.SLOT_TRAN),
            "ogee":                           s._tag_cfg(s.SLOT_TRAN),
            "radius offset":                  s._tag_cfg(s.SLOT_TRAN),
            "reducer":                        s._tag_cfg(s.SLOT_TRAN),
            "square to ø":                    s._tag_cfg(s.SLOT_TRAN),
            "transition":                     s._tag_cfg(s.SLOT_TRAN),
            "fire damper - type b":           s._tag_cfg(s.SLOT_MARK),
            "manbars":                        s._tag_cfg(s.SLOT_MARK),
        }
        return {self._norm(k): v for k, v in family_cfg.items()}

    def _build_rect_damper_switch_families(self):
        """All candidate tag family names for the rect damper MARK/TM swap logic."""
        return {
            name.strip().lower()
            for slot in (self.SLOT_MARK, self.SLOT_TYPE_MARK)
            for name in (self.TAG_SLOT_CANDIDATES.get(slot) or [])
        }

    # ------------------------------------------------------------------
    # Skip rule helpers
    # ------------------------------------------------------------------

    def _tag_pool_text(self, tag):
        fam_name = (
            tag.Family.Name if tag and tag.Family else "").strip().lower()
        sym_name = (getattr(tag, "Name", "") or "").strip().lower()
        return (fam_name + " " + sym_name).strip()

    def _is_extension_tag(self, tag):
        pool = self._tag_pool_text(tag)
        return any(needle in pool for needle in self._norm_ext_tags)

    def _is_degree_tag(self, tag):
        return "-fabduct_degree_mv_tag" in self._tag_pool_text(tag)

    @staticmethod
    def _is_angle_close(raw_angle, target, tol=0.5):
        try:
            return abs(abs(float(raw_angle)) - float(target)) <= float(tol)
        except Exception:
            return False

    @staticmethod
    def _connector_dz(element):
        """Return vertical distance between the two furthest connectors (feet)."""
        try:
            origins = RevitXYZ(element).connector_origins()
            if origins and len(origins) >= 2:
                return abs(origins[1].Z - origins[0].Z)
        except Exception:
            pass
        return 0.0

    def should_skip_by_param(self, duct):
        for param, skip_values in self.skip_parameters.items():
            param_val = getattr(duct, param, None)
            if not param_val:
                param_candidates = [param, param.title(), param.upper()]
                param_val = None
                for candidate in param_candidates:
                    try:
                        param_val = RevitElement(
                            self.doc, self.view, duct.element).get_param(candidate)
                    except Exception:
                        param_val = None
                    if param_val is not None:
                        break
                if param_val is None:
                    try:
                        type_element = self.doc.GetElement(
                            duct.element.GetTypeId())
                    except Exception:
                        type_element = None
                    if type_element is not None:
                        for candidate in param_candidates:
                            try:
                                param_val = RevitElement(
                                    self.doc, self.view, type_element).get_param(candidate)
                            except Exception:
                                param_val = None
                            if param_val is not None:
                                break
            if param_val is None:
                continue
            if str(param_val).strip().lower() in [v.strip().lower() for v in skip_values]:
                return True
        return False

    def should_skip_tag(self, duct, tag):
        if self.should_skip_by_param(duct):
            return True

        fam = self._norm(duct.family)

        # Never place a degree tag on a 90° fitting regardless of orientation.
        if self._is_degree_tag(tag) and self._is_angle_close(duct.angle, 90.0):
            return True

        # Square elbows at 45°/90°: degree tag only on vertical elbows.
        is_45_or_90 = False
        if fam in self._norm_square_elbow_fam:
            try:
                ang = duct.angle
                if ang is not None:
                    if ang in [45, 90, 45.0, 90.0]:
                        is_45_or_90 = True
                    else:
                        ang_f = abs(float(ang))
                        if (44.5 <= ang_f <= 45.5) or (89.5 <= ang_f <= 90.5):
                            is_45_or_90 = True
            except Exception:
                pass
        if is_45_or_90 and self._is_degree_tag(tag):
            # skip horizontal; allow vertical
            return self._connector_dz(duct.element) <= 0.01

        # Extension tags: skip for any elbow with vertical movement.
        if fam in self._norm_elbow_fam and self._is_extension_tag(tag):
            if self._connector_dz(duct.element) > 0.01:
                return True

        # Radius/gored elbows at 45/90°: degree tag only on vertical elbows.
        if fam in self._norm_angle_skip_fam and duct.angle in [45, 90] and self._is_degree_tag(tag):
            # skip horizontal; allow vertical
            return self._connector_dz(duct.element) <= 0.01

        # Extension tags: skip when extension equals the required throat allowance.
        if fam in self._norm_elbow_fam and self._is_extension_tag(tag):
            connector_types = [
                duct.connector_0_type,
                duct.connector_1_type,
                getattr(duct, 'connector_2_type', None),
            ]
            for ctype in connector_types:
                if not ctype:
                    continue
                key = ctype.lower().strip()
                if key in {'slip & drive', 'standing s&d', 'standing s and d', 's and d'}:
                    key = 's&d'
                required_ext = self.elbow_throat_allowances.get(key)
                if required_ext is None:
                    continue
                ext_values = (
                    duct.extension_top,
                    duct.extension_bottom,
                    getattr(duct, 'extension_left', None),
                    getattr(duct, 'extension_right', None),
                )
                for ev in ext_values:
                    if isinstance(ev, (int, float)) and abs(ev - required_ext) <= 0.01:
                        return True

        return False

    # ------------------------------------------------------------------
    # Element / tagging helpers
    # ------------------------------------------------------------------

    def _param_value_from_element_or_type(self, element, param_name):
        """Return the string value of param_name from the element or its type, or None."""
        target = (param_name or '').strip().lower()
        if not target or element is None:
            return None
        elem_type = None
        try:
            elem_type = self.doc.GetElement(element.GetTypeId())
        except Exception:
            pass
        for owner in [element, elem_type]:
            if owner is None:
                continue
            try:
                for p in owner.Parameters:
                    try:
                        dname = p.Definition.Name if p and p.Definition else None
                        if not dname or dname.strip().lower() != target:
                            continue
                        val = p.AsString() or p.AsValueString()
                        if val:
                            return str(val).strip()
                    except Exception:
                        pass
            except Exception:
                pass
        return None

    def _rect_volume_damper_tag_choice(self, duct):
        """Return (tag, label_lower) — MARK tag if the element has a mark, else TM tag."""
        mark_param = self.parameter_hierarchy_to_check[0] if self.parameter_hierarchy_to_check else 'mark'
        type_mark_param = self.parameter_hierarchy_to_check[1] if len(
            self.parameter_hierarchy_to_check) > 1 else 'type mark'
        if self._param_value_from_element_or_type(duct.element, mark_param):
            return self._resolve_slot(self.SLOT_MARK)
        if self._param_value_from_element_or_type(duct.element, type_mark_param):
            return self._resolve_slot(self.SLOT_TYPE_MARK)
        return self._resolve_slot(self.SLOT_TYPE_MARK)

    def _delete_conflicting_tags_for_element(self, element, keep_family_name_lower, candidate_family_names_lower):
        """Remove tags on this element whose family is a candidate but differs from the keeper."""
        try:
            tags_in_view = (
                FilteredElementCollector(self.doc, self.view.Id)
                .OfClass(IndependentTag)
                .ToElements()
            )
        except Exception:
            return
        for t in tags_in_view:
            try:
                tagged_ids = t.GetTaggedLocalElementIds()
            except Exception:
                tagged_ids = []
            if not tagged_ids:
                continue

            def _as_int_id(revit_id_like):
                """Normalize ElementId/LinkElementId-like objects to int id values."""
                if revit_id_like is None:
                    return None

                # LinkElementId may wrap the host element id.
                host_id = getattr(revit_id_like, 'HostElementId', None)
                if host_id is not None:
                    revit_id_like = host_id

                for attr in ('Value', 'IntegerValue'):
                    try:
                        value = getattr(revit_id_like, attr)
                        if value is not None:
                            return int(value)
                    except Exception:
                        pass
                return None

            target_id = _as_int_id(element.Id)
            is_for_element = any(
                _as_int_id(tid) == target_id
                for tid in tagged_ids
            )
            if not is_for_element:
                continue
            try:
                t_type = self.doc.GetElement(t.GetTypeId())
                fam_name = (
                    t_type.Family.Name if t_type and t_type.Family else '').strip().lower()
            except Exception:
                fam_name = ''
            if fam_name in candidate_family_names_lower and fam_name != keep_family_name_lower:
                try:
                    self.doc.Delete(t.Id)
                except Exception:
                    pass

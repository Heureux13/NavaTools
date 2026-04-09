# -*- coding: utf-8 -*-
import re
from revit.revit_element import RevitElement
from ducts.revit_xyz import RevitXYZ
from tagging.tag_config import (
    SLOT_ACCESS_PANEL as CFG_SLOT_ACCESS_PANEL,
    SLOT_CANVAS as CFG_SLOT_CANVAS,
    SLOT_DAMPER_FIRE as CFG_SLOT_DAMPER_FIRE,
    SLOT_DAMPER_VOLUME as CFG_SLOT_DAMPER_VOLUME,
    SLOT_ENDCAP_SD as CFG_SLOT_ENDCAP_SD,
    SLOT_ENDCAP_TDF as CFG_SLOT_ENDCAP_TDF,
    SLOT_MAN_BARS as CFG_SLOT_MAN_BARS,
    SLOT_SIZE as CFG_SLOT_SIZE,
    SLOT_TAP as CFG_SLOT_TAP,
    SLOT_EXT_BOT as CFG_SLOT_EXT_BOT,
    SLOT_EXT_TOP as CFG_SLOT_EXT_TOP,
    SLOT_EXT_LEFT as CFG_SLOT_EXT_LEFT,
    SLOT_EXT_RIGHT as CFG_SLOT_EXT_RIGHT,
    SLOT_DEGREE as CFG_SLOT_DEGREE,
    SLOT_OFFSET as CFG_SLOT_OFFSET,
    SLOT_MARK as CFG_SLOT_MARK,
    DEFAULT_TAG_SLOT_CANDIDATES,
    DEFAULT_PARAMETER_HIERARCHY,
    DEFAULT_SKIP_PARAMETERS,
    WRITE_PARAMETER,
)
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
        slot: list(candidates)
        for slot, candidates in DEFAULT_TAG_SLOT_CANDIDATES.items()
    }

    SLOT_ACCESS_PANEL = CFG_SLOT_ACCESS_PANEL
    SLOT_CANVAS = CFG_SLOT_CANVAS
    SLOT_DAMPER_FIRE = CFG_SLOT_DAMPER_FIRE
    SLOT_DAMPER_VOLUME = CFG_SLOT_DAMPER_VOLUME
    SLOT_ENDCAP_SD = CFG_SLOT_ENDCAP_SD
    SLOT_ENDCAP_TDF = CFG_SLOT_ENDCAP_TDF
    SLOT_MAN_BARS = CFG_SLOT_MAN_BARS
    SLOT_SIZE = CFG_SLOT_SIZE
    SLOT_TAP = CFG_SLOT_TAP
    SLOT_EXT_BOT = CFG_SLOT_EXT_BOT
    SLOT_EXT_TOP = CFG_SLOT_EXT_TOP
    SLOT_EXT_LEFT = CFG_SLOT_EXT_LEFT
    SLOT_EXT_RIGHT = CFG_SLOT_EXT_RIGHT
    SLOT_DEGREE = CFG_SLOT_DEGREE
    SLOT_OFFSET = CFG_SLOT_OFFSET
    SLOT_MARK = CFG_SLOT_MARK

    skip_parameters = {
        param: list(values)
        for param, values in DEFAULT_SKIP_PARAMETERS.items()
    }
    parameter_hierarchy_to_check = list(DEFAULT_PARAMETER_HIERARCHY)
    write_parameter = WRITE_PARAMETER

    elbow_throat_allowances = {'tdf': 6, 's&d': 6}

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

        # Build extension/degree tag sets from resolved slot candidates.
        _ext_slots = (self.SLOT_EXT_BOT, self.SLOT_EXT_TOP, self.SLOT_EXT_LEFT, self.SLOT_EXT_RIGHT)
        self._norm_ext_tags = {
            self._candidate_pool_needle(name)
            for slot in _ext_slots
            for name in (self.TAG_SLOT_CANDIDATES.get(slot) or [])
            if self._candidate_pool_needle(name)
        }
        self._norm_degree_tags = {
            self._candidate_pool_needle(name)
            for name in (self.TAG_SLOT_CANDIDATES.get(self.SLOT_DEGREE) or [])
            if self._candidate_pool_needle(name)
        }
        self.duct_families = self._build_duct_families()

    # ------------------------------------------------------------------
    # Tag resolution
    # ------------------------------------------------------------------

    @staticmethod
    def _norm(name):
        if not name:
            return ""
        return re.sub(r"[^a-z0-9]+", " ", str(name).strip().lower()).strip()

    @staticmethod
    def _candidate_pool_needle(candidate):
        """Build a normalized search needle against 'family type' pools."""
        if isinstance(candidate, tuple):
            fam = str(candidate[0]).strip()
            typ = str(candidate[1]).strip()
            return "{} {}".format(fam, typ).strip().lower()
        return str(candidate).strip().lower()

    @staticmethod
    def _candidate_family_name(candidate):
        """Return the family-name portion of a candidate for family-only compares."""
        if isinstance(candidate, tuple):
            return str(candidate[0]).strip().lower()
        return str(candidate).strip().lower()

    @staticmethod
    def _get_param_case_insensitive(element, param_name):
        target = (param_name or '').strip().lower()
        if not target or element is None:
            return None
        try:
            for param in element.Parameters:
                try:
                    definition = param.Definition
                    name = definition.Name if definition else None
                    if name and name.strip().lower() == target:
                        return param
                except Exception:
                    pass
        except Exception:
            pass
        return None

    @staticmethod
    def _get_param_text(param):
        if not param:
            return ''
        try:
            value = param.AsString()
            if not value:
                value = param.AsValueString()
            return value.strip() if value else ''
        except Exception:
            return ''

    def _resolve_slot(self, slot_name):
        """Return (tag, label_lower) for the first candidate found; cache the result.

        Each candidate can be a plain string (substring-matched against the tag pool)
        or a (family, type) tuple for an exact family-name + type-name lookup.
        """
        key = str(slot_name or '').strip().upper()
        if key in self._slot_resolution_cache:
            return self._slot_resolution_cache[key]
        result = (None, None)
        seen = set()
        attempted_names = []
        for candidate in (self.TAG_SLOT_CANDIDATES.get(key) or []):
            if not candidate:
                continue
            if isinstance(candidate, tuple):
                fam = str(candidate[0]).strip()
                typ = str(candidate[1]).strip()
                dedup_key = "{}::{}".format(fam.lower(), typ.lower())
                label = "{}::{}".format(fam, typ)
            else:
                fam = None
                dedup_key = str(candidate).strip().lower()
                label = str(candidate).strip()
            if not dedup_key or dedup_key in seen:
                continue
            seen.add(dedup_key)
            attempted_names.append(label)
            try:
                if fam is not None:
                    tag = self.tagger.get_label_exact(fam, typ)
                else:
                    tag = self.tagger.get_label(label)
            except LookupError:
                tag = None
            if tag is not None:
                result = (tag, dedup_key)
                break

        # Only report missing labels if no candidate in this slot could be resolved.
        if result[0] is None:
            self.missing_tag_labels.update(attempted_names)

        self._slot_resolution_cache[key] = result
        return result

    def _get_hierarchy_value(self, element):
        """Return the last non-empty value from the configured hierarchy.

        Precedence is determined by order in parameter_hierarchy_to_check, so later
        parameters override earlier ones when they have a value.
        """
        if element is None:
            return ''

        elem_type = None
        try:
            elem_type = self.doc.GetElement(element.GetTypeId())
        except Exception:
            elem_type = None

        result = ''
        for param_name in self.parameter_hierarchy_to_check:
            value = self._get_param_text(
                self._get_param_case_insensitive(element, param_name)
            )
            if not value and elem_type is not None:
                value = self._get_param_text(
                    self._get_param_case_insensitive(elem_type, param_name)
                )
            if value:
                result = value
        return result

    def update_write_parameter_from_hierarchy(self, element):
        """Write the resolved hierarchy value to the configured output parameter."""
        value = self._get_hierarchy_value(element)
        if element is None or not self.write_parameter:
            return False, value
        updated = RevitElement(self.doc, self.view, element).set_param(
            self.write_parameter,
            value,
        )
        return updated, value

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
            # tag choice resolved at runtime
            "8inch long coupler wdamper": s._tag_cfg(s.SLOT_DAMPER_VOLUME),
            "conical tap - wdamper": s._tag_cfg(s.SLOT_DAMPER_VOLUME),
            "rect volume damper": s._tag_cfg(s.SLOT_DAMPER_VOLUME),
            "boot tap - wdamper": s._tag_cfg(s.SLOT_DAMPER_VOLUME),
            "access panel": s._tag_cfg(s.SLOT_ACCESS_PANEL),
            "cap": s._tag_cfg(s.SLOT_ENDCAP_SD),
            "canvas": s._tag_cfg(s.SLOT_CANVAS),
            "end cap": s._tag_cfg(s.SLOT_ENDCAP_SD),
            "tdf end cap": s._tag_cfg(s.SLOT_ENDCAP_TDF),
            "drop cheek": s._tag_cfg(s.SLOT_SIZE),
            "radius bend": s._tag_cfg(s.SLOT_SIZE),
            "square bend": s._tag_cfg(s.SLOT_SIZE),
            "tap": s._tag_cfg(s.SLOT_TAP),
            "elbow": s._tag_cfg(s.SLOT_EXT_BOT, s.SLOT_EXT_TOP, s.SLOT_DEGREE),
            "elbow 90 degree": s._tag_cfg(s.SLOT_EXT_BOT, s.SLOT_EXT_TOP, s.SLOT_DEGREE),
            "gored elbow": s._tag_cfg(s.SLOT_DEGREE),
            "radius elbow": s._tag_cfg(s.SLOT_DEGREE),
            "tee": s._tag_cfg(s.SLOT_EXT_BOT, s.SLOT_EXT_LEFT, s.SLOT_EXT_RIGHT),
            "mitred offset": s._tag_cfg(s.SLOT_OFFSET),
            "cid330 - (radius 2-way offset)": s._tag_cfg(s.SLOT_OFFSET),
            "offset": s._tag_cfg(s.SLOT_OFFSET),
            "ogee": s._tag_cfg(s.SLOT_OFFSET),
            "radius offset": s._tag_cfg(s.SLOT_OFFSET),
            "reducer": s._tag_cfg(s.SLOT_OFFSET),
            "square to ø": s._tag_cfg(s.SLOT_OFFSET),
            "transition": s._tag_cfg(s.SLOT_OFFSET),
            "fire damper - type b": s._tag_cfg(s.SLOT_DAMPER_FIRE),
            "manbars": s._tag_cfg(s.SLOT_MAN_BARS),
        }
        return {self._norm(k): v for k, v in family_cfg.items()}

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
        pool = self._tag_pool_text(tag)
        return any(needle in pool for needle in self._norm_degree_tags)

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

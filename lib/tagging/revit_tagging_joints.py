# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

import math
from ducts.revit_xyz import RevitXYZ
from ducts.revit_duct import JointSize
from config.tag_config import (
    SLOT_LENGTH as CFG_SLOT_LENGTH,
    SLOT_STACK as CFG_SLOT_STACK,
    DEFAULT_TAG_SLOT_CANDIDATES,
)


class Joints:
    """Tag joint management for fabricated duct elements.

    Class attributes define the config — override them per project.
    Instantiate with the Revit context: Joints(doc, view, tagger).
    """

    # ------------------------------------------------------------------
    # CONFIG — edit per project
    # ------------------------------------------------------------------

    # Slot configuration
    SLOT_LENGTH = CFG_SLOT_LENGTH
    TAG_SLOT_CANDIDATES = {
        slot: list(candidates)
        for slot, candidates in DEFAULT_TAG_SLOT_CANDIDATES.items()
    }

    # Allowed element families: name -> min_length_threshold (None = no threshold)
    ELEMENT_FAMILIES = {
        'straight': None,
        'spiral': 12,
        'spiral duct': 12,
    }

    # Skip parameters: param_name -> list of skip values
    SKIP_PARAMETERS = {
        'mark': ['skip', 'skip n/a'],
    }

    # Default threshold for INVALID joint sizes (in inches)
    DEFAULT_SHORT_THRESHOLD_IN = 56.0

    # Tagging constants
    PROGRESS_EVERY = 500
    BATCH_SIZE = 200

    # ------------------------------------------------------------------
    # Init
    # ------------------------------------------------------------------

    def __init__(self, doc, view, tagger):
        self.doc = doc
        self.view = view
        self.tagger = tagger
        # Build reverse mapping: tag_family_lower -> slot_name
        self._tag_family_to_slot = {}
        for slot, candidates in self.TAG_SLOT_CANDIDATES.items():
            for tag_family, _ in candidates:
                tag_family_lower = tag_family.strip().lower()
                if tag_family_lower not in self._tag_family_to_slot:
                    self._tag_family_to_slot[tag_family_lower] = []
                self._tag_family_to_slot[tag_family_lower].append(slot)

    # ------------------------------------------------------------------
    # Filtering methods
    # ------------------------------------------------------------------

    @staticmethod
    def should_skip_by_param(element, param_rules):
        """Check if element parameter matches skip rules.

        Returns:
            tuple: (should_skip, param_name, raw_value)
        """
        for param_name, skip_values in param_rules.items():
            param = element.LookupParameter(param_name)
            if not param:
                continue
            raw_val = None
            try:
                raw_val = param.AsString()
            except Exception:
                raw_val = None
            if not raw_val:
                try:
                    raw_val = param.AsValueString()
                except Exception:
                    raw_val = None
            if raw_val is None:
                continue
            val = raw_val.strip().lower()
            if val in {v.strip().lower() for v in skip_values}:
                return True, param_name, raw_val
        return False, None, None

    @staticmethod
    def passes_family_filter(duct, allowed_families):
        """Check if duct family is in allowed families."""
        fam = (duct.family or "").strip().lower()
        return fam in allowed_families

    @staticmethod
    def passes_length_threshold(duct, element_families):
        """Check if duct length passes minimum threshold for its family."""
        fam = (duct.family or "").strip().lower()
        min_length = element_families.get(fam)
        if min_length is not None and duct.length is not None:
            try:
                length_val = float(duct.length) if isinstance(
                    duct.length, str) else duct.length
                if isinstance(length_val, (int, float)) and length_val <= min_length:
                    return False
            except (ValueError, TypeError):
                pass
        return True

    @staticmethod
    def passes_joint_size_filter(duct, default_threshold):
        """Check if duct joint size matches SHORT or INVALID criteria."""
        joint_size = duct.joint_size
        if joint_size == JointSize.INVALID:
            if duct.length is None or duct.length > default_threshold:
                return False
        elif joint_size != JointSize.SHORT:
            return False
        return True

    def filter_ducts(self, ducts):
        """Filter ducts based on family, skip parameters, and joint size.

        Returns:
            tuple: (filtered_ducts, skipped_by_param)
        """
        filtered = []
        skipped = []

        for d in ducts:
            # Check family filter
            if not self.passes_family_filter(d, self.ELEMENT_FAMILIES):
                continue

            # Check skip parameters
            skip_param, skip_name, skip_val = self.should_skip_by_param(
                d.element, self.SKIP_PARAMETERS)
            if skip_param:
                skipped.append((d, skip_name, skip_val))
                continue

            # Check length threshold
            if not self.passes_length_threshold(d, self.ELEMENT_FAMILIES):
                continue

            # Check joint size
            if not self.passes_joint_size_filter(d, self.DEFAULT_SHORT_THRESHOLD_IN):
                continue

            filtered.append(d)

        return filtered, skipped

    # ------------------------------------------------------------------
    # Tagging methods
    # ------------------------------------------------------------------

    def place_tag_with_rotation(self, duct, tag_symbol, attempt_rotation=True):
        """Place tag on duct with optional rotation.

        Returns:
            placed_tag or None
        """
        # Get angle for rotation
        angle_rad = None
        if attempt_rotation:
            try:
                angle_deg = RevitXYZ(duct.element).straight_joint_degree()
                if isinstance(angle_deg, (int, float)):
                    angle_rad = math.radians(angle_deg)
            except Exception:
                pass

        # Try different location strategies
        loc = duct.element.Location
        placed_tag = None

        # Strategy 1: Point location
        if hasattr(loc, "Point") and loc.Point is not None:
            placed_tag = self.tagger.place_tag(
                duct.element, tag_symbol, loc.Point)
            if angle_rad is not None and placed_tag is not None:
                try:
                    placed_tag.Rotation = angle_rad
                except Exception:
                    pass
            return placed_tag

        # Strategy 2: Curve location (midpoint)
        if hasattr(loc, "Curve") and loc.Curve is not None:
            curve = loc.Curve
            midpoint = curve.Evaluate(0.5, True)
            placed_tag = self.tagger.place_tag(
                duct.element, tag_symbol, midpoint)
            if angle_rad is not None and placed_tag is not None:
                try:
                    placed_tag.Rotation = angle_rad
                except Exception:
                    pass
            return placed_tag

        # Strategy 3: Face facing view
        ref, centroid = self.tagger.get_face_facing_view(duct.element)
        if ref is not None and centroid is not None:
            placed_tag = self.tagger.place_tag(ref, tag_symbol, centroid)
            if angle_rad is not None and placed_tag is not None:
                try:
                    placed_tag.Rotation = angle_rad
                except Exception:
                    pass
            return placed_tag

        return None

    def is_tagged_with_slots(self, element, slots):
        """Check if element is already tagged with tags from any of the specified slots.

        This checks against existing tags on the element, mapping their families to slots.
        Works with any tag family name, not just those in the config, using pattern matching.

        Args:
            element: The Revit element to check
            slots: List of slot names to check (e.g., [SLOT_LENGTH, SLOT_STACK])

        Returns:
            bool: True if element is tagged with any family belonging to the target slots
        """
        # Get existing tag families on the element
        existing_families = self.tagger.get_existing_tag_families(element)
        target_slots = set(slots)

        # Check if any existing tag family belongs to a target slot
        for existing_family in existing_families:
            existing_family_lower = existing_family.strip().lower()

            # First check explicit mapping from config
            mapped_slots = self._tag_family_to_slot.get(
                existing_family_lower, [])
            if any(slot in target_slots for slot in mapped_slots):
                return True

            # Fallback: pattern-based matching for tag families not in config
            # Check if tag family name contains keywords for target slots
            for slot in target_slots:
                # Map slot to expected keywords in tag family name
                if slot == CFG_SLOT_LENGTH and 'length' in existing_family_lower:
                    return True
                if slot == CFG_SLOT_STACK and (
                        'stack' in existing_family_lower or 'sizestack' in existing_family_lower):
                    return True

        return False

    # ------------------------------------------------------------------
    # Output formatting
    # ------------------------------------------------------------------

    @staticmethod
    def format_newly_tagged(duct):
        """Format a duct for newly tagged output."""
        return (
            "No.{} | ID: {} | Fam: {} | Size: {} | Le: {:06.2f} | Ex: {}".format(
                "",  # Index added by caller
                duct.element.Id,
                duct.family,
                duct.size,
                duct.length if duct.length else 0.0,
                duct.extension_bottom if duct.extension_bottom else 0.0
            )
        )

    @staticmethod
    def format_already_tagged(duct):
        """Format a duct for already tagged output."""
        return (
            "Size: {} | Family: {} | Length: {:06.2f} | Element ID: {}".format(
                duct.size,
                duct.family,
                duct.length if duct.length else 0.0,
                duct.element.Id
            )
        )

    @staticmethod
    def format_skipped_by_param(duct, skip_name, skip_val):
        """Format a duct for skipped by parameter output."""
        return (
            "Param: {} | Value: {} | Family: {} | Length: {:06.2f} | Element ID: {}".format(
                skip_name,
                skip_val,
                duct.family,
                duct.length if duct.length else 0.0,
                duct.element.Id
            )
        )

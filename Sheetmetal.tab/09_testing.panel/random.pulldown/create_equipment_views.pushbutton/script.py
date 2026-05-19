# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import script, revit, forms, DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BoundingBoxXYZ,
    BuiltInCategory,
    ElementId,
    Category,
    View3D,
    ViewSection,
    ViewFamily,
    ViewFamilyType,
    XYZ,
    Transform,
    View,
)
from System.Collections.Generic import List
from System.Windows.Forms import (
    Form,
    Button,
    DialogResult,
    TextBox,
    TreeView,
    TreeNode,
    CheckBox,
    FormStartPosition,
)

# Button info
# ======================================================================
__title__ = 'Create Equipment Views'
__doc__ = '''
Select a Mechanical Equipment family and create optional 3D and section views
'''

# Main Script
# ======================================================================

doc = revit.doc
uidoc = revit.uidoc

# Keep output panel hidden unless explicitly enabled.
SHOW_LOG_WINDOW = False


class _SilentOutput(object):
    def print_md(self, message):
        pass


output = script.get_output() if SHOW_LOG_WINDOW else _SilentOutput()

templet_name_3d = 'FRANK 3D'
templet_name_section = '1-SHEETS - MD - ALL - JN - 3/8'
section_view = '-Working View - Frank'

# If True, use only templet_name_section and do not fall back to section type default template.
FORCE_NAMED_SECTION_TEMPLATE = True
ONE_INCH_FT = 1.0 / 12.0


def log_md(message):
    if output:
        output.print_md(message)


class EquipmentSelectionForm(Form):
    def __init__(self, family_names):
        Form.__init__(self)
        self.Text = "Select Mechanical Equipment"
        self.Width = 700
        self.Height = 630
        self.StartPosition = FormStartPosition.CenterScreen

        self.family_names = family_names
        self.hierarchy = self._build_hierarchy(family_names)
        self.checked_families = set()
        self._suppress_after_check = False

        self.search_box = TextBox()
        self.search_box.Top = 10
        self.search_box.Left = 10
        self.search_box.Width = 660
        self.search_box.TextChanged += self._filter_tree
        self.Controls.Add(self.search_box)

        self.tree_view = TreeView()
        self.tree_view.Top = 40
        self.tree_view.Left = 10
        self.tree_view.Width = 660
        self.tree_view.Height = 380
        self.tree_view.CheckBoxes = True
        self.tree_view.AfterCheck += self._on_node_checked
        self.Controls.Add(self.tree_view)

        self.create_north_checkbox = CheckBox()
        self.create_north_checkbox.Text = "Create horizontal section (North)"
        self.create_north_checkbox.Top = 430
        self.create_north_checkbox.Left = 10
        self.create_north_checkbox.Width = 320
        self.create_north_checkbox.Checked = False
        self.Controls.Add(self.create_north_checkbox)

        self.create_east_checkbox = CheckBox()
        self.create_east_checkbox.Text = "Create vertical section (East)"
        self.create_east_checkbox.Top = 455
        self.create_east_checkbox.Left = 10
        self.create_east_checkbox.Width = 320
        self.create_east_checkbox.Checked = False
        self.Controls.Add(self.create_east_checkbox)

        self.create_3d_checkbox = CheckBox()
        self.create_3d_checkbox.Text = "Create 3D view"
        self.create_3d_checkbox.Top = 480
        self.create_3d_checkbox.Left = 10
        self.create_3d_checkbox.Width = 320
        self.create_3d_checkbox.Checked = False

        self.Controls.Add(self.create_3d_checkbox)

        btn_all = Button()
        btn_all.Text = "Select All"
        btn_all.Top = 520
        btn_all.Left = 10
        btn_all.Width = 120
        btn_all.Click += self._on_select_all
        self.Controls.Add(btn_all)

        btn_select = Button()
        btn_select.Text = "Create Views"
        btn_select.Top = 520
        btn_select.Left = 140
        btn_select.Width = 140
        btn_select.DialogResult = DialogResult.Yes
        self.Controls.Add(btn_select)
        self.AcceptButton = btn_select

        btn_cancel = Button()
        btn_cancel.Text = "Cancel"
        btn_cancel.Top = 520
        btn_cancel.Left = 290
        btn_cancel.Width = 120
        btn_cancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(btn_cancel)
        self.CancelButton = btn_cancel

        self._build_tree()
        self.tree_view.CollapseAll()

    def _build_hierarchy(self, names):
        hierarchy = {}
        for full_name in names:
            base_name = full_name.split("(")[0].strip()
            if base_name not in hierarchy:
                hierarchy[base_name] = []
            hierarchy[base_name].append(full_name)
        return hierarchy

    def _build_tree(self, search_filter=None):
        self.tree_view.Nodes.Clear()

        for base_name in sorted(self.hierarchy.keys()):
            variants = self.hierarchy[base_name]
            base_matches = (not search_filter) or (
                search_filter in base_name.lower())
            variant_matches = [v for v in variants if (
                not search_filter) or (search_filter in v.lower())]

            if not base_matches and not variant_matches:
                continue

            parent_node = TreeNode(base_name)
            parent_node.Tag = ("parent", base_name)

            for variant in sorted(variants):
                if (not search_filter) or (search_filter in variant.lower()):
                    child_node = TreeNode(variant)
                    child_node.Tag = ("child", variant)
                    if variant in self.checked_families:
                        child_node.Checked = True
                    parent_node.Nodes.Add(child_node)

            for child_node in parent_node.Nodes:
                if child_node.Checked:
                    parent_node.Checked = True
                    break

            if parent_node.Nodes.Count > 0:
                self.tree_view.Nodes.Add(parent_node)

    def _filter_tree(self, sender, args):
        search = sender.Text.lower().strip()
        self._build_tree(search if search else None)
        self.tree_view.CollapseAll()

    def _on_select_all(self, sender, args):
        self._suppress_after_check = True
        try:
            for parent_node in self.tree_view.Nodes:
                parent_node.Checked = True
                for child_node in parent_node.Nodes:
                    child_node.Checked = True
                    if child_node.Tag and child_node.Tag[0] == "child":
                        self.checked_families.add(child_node.Tag[1])
        finally:
            self._suppress_after_check = False

    def _on_node_checked(self, sender, args):
        if self._suppress_after_check:
            return

        node = args.Node
        if not node or not node.Tag:
            return

        self._suppress_after_check = True
        try:
            kind, value = node.Tag

            if kind == "parent":
                for child_node in node.Nodes:
                    child_node.Checked = node.Checked
                    if child_node.Tag and child_node.Tag[0] == "child":
                        child_variant = child_node.Tag[1]
                        if node.Checked:
                            self.checked_families.add(child_variant)
                        else:
                            self.checked_families.discard(child_variant)

            elif kind == "child":
                if node.Checked:
                    self.checked_families.add(value)
                else:
                    self.checked_families.discard(value)
                    if node.Parent:
                        parent_checked = any(
                            c.Checked for c in node.Parent.Nodes)
                        node.Parent.Checked = parent_checked
        finally:
            self._suppress_after_check = False

    def get_selected_families(self):
        # Read from visible tree and merge with tracked checked state.
        selected = set(self.checked_families)
        for parent_node in self.tree_view.Nodes:
            for child_node in parent_node.Nodes:
                if child_node.Checked and child_node.Tag and child_node.Tag[0] == "child":
                    selected.add(child_node.Tag[1])
        return sorted(selected)

    def get_create_north_section(self):
        return self.create_north_checkbox.Checked

    def get_create_east_section(self):
        return self.create_east_checkbox.Checked

    def get_create_3d_view(self):
        return self.create_3d_checkbox.Checked


def get_family_name(elem):
    symbol = getattr(elem, "Symbol", None)
    if symbol and getattr(symbol, "Family", None):
        return symbol.Family.Name

    fam = getattr(elem, "Family", None)
    if fam:
        return fam.Name

    return None


def get_view_family_type(doc, target_view_family):
    for view_type in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if view_type.ViewFamily == target_view_family:
            return view_type
    return None


def get_view_family_type_by_name(doc, target_view_family, type_name):
    target_name = (type_name or "").strip().lower()
    for view_type in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        try:
            if view_type.ViewFamily != target_view_family:
                continue
            if (view_type.Name or "").strip().lower() == target_name:
                return view_type
        except BaseException:
            pass
    return None


def build_combined_bbox(elements):
    bbox_min = None
    bbox_max = None

    for elem in elements:
        try:
            bbox = elem.get_BoundingBox(None)
            if not bbox:
                continue

            if bbox_min is None:
                bbox_min = bbox.Min
                bbox_max = bbox.Max
            else:
                bbox_min = XYZ(
                    min(bbox_min.X, bbox.Min.X),
                    min(bbox_min.Y, bbox.Min.Y),
                    min(bbox_min.Z, bbox.Min.Z)
                )
                bbox_max = XYZ(
                    max(bbox_max.X, bbox.Max.X),
                    max(bbox_max.Y, bbox.Max.Y),
                    max(bbox_max.Z, bbox.Max.Z)
                )
        except BaseException:
            pass

    return bbox_min, bbox_max


def fit_bbox_to_padding_and_levels(bbox_min, bbox_max, xy_padding_ft=2.0):
    levels = list(FilteredElementCollector(doc).OfClass(DB.Level).ToElements())
    elevations = sorted([lvl.Elevation for lvl in levels])

    z_bottom = None
    z_top = None

    for elev in elevations:
        if elev <= bbox_min.Z:
            z_bottom = elev
        if z_top is None and elev >= bbox_max.Z:
            z_top = elev

    # Apply fallback independently for each missing side.
    if z_bottom is None:
        z_bottom = bbox_min.Z - 2.0
    if z_top is None:
        z_top = bbox_max.Z + 2.0
    if z_top <= z_bottom:
        z_bottom = bbox_min.Z - 2.0
        z_top = bbox_max.Z + 2.0

    return (
        XYZ(bbox_min.X - xy_padding_ft, bbox_min.Y - xy_padding_ft, z_bottom),
        XYZ(bbox_max.X + xy_padding_ft, bbox_max.Y + xy_padding_ft, z_top),
    )


def get_section_vertical_limits(bbox_min_z, bbox_max_z, offset_ft=ONE_INCH_FT):
    levels = list(FilteredElementCollector(doc).OfClass(DB.Level).ToElements())
    elevations = sorted([lvl.Elevation for lvl in levels])

    level_below = None
    level_above = None

    for elev in elevations:
        if elev <= bbox_min_z:
            level_below = elev
        if level_above is None and elev >= bbox_max_z:
            level_above = elev

    z_bottom = (
        level_below + offset_ft) if level_below is not None else (bbox_min_z - offset_ft)
    z_top = (level_above -
             offset_ft) if level_above is not None else (bbox_max_z + offset_ft)

    if z_top <= z_bottom:
        z_bottom = bbox_min_z - 1.0
        z_top = bbox_max_z + 1.0

    return z_bottom, z_top


def get_unique_view_name(base_name):
    existing = set()
    for view in FilteredElementCollector(doc).OfClass(View):
        try:
            if not view.IsTemplate:
                existing.add(view.Name)
        except BaseException:
            pass

    if base_name not in existing:
        return base_name

    i = 2
    while True:
        candidate = "{} ({})".format(base_name, i)
        if candidate not in existing:
            return candidate
        i += 1


def find_view_by_name(name):
    for view in FilteredElementCollector(doc).OfClass(View):
        try:
            if not view.IsTemplate and view.Name == name:
                return view
        except BaseException:
            pass
    return None


def get_mark_or_family(elem, fallback):
    try:
        p = elem.LookupParameter("Mark")
        if p:
            val = p.AsString()
            if val and val.strip():
                return val.strip()
    except BaseException:
        pass
    return fallback


def get_3d_view_template_by_name(name):
    for view in FilteredElementCollector(doc).OfClass(View3D):
        try:
            if view.IsTemplate and view.Name == name:
                return view
        except BaseException:
            pass
    return None


def get_view_templates_by_name(name):
    matches = []
    target_name = (name or "").strip().lower()
    for view in FilteredElementCollector(doc).OfClass(View):
        try:
            if view.IsTemplate and (view.Name or "").strip().lower() == target_name:
                matches.append(view)
        except BaseException:
            pass
    return matches


def get_default_template_from_view_type(view_family_type):
    if not view_family_type:
        return None

    try:
        default_template_id = getattr(
            view_family_type, "DefaultTemplateId", ElementId.InvalidElementId)
        if default_template_id and default_template_id != ElementId.InvalidElementId:
            template_view = doc.GetElement(default_template_id)
            if template_view and getattr(template_view, "IsTemplate", False):
                return template_view
    except BaseException:
        pass

    try:
        p = view_family_type.LookupParameter(
            "View Template applied to new views")
        if p:
            tid = p.AsElementId()
            if tid and tid != ElementId.InvalidElementId:
                template_view = doc.GetElement(tid)
                if template_view and getattr(template_view, "IsTemplate", False):
                    return template_view
    except BaseException:
        pass

    return None


def set_default_template_on_view_type(view_family_type, template_view):
    if not view_family_type or not template_view:
        return False

    try:
        if view_family_type.DefaultTemplateId != template_view.Id:
            view_family_type.DefaultTemplateId = template_view.Id
        return True
    except BaseException:
        pass

    try:
        p = view_family_type.LookupParameter(
            "View Template applied to new views")
        if p and not p.IsReadOnly:
            p.Set(template_view.Id)
            return True
    except BaseException:
        pass

    return False


def get_section_view_type_by_default_template_name(doc, template_name):
    target_name = (template_name or "").strip().lower()
    for view_type in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        try:
            if view_type.ViewFamily != ViewFamily.Section:
                continue
            default_template = get_default_template_from_view_type(view_type)
            if default_template and (default_template.Name or "").strip().lower() == target_name:
                return view_type
        except BaseException:
            pass
    return None


def apply_first_valid_template(target_view, template_candidates):
    for template_view in template_candidates:
        try:
            # Try direct assignment first; some view/type combos report validity inconsistently.
            target_view.ViewTemplateId = template_view.Id
            if target_view.ViewTemplateId == template_view.Id:
                return template_view
        except BaseException:
            pass

        try:
            # Second attempt: clear template then assign again.
            target_view.ViewTemplateId = ElementId.InvalidElementId
            target_view.ViewTemplateId = template_view.Id
            if target_view.ViewTemplateId == template_view.Id:
                return template_view
        except BaseException:
            pass
    return None


def apply_section_view_type(target_view, section_view_type):
    if not target_view or not section_view_type:
        return False

    try:
        if target_view.GetTypeId() != section_view_type.Id:
            target_view.ChangeTypeId(section_view_type.Id)
        return target_view.GetTypeId() == section_view_type.Id
    except BaseException:
        pass

    return False


def prepare_view_visibility(view, clear_template=True):
    # Remove template influence and ensure ME category is visible.
    if clear_template:
        view.ViewTemplateId = ElementId.InvalidElementId

    me_cat = Category.GetCategory(doc, BuiltInCategory.OST_MechanicalEquipment)
    if me_cat:
        try:
            if view.CanCategoryBeHidden(me_cat.Id):
                view.SetCategoryHidden(me_cat.Id, False)
        except BaseException:
            pass


try:
    output.print_md("## Mechanical Equipment View Builder")

    mep_equipment = (FilteredElementCollector(doc)
                     .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                     .WhereElementIsNotElementType()
                     .ToElements())

    families_dict = {}
    for elem in mep_equipment:
        family_name = get_family_name(elem)
        group_name = get_mark_or_family(
            elem, family_name if family_name else "(No Family)")
        if not group_name:
            group_name = "(Unnamed)"

        if group_name not in families_dict:
            families_dict[group_name] = []
        families_dict[group_name].append(elem)

    if not families_dict:
        output.print_md("No Mechanical Equipment elements found in document.")
        script.exit()

    # Sort family names
    family_names = sorted(families_dict.keys())

    # Show enhanced searchable picker (family + view options)
    selection_form = EquipmentSelectionForm(family_names)
    if selection_form.ShowDialog() != DialogResult.Yes:
        output.print_md("Selection canceled.")
        script.exit()

    selected_families = selection_form.get_selected_families()
    if not selected_families:
        output.print_md("No families selected.")
        script.exit()

    create_north_section = selection_form.get_create_north_section()
    create_east_section = selection_form.get_create_east_section()
    create_3d_view = selection_form.get_create_3d_view()

    if not (create_north_section or create_east_section or create_3d_view):
        output.print_md("No view types selected.")
        script.exit()

    view_family_3d = get_view_family_type(doc, ViewFamily.ThreeDimensional)
    view_family_section = get_view_family_type_by_name(
        doc, ViewFamily.Section, section_view)
    if not view_family_section:
        view_family_section = get_view_family_type(doc, ViewFamily.Section)

    if create_3d_view and not view_family_3d:
        output.print_md("Could not find 3D view family type.")
        script.exit()

    template_view = get_3d_view_template_by_name(templet_name_3d)
    section_template_candidates = []

    # Priority 1: explicit template name configured in script.
    named_section_templates = get_view_templates_by_name(templet_name_section)
    if FORCE_NAMED_SECTION_TEMPLATE and not named_section_templates:
        forms.alert("Section template '{}' was not found. Update templet_name_section or disable FORCE_NAMED_SECTION_TEMPLATE.".format(
            templet_name_section), title="Create Views Warning")
    for tpl in named_section_templates:
        already_added = any(
            existing.Id == tpl.Id for existing in section_template_candidates)
        if not already_added:
            section_template_candidates.append(tpl)

    # Hard enforcement for dependent section types:
    # 1) Prefer a section view type that already defaults to the named template.
    # 2) Otherwise try to rewrite the selected section type default template.
    if FORCE_NAMED_SECTION_TEMPLATE and view_family_section and section_template_candidates:
        preferred_section_type = get_section_view_type_by_default_template_name(
            doc, templet_name_section)
        if preferred_section_type:
            view_family_section = preferred_section_type
        else:
            with revit.Transaction("Set Section View Type Default Template"):
                set_ok = set_default_template_on_view_type(
                    view_family_section, section_template_candidates[0])
            if not set_ok:
                forms.alert(
                    "Could not set default template '{}' on section type '{}'. Revit may be locking this type/template relationship.".format(
                        templet_name_section,
                        view_family_section.Name if view_family_section else "<None>"),
                    title="Create Views Warning")

    # Priority 2 (optional): template attached to selected section view type.
    if not FORCE_NAMED_SECTION_TEMPLATE:
        section_type_default_template = get_default_template_from_view_type(
            view_family_section)
        if section_type_default_template:
            already_added = any(
                existing.Id == section_type_default_template.Id for existing in section_template_candidates)
            if not already_added:
                section_template_candidates.append(
                    section_type_default_template)

    all_element_ids = []
    total_section_views = 0
    last_3d_view = None

    for selected_family in selected_families:
        elements = families_dict.get(selected_family, [])
        element_ids = [e.Id for e in elements if e and e.Id]
        if not elements or not element_ids:
            continue

        all_element_ids.extend(element_ids)

        first_elem = elements[0]
        view_base_name = get_mark_or_family(first_elem, selected_family)

        bbox_min, bbox_max = build_combined_bbox(elements)
        if not (bbox_min and bbox_max):
            output.print_md(
                "Could not build a bounding box for family **{}**. Skipping.".format(selected_family))
            continue

        fitted_min, fitted_max = fit_bbox_to_padding_and_levels(
            bbox_min, bbox_max, 2.0)
        section_bottom_z, section_top_z = get_section_vertical_limits(
            bbox_min.Z, bbox_max.Z)
        desired_3d_name = "3D - {}".format(view_base_name)
        section_views = []
        new_3d_view = None

        with revit.Transaction("Create Mechanical Equipment Views - {}".format(selected_family)):
            if create_3d_view:
                existing_3d = find_view_by_name(desired_3d_name)
                if existing_3d:
                    new_3d_view = existing_3d
                    if template_view:
                        try:
                            if new_3d_view.IsValidViewTemplate(template_view.Id):
                                new_3d_view.ViewTemplateId = template_view.Id
                        except BaseException:
                            pass
                else:
                    new_3d_view = View3D.CreateIsometric(
                        doc, view_family_3d.Id)
                    new_3d_view.Name = desired_3d_name

                    if template_view:
                        new_3d_view.ViewTemplateId = template_view.Id
                    new_3d_view.IsSectionBoxActive = False
                    prepare_view_visibility(
                        new_3d_view, clear_template=(template_view is None))

            if create_north_section or create_east_section:
                if view_family_section:
                    center = XYZ(
                        (fitted_min.X + fitted_max.X) / 2.0,
                        (fitted_min.Y + fitted_max.Y) / 2.0,
                        (fitted_min.Z + fitted_max.Z) / 2.0,
                    )
                    half_x = (fitted_max.X - fitted_min.X) / 2.0
                    half_y = (fitted_max.Y - fitted_min.Y) / 2.0

                    section_defs = []

                    if create_north_section:
                        section_defs.append({
                            "label": "North",
                            "bx": XYZ(1, 0, 0),
                            "by": XYZ(0, 0, 1),
                            "bz": XYZ(0, 1, 0),
                            "half_w": half_x,
                            "depth": half_y * 2,
                        })

                    if create_east_section:
                        section_defs.append({
                            "label": "East",
                            "bx": XYZ(0, -1, 0),
                            "by": XYZ(0, 0, 1),
                            "bz": XYZ(1, 0, 0),
                            "half_w": half_y,
                            "depth": half_x * 2,
                        })

                    for sdef in section_defs:
                        desired_section_name = "{} - Section {}".format(
                            view_base_name, sdef["label"])
                        existing_section = find_view_by_name(
                            desired_section_name)
                        if existing_section:
                            try:
                                if existing_section.ViewType == DB.ViewType.Section:
                                    apply_section_view_type(
                                        existing_section, view_family_section)
                            except BaseException:
                                pass
                            if section_template_candidates:
                                apply_first_valid_template(
                                    existing_section, section_template_candidates)
                            continue

                        t = Transform.Identity
                        t.BasisX = sdef["bx"]
                        t.BasisY = sdef["by"]
                        t.BasisZ = sdef["bz"]
                        t.Origin = center

                        # Section vertical extents follow adjacent levels with 1-inch offsets.
                        p_bottom = XYZ(center.X, center.Y, section_bottom_z)
                        p_top = XYZ(center.X, center.Y, section_top_z)
                        local_bottom_y = t.Inverse.OfPoint(p_bottom).Y
                        local_top_y = t.Inverse.OfPoint(p_top).Y

                        if local_top_y <= local_bottom_y:
                            old_bottom_y = local_bottom_y
                            old_top_y = local_top_y
                            local_bottom_y = min(old_bottom_y, old_top_y)
                            local_top_y = max(old_bottom_y, old_top_y)

                        sbox = BoundingBoxXYZ()
                        sbox.Transform = t
                        sbox.Min = XYZ(-sdef["half_w"], local_bottom_y, 0)
                        sbox.Max = XYZ(
                            sdef["half_w"], local_top_y, sdef["depth"])

                        sv = ViewSection.CreateSection(
                            doc, view_family_section.Id, sbox)
                        sv.Name = desired_section_name

                        # Enforce configured section view type so type-based callout/tag settings match.
                        apply_section_view_type(sv, view_family_section)

                        applied_section_template = None
                        if section_template_candidates:
                            applied_section_template = apply_first_valid_template(
                                sv, section_template_candidates)
                        if not applied_section_template:
                            prepare_view_visibility(sv, clear_template=True)

                        section_views.append(sv)

        if new_3d_view:
            with revit.Transaction("Fit 3D Section Box - {}".format(selected_family)):
                section_box_3d = BoundingBoxXYZ()
                section_box_3d.Min = fitted_min
                section_box_3d.Max = fitted_max
                new_3d_view.SetSectionBox(section_box_3d)
                new_3d_view.IsSectionBoxActive = True
            last_3d_view = new_3d_view

        total_section_views += len(section_views)

    if last_3d_view:
        uidoc.ActiveView = last_3d_view

    if all_element_ids:
        unique_ids = {}
        for eid in all_element_ids:
            unique_ids[eid.IntegerValue] = eid
        ids_list = List[ElementId](list(unique_ids.values()))
        uidoc.Selection.SetElementIds(ids_list)
        uidoc.ShowElements(ids_list)

except Exception as ex:
    forms.alert("{}".format(ex), title="Create Views Error")
    output.print_md("## Error: {}".format(ex))
    import traceback
    output.print_md("```\n{}\n```".format(traceback.format_exc()))

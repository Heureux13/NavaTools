# -*- coding: utf-8 -*-
# duct_tagger.py
# Copyright (c) 2025 Jose Francisco Nava Perez
# All rights reserved. No part of this code may be reproduced without permission.

from pyrevit import DB
from Autodesk.Revit.DB import FabricationPart
from duct_shadow import DuctShadow


class DuctTagger:
    def __init__(self, doc, view):
        self.doc = doc
        self.view = view
        # Collect fabrication ductwork tag symbols
        self.tag_syms = (DB.FilteredElementCollector(doc)
                         .OfClass(DB.FamilySymbol)
                         .OfCategory(DB.BuiltInCategory.OST_FabricationDuctworkTags)
                         .ToElements())

    # ---- helpers ----
    def get_label(self, name_contains):
        needle = name_contains.lower()
        for ts in self.tag_syms:
            fam = getattr(ts, "Family", None)
            fam_name = fam.Name if fam else ""
            ts_name = getattr(ts, "Name", "")
            pool = (fam_name + " " + ts_name).lower()
            if needle in pool:
                return ts
        raise Exception("No tag symbol found with: " + name_contains)

    def already_tagged(self, elem, tag_fam_name):
        existing = (DB.FilteredElementCollector(self.doc, self.view.Id)
                    .OfClass(DB.IndependentTag)
                    .ToElements())
        for itag in existing:
            ref_elem = None
            try:
                ref_elem = itag.GetTaggedLocalElement()
            except:
                pass
            if ref_elem and ref_elem.Id == elem.Id:
                fam_name = ""
                try:
                    t = itag.GetType()
                    fam_name = t.FamilyName if t else ""
                except:
                    fam_name = ""
                if fam_name == tag_fam_name:
                    return True
        return False

    def place(self, element, tag_symbol, point_xyz):
        ref = DB.Reference(element)
        tag = DB.IndependentTag.Create(
            self.doc,
            self.view.Id,
            ref,
            False,
            DB.TagMode.TM_ADDBY_CATEGORY,
            DB.TagOrientation.Horizontal,
            point_xyz
        )
        if tag_symbol and tag_symbol.Id:
            tag.ChangeTypeId(tag_symbol.Id)
        tag.HasLeader = False
        # Force the head where we want it
        tag.TagHeadPosition = point_xyz
        return tag

    def _pick_anchor_by_orientation(self, axis_vec, override=None):
        if override:
            return override
        # Project axis into view plane
        u = axis_vec.DotProduct(self.view.RightDirection)
        v = axis_vec.DotProduct(self.view.UpDirection)
        # Horizontal bias bottom-left; Vertical bias bottom-right
        return "bottom_left" if abs(u) >= abs(v) else "bottom_right"

    def tag_all(self, tag_keyword="length", anchor=None, debug=False):
        ducts = (DB.FilteredElementCollector(self.doc, self.view.Id)
                .OfCategory(DB.BuiltInCategory.OST_FabricationDuctwork)
                .WhereElementIsNotElementType()
                .ToElements())

        if debug:
            print("Ducts found:", len(ducts))
            print("Tag symbols found:", len(self.tag_syms))
            for ts in self.tag_syms:
                fam = getattr(ts, "Family", None)
                fam_name = getattr(fam, "Name", "<no family>") if fam is not None else "<no family>"
                type_name = getattr(ts, "Name", "<no type name>")
                print("Tag symbol:", fam_name, type_name)

        # Find tag symbol
        tag_sym = None
        try:
            tag_sym = self.get_label(tag_keyword)
        except Exception as e:
            print("Tag symbol lookup failed:", e)
            return

        if tag_sym is None:
            print("No tag symbol matched keyword:", tag_keyword)
            return

        t = DB.Transaction(self.doc, "Tag fabrication ducts with DuctShadow")
        t.Start()

        try:
            if not tag_sym.IsActive:
                tag_sym.Activate()
                self.doc.Regenerate()

            fam_obj = getattr(tag_sym, "Family", None)
            tag_fam_name = fam_obj.Name if fam_obj else ""
        except Exception as e:
            print("Failed to activate tag symbol:", e)
            t.RollBack()
            return

        # inner helper: robust size lookup
        def _get_size_from_params(instance):
            ps = {}
            try:
                for p in instance.Parameters:
                    try:
                        ps[p.Definition.Name.lower()] = p
                    except:
                        continue
            except:
                pass

            w = h = None
            for name, p in ps.items():
                lname = name.lower()
                try:
                    val = p.AsDouble()
                except:
                    try:
                        vstr = p.AsValueString()
                        val = float(vstr) if vstr and vstr.strip() != "" else None
                    except:
                        val = None
                if val is None:
                    continue
                if "width" in lname or lname in ("w", "rectangular width", "actual width"):
                    w = val
                elif "height" in lname or lname in ("h", "rectangular height", "actual height"):
                    h = val
                elif "diameter" in lname or "od" in lname:
                    w = h = val
                elif lname in ("size", "size1", "size2") and (w is None or h is None):
                    if w is None:
                        w = val
                    elif h is None:
                        h = val
            return w, h

        placed = 0
        for duct in ducts:
            # Verify element type
            if not isinstance(duct, FabricationPart):
                if debug: print("Skip: not FabricationPart", getattr(duct, "Id", "<no id>"))
                continue

            # Already tagged check (defensive)
            try:
                if tag_fam_name and self.already_tagged(duct, tag_fam_name):
                    if debug: print("Skip: already tagged", duct.Id)
                    continue
            except Exception as e:
                if debug: print("already_tagged error:", e)

            # Location + curve
            loc = getattr(duct, "Location", None)
            if not isinstance(loc, DB.LocationCurve):
                if debug: print("Skip: no LocationCurve", duct.Id)
                continue
            curve = getattr(loc, "Curve", None)
            if not curve:
                if debug: print("Skip: no curve", duct.Id)
                continue

            # Axis
            try:
                p0 = curve.GetEndPoint(0)
                p1 = curve.GetEndPoint(1)
                axis_vec = (p1 - p0).Normalize()
            except Exception as e:
                if debug: print("Skip: axis normalize failed", duct.Id, e)
                continue

            # --- robust size lookup start ---
            # Try instance params first
            width, height = _get_size_from_params(duct)

            # If not found, try type parameters
            if (width is None or height is None):
                try:
                    typ = duct.GetType()
                    if typ is not None:
                        tw, th = _get_size_from_params(typ)
                        width = width or tw
                        height = height or th
                except:
                    pass

            # Fallback: bounding box approximation (view-aware if possible)
            if (width is None or height is None):
                try:
                    bb = duct.get_BoundingBox(self.view)
                    if not bb:
                        bb = duct.get_BoundingBox(None)
                    if bb:
                        diag = bb.Max - bb.Min
                        comps = [abs(diag.X), abs(diag.Y), abs(diag.Z)]
                        comps.sort(reverse=True)
                        if width is None:
                            width = comps[0]
                        if height is None:
                            height = comps[1] if len(comps) > 1 else comps[0]
                except:
                    pass

            # If still missing, skip this duct
            if width is None or height is None:
                if debug: print("Skip: missing size params", duct.Id)
                continue
            # --- robust size lookup end ---

            length = getattr(curve, "Length", None)
            if length is None:
                if debug: print("Skip: missing curve length", duct.Id)
                continue

            # Anchor selection fallback name (kept for compatibility but not used below)
            anchor_name = self._pick_anchor_by_orientation(axis_vec, override=anchor)

            # --- compute robust anchors that lie inside the openings ---
            try:
                shadow = DuctShadow(p0, axis_vec, width, height, length, self.view)
            except Exception as e:
                if debug: print("Skip: shadow construction failed", duct.Id, e)
                continue

            # start opening centroid -> nudge inward along axis_vec (toward p1) so tag is inside opening
            sc = getattr(shadow, "start_corners", None)
            if sc:
                start_centroid = DB.XYZ(sum(c.X for c in sc)/len(sc),
                                        sum(c.Y for c in sc)/len(sc),
                                        sum(c.Z for c in sc)/len(sc))
            else:
                start_centroid = p0
            # small inward nudge (feet). Positive along axis_vec moves from start into duct.
            INWARD = 0.02
            start_anchor_model = start_centroid + axis_vec.Multiply(INWARD)

            # end opening centroid -> nudge inward along -axis_vec (toward p0) so tag is inside opening
            ec = getattr(shadow, "end_corners", None)
            if ec:
                end_centroid = DB.XYZ(sum(c.X for c in ec)/len(ec),
                                    sum(c.Y for c in ec)/len(ec),
                                    sum(c.Z for c in ec)/len(ec))
            else:
                end_centroid = p1
            end_anchor_model = end_centroid + axis_vec.Multiply(-INWARD)

            # Project to view plane and rebuild model points for stable TagHeadPosition in view
            s2 = ( (start_anchor_model - self.view.Origin).DotProduct(self.view.RightDirection),
                (start_anchor_model - self.view.Origin).DotProduct(self.view.UpDirection) )
            e2 = ( (end_anchor_model - self.view.Origin).DotProduct(self.view.RightDirection),
                (end_anchor_model - self.view.Origin).DotProduct(self.view.UpDirection) )

            start_anchor = self.view.Origin + self.view.RightDirection.Multiply(s2[0]) + self.view.UpDirection.Multiply(s2[1])
            end_anchor   = self.view.Origin + self.view.RightDirection.Multiply(e2[0]) + self.view.UpDirection.Multiply(e2[1])

            # tiny vertical offsets so start/end tags don't collide visually
            start_anchor = start_anchor + self.view.UpDirection.Multiply(0.03)
            end_anchor   = end_anchor   + self.view.UpDirection.Multiply(-0.03)

            if debug:
                print("Duct", duct.Id, "start_anchor (inward):", start_anchor, "end_anchor (inward):", end_anchor)

            # --- single combined tag (one tag per duct showing both openings) ---
            try:
                # compute opening sizes from shadow corners if available, else fall back to width/height
                sc = getattr(shadow, "start_corners", None)
                if sc:
                    s_w, s_h = (abs(max(c.X for c in sc) - min(c.X for c in sc)),
                                abs(max(c.Y for c in sc) - min(c.Y for c in sc))) if len(sc) >= 2 else (width, height)
                else:
                    s_w, s_h = width, height

                ec = getattr(shadow, "end_corners", None)
                if ec:
                    e_w, e_h = (abs(max(c.X for c in ec) - min(c.X for c in ec)),
                                abs(max(c.Y for c in ec) - min(c.Y for c in ec))) if len(ec) >= 2 else (width, height)
                else:
                    e_w, e_h = width, height

                # format sizes to simple strings in model units (feet). Adjust formatting if you want inches or mm.
                def fmt(val):
                    return "{:.2f}".format(val) if val is not None else "?"
                start_size_str = "{}x{}".format(fmt(s_w), fmt(s_h))
                end_size_str   = "{}x{}".format(fmt(e_w), fmt(e_h))
                combined_label = "{} / {}".format(start_size_str, end_size_str)

                if debug:
                    print("Duct", duct.Id, "start_size(ft):", start_size_str, "end_size(ft):", end_size_str)

                # choose the start opening's lower-left corner (view-space) and nudge inward a little
                if getattr(shadow, "start_corners", None):
                    centroid = DB.XYZ(sum(c.X for c in shadow.start_corners)/len(shadow.start_corners),
                                    sum(c.Y for c in shadow.start_corners)/len(shadow.start_corners),
                                    sum(c.Z for c in shadow.start_corners)/len(shadow.start_corners))
                else:
                    centroid = start_centroid
                mid_model = centroid + axis_vec.Multiply(0.08)
                mid2 = ((mid_model - self.view.Origin).DotProduct(self.view.RightDirection),
                        (mid_model - self.view.Origin).DotProduct(self.view.UpDirection))
                mid_anchor = self.view.Origin + self.view.RightDirection.Multiply(mid2[0]) + self.view.UpDirection.Multiply(mid2[1])
                mid_anchor = mid_anchor + self.view.UpDirection.Multiply(0.03)

                mid2 = ((mid_model - self.view.Origin).DotProduct(self.view.RightDirection),
                        (mid_model - self.view.Origin).DotProduct(self.view.UpDirection))
                mid_anchor = self.view.Origin + self.view.RightDirection.Multiply(mid2[0]) + self.view.UpDirection.Multiply(mid2[1])
                mid_anchor = mid_anchor + self.view.UpDirection.Multiply(0.03)

                tag = self.place(duct, tag_sym, mid_anchor)
                tag.TagOrientation = DB.TagOrientation.Horizontal
                tag.HasLeader = False

                # optional: set the tag's visible text if your tag family exposes a parameter or if it's an independent text tag
                try:
                    p = None
                    for pname in ("Label","Text","Type Mark","Comments"):
                        try:
                            p = tag.LookupParameter(pname)
                        except:
                            p = None
                        if p:
                            try:
                                p.Set(combined_label)
                                break
                            except:
                                p = None
                    if debug and not p:
                        print("No writable tag parameter found to set combined text; tag placed only.")
                except Exception as e:
                    if debug: print("Setting tag text failed:", e)

                placed += 1

            except Exception as e:
                if debug: print("Combined tag placement failed for", duct.Id, e)

        if debug:
            print("Tags placed:", placed)

        self.doc.Regenerate()
        t.Commit()
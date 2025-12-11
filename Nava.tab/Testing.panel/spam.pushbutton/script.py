# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from offsets import Offsets
from pyrevit import revit, script

# Button info
# ===================================================
__title__ = "SPam"
__doc__ = """
Test the Offsets class with selected duct/pipe elements
"""

# Variables
# ==================================================
output = script.get_output()

# Main
# ==================================================
selection = revit.get_selection()

if not selection:
    output.print_md("Select at least one element.")
else:
    for el in selection:
        loc = getattr(el, "Location", None)
        curve = getattr(loc, "Curve", None) if loc else None
        if curve:
            sp = curve.GetEndPoint(0)
            ep = curve.GetEndPoint(1)
            output.print_md("Element {}: start=({:.3f}, {:.3f}, {:.3f}), end=({:.3f}, {:.3f}, {:.3f})".format(
                el.Id.Value,
                sp.X, sp.Y, sp.Z,
                ep.X, ep.Y, ep.Z,
            ))
        else:
            # Try connectors for fabrication parts
            printed = False
            try:
                cm = getattr(el, 'ConnectorManager', None)
                connectors = cm.Connectors if cm else getattr(
                    el, 'Connectors', None)
                count = getattr(connectors, 'Size', getattr(
                    connectors, 'Count', 0)) if connectors else 0
                origins = []
                if connectors:
                    if count > 0 and hasattr(connectors, 'Item'):
                        for i in range(count):
                            c = connectors.Item(i)
                            o = getattr(c, 'Origin', None)
                            if o:
                                origins.append(o)
                    else:
                        try:
                            for c in connectors:
                                o = getattr(c, 'Origin', None)
                                if o:
                                    origins.append(o)
                        except Exception:
                            pass
                if origins:
                    printed = True
                    for idx, o in enumerate(origins):
                        output.print_md("Element {}: connector {} origin=({:.3f}, {:.3f}, {:.3f})".format(
                            el.Id.Value, idx, o.X, o.Y, o.Z))
            except Exception:
                pass

            # FabricationPart specific primary/secondary connectors
            if not printed:
                try:
                    pc = getattr(el, 'PrimaryConnector', None)
                    sc = getattr(el, 'SecondaryConnector', None)
                    origins = []
                    if pc and getattr(pc, 'Origin', None):
                        origins.append(('primary', pc.Origin))
                    if sc and getattr(sc, 'Origin', None):
                        origins.append(('secondary', sc.Origin))
                    if origins:
                        printed = True
                        for label, o in origins:
                            output.print_md("Element {}: {} connector origin=({:.3f}, {:.3f}, {:.3f})".format(
                                el.Id.Value, label, o.X, o.Y, o.Z))
                except Exception:
                    pass

            # Some fabrication APIs expose GetConnectors()
            if not printed:
                try:
                    get_conns = getattr(el, 'GetConnectors', None)
                    if get_conns:
                        conns = get_conns()
                        if conns:
                            origins = []
                            for c in conns:
                                o = getattr(c, 'Origin', None)
                                if o:
                                    origins.append(o)
                            if origins:
                                printed = True
                                for idx, o in enumerate(origins):
                                    output.print_md("Element {}: connector {} origin=({:.3f}, {:.3f}, {:.3f})".format(
                                        el.Id.Value, idx, o.X, o.Y, o.Z))
                except Exception:
                    pass

            if not printed:
                output.print_md(
                    "Element {}: no Location.Curve and no connectors".format(el.Id.Value))

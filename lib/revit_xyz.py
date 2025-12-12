# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

# Constants
TOL = 1e-6


class RevitXYZ(object):
    """Extract XYZ coordinates and orientation from element connectors."""

    def __init__(self, element):
        self.element = element
        self.loc = getattr(element, "Location", None)
        self.curve = getattr(self.loc, "Curve", None) if self.loc else None

    def _get_all_connectors(self):
        """Return list of all connector objects from element."""
        connectors = []

        try:
            # ConnectorManager / Connectors (standard MEP elements)
            cm = getattr(self.element, 'ConnectorManager', None)
            conn_collection = cm.Connectors if cm else getattr(
                self.element, 'Connectors', None)

            if conn_collection:
                count = getattr(conn_collection, 'Size',
                                getattr(conn_collection, 'Count', 0))
                if count and hasattr(conn_collection, 'Item'):
                    for i in range(count):
                        c = conn_collection.Item(i)
                        if c:
                            connectors.append(c)
                try:
                    for c in conn_collection:
                        if c:
                            connectors.append(c)
                except Exception:
                    pass

            # Primary/Secondary connectors (fabrication parts)
            pc = getattr(self.element, 'PrimaryConnector', None)
            sc = getattr(self.element, 'SecondaryConnector', None)
            if pc:
                connectors.append(pc)
            if sc:
                connectors.append(sc)

            # GetConnectors API (some fabrication elements)
            get_conns = getattr(self.element, 'GetConnectors', None)
            if get_conns:
                try:
                    conns = get_conns()
                    if conns:
                        for c in conns:
                            if c:
                                connectors.append(c)
                except Exception:
                    pass
        except Exception:
            pass

        return connectors

    def connector_data(self):
        """Return list of dicts with connector origin and orientation vectors.

        Each dict contains:
            'origin': XYZ point
            'basis_x': XYZ vector (width direction)
            'basis_y': XYZ vector (height direction)  
            'basis_z': XYZ vector (flow direction)

        Returns empty list if no connectors found.
        """
        connectors = self._get_all_connectors()
        data = []
        seen = set()

        for conn in connectors:
            origin = getattr(conn, 'Origin', None)
            if not origin:
                continue

            # Deduplicate by origin
            key = (round(origin.X, 9), round(origin.Y, 9), round(origin.Z, 9))
            if key in seen:
                continue
            seen.add(key)

            # Try to get coordinate system
            coord_sys = getattr(conn, 'CoordinateSystem', None)
            if coord_sys:
                data.append({
                    'origin': origin,
                    'basis_x': coord_sys.BasisX,
                    'basis_y': coord_sys.BasisY,
                    'basis_z': coord_sys.BasisZ
                })
            else:
                # No coordinate system available
                data.append({
                    'origin': origin,
                    'basis_x': None,
                    'basis_y': None,
                    'basis_z': None
                })

        return data

    def connector_origins(self):
        """Return list of unique connector origin XYZ points from element.

        Tries multiple connector APIs and deduplicates by rounding to 9 decimals.
        Returns empty list if no connectors found.
        """
        data = self.connector_data()
        return [d['origin'] for d in data]

    def curve_endpoints(self):
        """Return (start, end) XYZ points from Location.Curve.

        Returns (None, None) if element has no curve.
        """
        if self.curve:
            return self.curve.GetEndPoint(0), self.curve.GetEndPoint(1)
        return None, None

    def inlet_outlet_data(self):
        """Return inlet and outlet connector data with orientation.

        Returns tuple: (inlet_dict, outlet_dict) where each dict contains:
            'origin': XYZ point
            'basis_x': XYZ vector (width direction)
            'basis_y': XYZ vector (height direction)  
            'basis_z': XYZ vector (flow direction)

        For elements with 2+ connectors, returns the two farthest apart.
        Returns (None, None) if no data available or orientation unavailable.
        """
        # Prefer Primary/Secondary connectors when present
        pc = getattr(self.element, 'PrimaryConnector', None)
        sc = getattr(self.element, 'SecondaryConnector', None)

        def _build_data(conn):
            if not conn:
                return None
            origin = getattr(conn, 'Origin', None)
            if not origin:
                return None
            cs = getattr(conn, 'CoordinateSystem', None)
            return {
                'origin': origin,
                'basis_x': cs.BasisX if cs else None,
                'basis_y': cs.BasisY if cs else None,
                'basis_z': cs.BasisZ if cs else None,
            }

        if pc and sc:
            inlet = _build_data(pc)
            outlet = _build_data(sc)
            if inlet and outlet:
                return inlet, outlet

        conn_data = self.connector_data()

        # Two or more distinct connectors
        if len(conn_data) >= 2:
            # Find the pair with maximum distance
            max_dist_sq = -1.0
            inlet, outlet = conn_data[0], conn_data[1]
            for i in range(len(conn_data)):
                for j in range(i + 1, len(conn_data)):
                    o1 = conn_data[i]['origin']
                    o2 = conn_data[j]['origin']
                    dx = o1.X - o2.X
                    dy = o1.Y - o2.Y
                    dz = o1.Z - o2.Z
                    dist_sq = dx * dx + dy * dy + dz * dz
                    if dist_sq > max_dist_sq:
                        max_dist_sq = dist_sq
                        inlet, outlet = conn_data[i], conn_data[j]
            return inlet, outlet

        return None, None

    def inlet_outlet_points(self):
        """Return (inlet, outlet) XYZ points using connectors first, then curve.

        For elements with 2+ connectors, returns the two farthest apart.
        For single connector + curve, pairs connector with farthest endpoint.
        For curve only, returns start and end points.

        Returns (None, None) if no data available.
        """
        origins = self.connector_origins()

        # Two or more distinct connector origins
        if len(origins) >= 2:
            # Find the pair with maximum distance (inlet/outlet direction)
            max_dist_sq = -1.0
            inlet, outlet = origins[0], origins[1]
            for i in range(len(origins)):
                for j in range(i + 1, len(origins)):
                    dx = origins[i].X - origins[j].X
                    dy = origins[i].Y - origins[j].Y
                    dz = origins[i].Z - origins[j].Z
                    dist_sq = dx * dx + dy * dy + dz * dz
                    if dist_sq > max_dist_sq:
                        max_dist_sq = dist_sq
                        inlet, outlet = origins[i], origins[j]
            return inlet, outlet

        # One connector + curve available
        if len(origins) == 1 and self.curve:
            p0, p1 = self.curve_endpoints()
            c = origins[0]
            d0_sq = (c.X - p0.X) ** 2 + (c.Y - p0.Y) ** 2 + (c.Z - p0.Z) ** 2
            d1_sq = (c.X - p1.X) ** 2 + (c.Y - p1.Y) ** 2 + (c.Z - p1.Z) ** 2
            # Use farthest curve endpoint as outlet to create a real direction
            outlet = p1 if d0_sq <= d1_sq else p0
            return c, outlet

        # Fallback to curve endpoints
        if self.curve:
            p0, p1 = self.curve_endpoints()
            return p0, p1

        # No data
        return None, None

    def all_connector_origins(self):
        """Return all connector origins as a list.

        Useful for tees, crosses, and other multi-connector elements.
        Returns the same list as connector_origins() but clearly marked
        for use cases where you need all connection points, not just two.
        """
        return self.connector_origins()

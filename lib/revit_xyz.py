# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

import math

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
            try:
                origin = conn.Origin
            except Exception:
                continue
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
                    'basis_z': coord_sys.BasisZ,
                    'connector': conn,
                })
            else:
                # No coordinate system available
                data.append({
                    'origin': origin,
                    'basis_x': None,
                    'basis_y': None,
                    'basis_z': None,
                    'connector': conn,
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

    def inlet_outlet_data(self, size_obj=None):
        """Return inlet and outlet connector data with orientation.

        Returns tuple: (inlet_dict, outlet_dict) where each dict contains:
            'origin': XYZ point
            'basis_x': XYZ vector (width direction)
            'basis_y': XYZ vector (height direction)
            'basis_z': XYZ vector (flow direction)

        Args:
            size_obj: Optional Size object to match physical connectors to inlet/outlet sizes

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
                'connector': conn,
            } if origin else None

        if pc and sc:
            inlet = _build_data(pc)
            outlet = _build_data(sc)
            if inlet and outlet:
                # If size object provided, verify assignment matches
                if size_obj and size_obj.in_size != size_obj.out_size:
                    inlet, outlet = self._match_by_size(
                        inlet, outlet, size_obj)
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

            # If size object provided and sizes differ, match by dimensions
            if size_obj and size_obj.in_size != size_obj.out_size:
                inlet, outlet = self._match_by_size(inlet, outlet, size_obj)
            # Otherwise use flow direction from basis_z
            else:
                o1 = inlet['origin']
                o2 = outlet['origin']
                bz1 = inlet.get('basis_z')
                bz2 = outlet.get('basis_z')

                # If we have both basis_z vectors, check if they point toward each other
                if bz1 and bz2:
                    # Vector from inlet to outlet
                    dx = o2.X - o1.X
                    dy = o2.Y - o1.Y
                    dz = o2.Z - o1.Z
                    # Dot product: if inlet basis_z points toward outlet, it's correct
                    dot1 = bz1.X * dx + bz1.Y * dy + bz1.Z * dz
                    # If negative, inlet basis is pointing away - swap them
                    if dot1 < 0:
                        inlet, outlet = outlet, inlet

            return inlet, outlet

        return None, None

    def _match_by_size(self, conn1, conn2, size_obj):
        """Match physical connectors to Size parameter inlet/outlet by dimensions.

        Returns (inlet, outlet) tuple matched to size_obj.in_size and size_obj.out_size.
        """
        # Get connector dimensions (in inches)
        def _get_dimensions(conn_dict):
            conn = conn_dict.get('connector')
            if not conn:
                return None, None, None

            # Try to get width, height, diameter from connector
            try:
                width = getattr(conn, 'Width', None)
                height = getattr(conn, 'Height', None)
                radius = getattr(conn, 'Radius', None)

                # Convert feet to inches
                w = width * 12 if width else None
                h = height * 12 if height else None
                d = radius * 2 * 12 if radius else None

                return w, h, d
            except Exception:
                return None, None, None

        w1, h1, d1 = _get_dimensions(conn1)
        w2, h2, d2 = _get_dimensions(conn2)

        # Match conn1 dimensions to either inlet or outlet size
        # Compare using the dominant dimension
        def _size_match_score(conn_w, conn_h, conn_d, size_w, size_h, size_d):
            """Calculate how well connector dimensions match size dimensions."""
            score = 0
            tol = 1.0  # 1 inch tolerance

            if conn_d and size_d:
                # Round comparison
                if abs(conn_d - size_d) < tol:
                    score += 10
            elif conn_w and conn_h and size_w and size_h:
                # Rectangular comparison
                if abs(conn_w - size_w) < tol and abs(conn_h - size_h) < tol:
                    score += 10
                # Allow for swapped dimensions
                elif abs(conn_w - size_h) < tol and abs(conn_h - size_w) < tol:
                    score += 8

            return score

        # Score conn1 against inlet size
        score_1_to_in = _size_match_score(
            w1, h1, d1,
            size_obj.in_width, size_obj.in_height, size_obj.in_diameter
        )
        # Score conn1 against outlet size
        score_1_to_out = _size_match_score(
            w1, h1, d1,
            size_obj.out_width, size_obj.out_height, size_obj.out_diameter
        )

        # If conn1 matches inlet better, return as-is; otherwise swap
        if score_1_to_in >= score_1_to_out:
            return conn1, conn2
        else:
            return conn2, conn1

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
            if p0 and p1:
                c = origins[0]
                d0_sq = (c.X - p0.X) ** 2 + \
                    (c.Y - p0.Y) ** 2 + (c.Z - p0.Z) ** 2
                d1_sq = (c.X - p1.X) ** 2 + \
                    (c.Y - p1.Y) ** 2 + (c.Z - p1.Z) ** 2
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

    def straight_joint_degree(self):
        """Returns the angle in degrees between the duct and the horizontal (XY) plane."""
        inlet, outlet = self.inlet_outlet_points()
        if not inlet or not outlet:
            return None

        dx = outlet.X - inlet.X
        dy = outlet.Y - inlet.Y
        dz = outlet.Z - inlet.Z

        horizontal_length = math.sqrt(dx**2 + dy**2)
        if horizontal_length == 0:
            return 90.0 if dz != 0 else 0.0

        angle_rad = math.atan2(dz, horizontal_length)
        angle_deg = math.degrees(angle_rad)
        return round(angle_deg, 2)

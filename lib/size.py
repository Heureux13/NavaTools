# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

# Imports
# ========================================================================
import re


class Size:
    def __init__(self, size):
        self.size = size
        parsed = self._parse_size()
        self.in_size = parsed['in_size']
        self.in_width = parsed['in_width']
        self.in_height = parsed['in_height']
        self.in_diameter = parsed['in_diameter']
        self.in_oval_dia = parsed['in_oval_dia']
        self.in_oval_flat = parsed['in_oval_flat']
        self.out_size = parsed['out_size']
        self.out_width = parsed['out_width']
        self.out_height = parsed['out_height']
        self.out_diameter = parsed['out_diameter']
        self.out_oval_dia = parsed['out_oval_dia']
        self.out_oval_flat = parsed['out_oval_flat']

    def _parse_size(self):
        s = str(self.size).strip().replace('"', '').lower()
        if '-' in s:
            inlet, outlet = [p.strip() for p in s.split('-', 1)]
        else:
            inlet, outlet = s, None
        in_data = self._parse_token(inlet)
        out_data = self._parse_token(outlet) if outlet else in_data.copy()
        return {
            'in_size': inlet,
            'in_width': in_data.get('width'),
            'in_height': in_data.get('height'),
            'in_diameter': in_data.get('diameter'),
            'in_oval_dia': in_data.get('oval_dia'),
            'in_oval_flat': in_data.get('oval_flat'),
            'out_size': outlet if outlet else inlet,
            'out_width': out_data.get('width'),
            'out_height': out_data.get('height'),
            'out_diameter': out_data.get('diameter'),
            'out_oval_dia': out_data.get('oval_dia'),
            'out_oval_flat': out_data.get('oval_flat'),
        }

    def _parse_token(self, token):
        result = {}
        if not token:
            result['width'] = result['height'] = result['diameter'] = None
            result['oval_dia'] = result['oval_flat'] = None
            return result

        # Logic for rounds
        m = re.match(r'(\d+(?:\.\d+)?)\s*[øØ]', token)
        if m:
            result['diameter'] = float(m.group(1))
            result['width'] = result['height'] = None
            result['oval_dia'] = result['oval_flat'] = None
            return result

        # Logic for ovals
        m = re.match(r'(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)', token)
        if m:
            w = float(m.group(1))
            h = float(m.group(2))
            result['width'] = w
            result['height'] = h
            result['diameter'] = h  # per user: oval diameter = height
            result['oval_dia'] = (w + h) / 2.0
            result['oval_flat'] = (max(w, h) - min(w, h)) if w != h else 0.0
            return result

        # Logic for rectangle / square
        m = re.match(r'(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)', token)
        if m:
            result['width'] = float(m.group(1))
            result['height'] = float(m.group(2))
            result['diameter'] = None
            result['oval_dia'] = result['oval_flat'] = None
            return result

        result['width'] = result['height'] = result['diameter'] = None
        result['oval_dia'] = result['oval_flat'] = None
        return result

    def in_shape(self):
        if self.in_diameter is not None:
            return "round"
        if self.in_oval_dia is not None:
            return "oval"
        if self.in_width is not None and self.in_height is not None:
            return "rectangle"

    def out_shape(self):
        if self.out_diameter is not None:
            return "round"
        if self.out_oval_dia is not None:
            return "oval"
        if self.out_width is not None and self.out_height is not None:
            return "rectangle"


if __name__ == "__main__":
    # Quick sanity examples
    for sample in [
        "40/20-12ø", "40/20", "12x12", "12ø"
    ]:
        rs = Size(sample)
        print("\nSample:", sample)
        print("  inlet size:", rs.in_size)
        print("    width:", rs.in_width)
        print("    height:", rs.in_height)
        print("    diameter:", rs.in_diameter)
        print("    oval_flat:", rs.in_oval_flat)
        print("  outlet size:", rs.out_size)
        print("    width:", rs.out_width)
        print("    height:", rs.out_height)
        print("    diameter:", rs.out_diameter)
        print("    oval_flat:", rs.out_oval_flat)

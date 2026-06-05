# -*- coding: utf-8 -*-

"""Backward-compatible bridge for legacy imports.

Existing tools import tag configuration as `tagging.tag_config`.
The canonical module now lives at `config.tag_config`, so re-export
its public names here to keep old imports working.
"""

from config.tag_config import *

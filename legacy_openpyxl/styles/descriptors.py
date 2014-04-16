# Copyright (c) 2010-2014 openpyxl

"""
Package level descriptors
"""

from legacy_openpyxl.descriptors import Typed, Default
from .colors import Color


class Color(Default):

    expected_type = Color
# file openpyxl/reader/worksheet.py

# Copyright (c) 2010 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: Eric Gazoni

"""Reader for a single worksheet."""

# Python stdlib imports
try:
    from xml.etree.cElementTree import iterparse
except ImportError:
    from xml.etree.ElementTree import iterparse
from itertools import ifilter
from StringIO import StringIO

# package imports
from legacy_openpyxl.cell import Cell
from legacy_openpyxl.worksheet import Worksheet

def filter_cells((event, element)):

    return element.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'

def fast_parse(ws, xml_source, string_table, style_table):

    source = StringIO(xml_source)

    it = iterparse(source)

    for event, element in ifilter(filter_cells, it):

        coordinate = element.get('r')
        data_type = element.get('t', 'n')
        style_id = element.get('s')
        value = element.findtext('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')

        if value is not None:

            if data_type == Cell.TYPE_STRING:
                value = string_table.get(int(value))

            ws.cell(coordinate).value = value

            if style_id is not None:
                ws._styles[coordinate] = style_table.get(int(style_id))

        # to avoid memory exhaustion, clear the item after use
        element.clear()

def read_worksheet(xml_source, parent, preset_title,
        string_table, style_table):
    """Read an xml worksheet"""
    ws = Worksheet(parent, preset_title)
    fast_parse(ws, xml_source, string_table, style_table)
    return ws

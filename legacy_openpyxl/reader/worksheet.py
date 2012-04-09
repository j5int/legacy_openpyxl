# file openpyxl/reader/worksheet.py

# Copyright (c) 2010-2011 openpyxl
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
# @author: see AUTHORS file

"""Reader for a single worksheet."""

# Python stdlib imports
try:
    from legacy_openpyxl.shared.compat import iterparse
except ImportError:
    from xml.etree.ElementTree import iterparse

# package imports
from legacy_openpyxl.cell import Cell, coordinate_from_string
from legacy_openpyxl.worksheet import Worksheet
from legacy_openpyxl.shared.compat import ifilter, StringIO, unicode

def _get_xml_iter(xml_source):
    if not hasattr(xml_source, 'name'):
        return StringIO(unicode(xml_source))
    else:
        xml_source.seek(0)
        return xml_source

def read_dimension(xml_source):

    source = _get_xml_iter(xml_source)

    it = iterparse(source)

    smax_col = None
    smax_row = None
    smin_col = None
    smin_row = None

    for event, element in it:

        if element.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}dimension':
            ref = element.get('ref')

            if ':' in ref:
                min_range, max_range = ref.split(':')
            else:
                min_range = max_range = ref

            min_col, min_row = coordinate_from_string(min_range)
            max_col, max_row = coordinate_from_string(max_range)

            return min_col, min_row, max_col, max_row

        if element.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c':
            # Supposedly the dimension is mandatory, but in practice it can be
            # left off sometimes, if so, observe the max/min extants and return
            # those instead.
            col, row = coordinate_from_string(element.get('r'))
            if smin_row is None:
                #initialize the observed max/min values
                smin_col = smax_col = col
                smin_row = smax_row = row
            else:
                #Keep track of the seen max and min (fallback if there's no dimension)
                smin_col = min(smin_col, col)
                smin_row = min(smin_row, row)
                smax_col = max(smax_col, col)
                smax_row = max(smax_row, row)
        else:
            element.clear()

    return smin_col, smin_row, smax_col, smax_row

def filter_cells(item):
    return item[1].tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'

def fast_parse(ws, xml_source, string_table, style_table):

    source = _get_xml_iter(xml_source)

    it = iterparse(source)

    for event, element in ifilter(filter_cells, it):

        value = element.findtext('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')

        if value is not None:

            coordinate = element.get('r')
            data_type = element.get('t', 'n')
            style_id = element.get('s')

            if data_type == Cell.TYPE_STRING:
                value = string_table.get(int(value))

            ws.cell(coordinate).value = value

            if style_id is not None:
                ws._styles[coordinate] = style_table.get(int(style_id))

        # to avoid memory exhaustion, clear the item after use
        element.clear()

from legacy_openpyxl.reader.iter_worksheet import IterableWorksheet

def read_worksheet(xml_source, parent, preset_title, string_table,
                   style_table, workbook_name=None, sheet_codename=None):
    """Read an xml worksheet"""
    if workbook_name and sheet_codename:
        ws = IterableWorksheet(parent, preset_title, workbook_name,
                sheet_codename, xml_source, string_table)
    else:
        ws = Worksheet(parent, preset_title)
        fast_parse(ws, xml_source, string_table, style_table)
    return ws

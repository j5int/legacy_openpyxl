# Copyright (c) 2010-2014 openpyxl
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

# Python stdlib imports
import datetime
from functools import partial

import pytest

# compatibility imports
from legacy_openpyxl.compat import BytesIO

# package imports
from legacy_openpyxl.reader.excel import load_workbook
from legacy_openpyxl.reader.style import read_style_table
from legacy_openpyxl.writer.excel import save_virtual_workbook

from legacy_openpyxl.styles import (
    NumberFormat,
    Color,
    Font,
    PatternFill,
    Borders,
    Protection,
    Style
)
from legacy_openpyxl.styles.border import Border
from legacy_openpyxl.styles import colors
from legacy_openpyxl.styles import fills

# test imports
from legacy_openpyxl.tests.helper import get_xml, compare_xml
from legacy_openpyxl.styles.alignment import Alignment


@pytest.fixture
def StyleReader():
    from ..style import SharedStylesParser
    return SharedStylesParser


@pytest.mark.parametrize("value, expected",
                         [
                             ({'indexed': '62'}, "00333399"),
                             ({'rgb': "FFFFFFFF"}, "FFFFFFFF"),
                             ({'theme': '0'}, 'theme:0:'),
                             ({'theme': '0', 'tint': "0.5"}, "theme:0:0.5")
])
def test_get_color(StyleReader, value, expected):
    reader = StyleReader("""
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></styleSheet>
    """)
    assert reader._get_relevant_color(value) == expected


def test_read_fills(StyleReader,datadir):
    datadir.chdir()
    expected = [
        PatternFill(),
        PatternFill(fill_type='gray125'),
        PatternFill(fill_type='solid',
             start_color=Color('theme:0:-0.14999847407452621'),
             end_color=Color('System Background')
             ),
        PatternFill(fill_type='solid',
             start_color=Color('theme:0:'),
             end_color=Color('System Background')
             ),
        PatternFill(fill_type='solid',
             start_color=Color("00333399"),
             end_color=Color('System Background')
             )
    ]
    with open("bug311-styles.xml") as src:
        reader = StyleReader(src.read())
        assert list(reader.parse_fills()) == expected


def test_read_cell_style(datadir):
    datadir.chdir()
    with open("empty-workbook-styles.xml") as content:
        style_properties = read_style_table(content.read())
        style_table = style_properties['table']
        assert len(style_table) == 2


def test_read_style(datadir):
    datadir.chdir()
    with open("simple-styles.xml") as content:
        style_properties = read_style_table(content.read())
        style_table = style_properties['table']
        style_list = style_properties['list']
        assert len(style_table) == 4
        assert NumberFormat._BUILTIN_FORMATS[9] == style_list[style_table[1]].number_format
        assert 'yyyy-mm-dd' == style_list[style_table[2]].number_format


def test_read_complex_style(datadir):
    datadir.chdir()
    wb = load_workbook("complex-styles.xlsx")
    ws = wb.get_active_sheet()
    assert ws.column_dimensions['A'].width == 31.1640625

    style = partial(ws.get_style)
    assert style('I').fill.start_color.index == 'FF006600'
    assert style('I').font.color.index == 'FF3300FF'
    assert style('A2').font.name == 'Arial'
    assert style('A2').font.size == 10
    assert not style('A2').font.bold
    assert not style('A2').font.italic
    assert style('A3').font.name == 'Arial'
    assert style('A3').font.size == 12
    assert style('A3').font.bold
    assert not style('A3').font.italic
    assert style('A4').font.name == 'Arial'
    assert style('A4').font.size == 14
    assert not style('A4').font.bold
    assert style('A4').font.italic
    assert style('A5').font.color.index == 'FF3300FF'
    assert style('A6').font.color.index == 'theme:9:'
    assert style('A7').fill.start_color.index == 'FFFFFF66'
    assert style('A8').fill.start_color.index == 'theme:8:'
    assert style('A9').alignment.horizontal == 'left'
    assert style('A10').alignment.horizontal == 'right'
    assert style('A11').alignment.horizontal == 'center'
    assert style('A12').alignment.vertical == 'top'
    assert style('A13').alignment.vertical == 'center'
    assert style('A14').alignment.vertical == 'bottom'
    assert style('A15').number_format == '0.00'
    assert style('A16').number_format == 'mm-dd-yy'
    assert style('A17').number_format == '0.00%'
    assert 'A18:B18' in ws._merged_cells
    assert ws.cell('B18').merged
    assert style('A19').borders.top.color.index == 'FF006600'
    assert style('A19').borders.bottom.color.index == 'FF006600'
    assert style('A19').borders.left.color.index == 'FF006600'
    assert style('A19').borders.right.color.index == 'FF006600'
    assert style('A21').borders.top.color.index == 'theme:7:'
    assert style('A21').borders.bottom.color.index == 'theme:7:'
    assert style('A21').borders.left.color.index == 'theme:7:'
    assert style('A21').borders.right.color.index == 'theme:7:'
    assert style('A23').fill.start_color.index == 'FFCCCCFF'
    assert style('A23').borders.top.color.index == 'theme:6:'
    assert 'A23:B24' in ws._merged_cells
    assert ws.cell('A24').merged
    assert ws.cell('B23').merged
    assert ws.cell('B24').merged
    assert style('A25').alignment.wrap_text
    assert style('A26').alignment.shrink_to_fit


def test_change_existing_styles(datadir):
    wb = load_workbook("complex-styles.xlsx")
    ws = wb.get_active_sheet()

    ws.column_dimensions['A'].width = 20
    i_style = ws.get_style('I')
    ws.set_style('I', i_style.copy(fill=PatternFill(fill_type='solid',
                                             start_color=Color('FF442200')),
                                   font=Font(color=Color('FF002244'))))
    ws.cell('A2').style = ws.cell('A2').style.copy(font=Font(name='Times New Roman',
                                                             size=12,
                                                             bold=True,
                                                             italic=True))
    ws.cell('A3').style = ws.cell('A3').style.copy(font=Font(name='Times New Roman',
                                                             size=14,
                                                             bold=False,
                                                             italic=True))
    ws.cell('A4').style = ws.cell('A4').style.copy(font=Font(name='Times New Roman',
                                                             size=16,
                                                             bold=True,
                                                             italic=False))
    ws.cell('A5').style = ws.cell('A5').style.copy(font=Font(color=Color('FF66FF66')))
    ws.cell('A6').style = ws.cell('A6').style.copy(font=Font(color=Color('theme:1:')))
    ws.cell('A7').style = ws.cell('A7').style.copy(fill=PatternFill(fill_type='solid',
                                                             start_color=Color('FF330066')))
    ws.cell('A8').style = ws.cell('A8').style.copy(fill=PatternFill(fill_type='solid',
                                                             start_color=Color('theme:2:')))
    ws.cell('A9').style = ws.cell('A9').style.copy(alignment=Alignment(horizontal='center'))
    ws.cell('A10').style = ws.cell('A10').style.copy(alignment=Alignment(horizontal='left'))
    ws.cell('A11').style = ws.cell('A11').style.copy(alignment=Alignment(horizontal='right'))
    ws.cell('A12').style = ws.cell('A12').style.copy(alignment=Alignment(vertical='bottom'))
    ws.cell('A13').style = ws.cell('A13').style.copy(alignment=Alignment(vertical='top'))
    ws.cell('A14').style = ws.cell('A14').style.copy(alignment=Alignment(vertical='center'))
    ws.cell('A15').style = ws.cell('A15').style.copy(number_format=NumberFormat('0.00%'))
    ws.cell('A16').style = ws.cell('A16').style.copy(number_format=NumberFormat('0.00'))
    ws.cell('A17').style = ws.cell('A17').style.copy(number_format=NumberFormat('mm-dd-yy'))
    ws.unmerge_cells('A18:B18')
    ws.cell('A19').style = ws.cell('A19').style.copy(borders=Borders(top=Border(border_style=Border.BORDER_THIN,
                                                                                color=Color('FF006600')),
                                                                     bottom=Border(border_style=Border.BORDER_THIN,
                                                                                   color=Color('FF006600')),
                                                                     left=Border(border_style=Border.BORDER_THIN,
                                                                                 color=Color('FF006600')),
                                                                     right=Border(border_style=Border.BORDER_THIN,
                                                                                  color=Color('FF006600'))))
    ws.cell('A21').style = ws.cell('A21').style.copy(borders=Borders(top=Border(border_style=Border.BORDER_THIN,
                                                                                color=Color('theme:7:')),
                                                                     bottom=Border(border_style=Border.BORDER_THIN,
                                                                                   color=Color('theme:7:')),
                                                                     left=Border(border_style=Border.BORDER_THIN,
                                                                                 color=Color('theme:7:')),
                                                                     right=Border(border_style=Border.BORDER_THIN,
                                                                                  color=Color('theme:7:'))))
    ws.cell('A23').style = ws.cell('A23').style.copy(borders=Borders(top=Border(border_style=Border.BORDER_THIN,
                                                                                color=Color('theme:6:'))),
                                                     fill=PatternFill(fill_type='solid',
                                                               start_color=Color('FFCCCCFF')))
    ws.unmerge_cells('A23:B24')
    ws.cell('A25').style = ws.cell('A25').style.copy(alignment=Alignment(wrap_text=False))
    ws.cell('A26').style = ws.cell('A26').style.copy(alignment=Alignment(shrink_to_fit=False))

    saved_wb = save_virtual_workbook(wb)
    new_wb = load_workbook(BytesIO(saved_wb))
    ws = new_wb.get_active_sheet()

    assert ws.column_dimensions['A'].width == 20.0

    style = partial(ws.get_style)

    assert ws.get_style('I').fill.start_color.index == 'FF442200'
    assert ws.get_style('I').font.color.index == 'FF002244'
    assert style('A2').font.name == 'Times New Roman'
    assert style('A2').font.size == 12
    assert style('A2').font.bold
    assert style('A2').font.italic
    assert style('A3').font.name == 'Times New Roman'
    assert style('A3').font.size == 14
    assert not style('A3').font.bold
    assert style('A3').font.italic
    assert style('A4').font.name == 'Times New Roman'
    assert style('A4').font.size == 16
    assert style('A4').font.bold
    assert not style('A4').font.italic
    assert style('A5').font.color.index == 'FF66FF66'
    assert style('A6').font.color.index == 'theme:1:'
    assert style('A7').fill.start_color.index == 'FF330066'
    assert style('A8').fill.start_color.index == 'theme:2:'
    assert style('A9').alignment.horizontal == 'center'
    assert style('A10').alignment.horizontal == 'left'
    assert style('A11').alignment.horizontal == 'right'
    assert style('A12').alignment.vertical == 'bottom'
    assert style('A13').alignment.vertical == 'top'
    assert style('A14').alignment.vertical == 'center'
    assert style('A15').number_format == '0.00%'
    assert style('A16').number_format == '0.00'
    assert style('A17').number_format == 'mm-dd-yy'
    assert 'A18:B18' not in ws._merged_cells
    assert not ws.cell('B18').merged
    assert style('A19').borders.top.color.index == 'FF006600'
    assert style('A19').borders.bottom.color.index == 'FF006600'
    assert style('A19').borders.left.color.index == 'FF006600'
    assert style('A19').borders.right.color.index == 'FF006600'
    assert style('A21').borders.top.color.index == 'theme:7:'
    assert style('A21').borders.bottom.color.index == 'theme:7:'
    assert style('A21').borders.left.color.index == 'theme:7:'
    assert style('A21').borders.right.color.index == 'theme:7:'
    assert style('A23').fill.start_color.index == 'FFCCCCFF'
    assert style('A23').borders.top.color.index == 'theme:6:'
    assert 'A23:B24' not in ws._merged_cells
    assert not ws.cell('A24').merged
    assert not ws.cell('B23').merged
    assert not ws.cell('B24').merged
    assert not style('A25').alignment.wrap_text
    assert not style('A26').alignment.shrink_to_fit

    # Verify that previously duplicate styles remain the same
    assert ws.column_dimensions['C'].width == 31.1640625
    assert style('C2').font.name == 'Arial'
    assert style('C2').font.size == 10
    assert not style('C2').font.bold
    assert not style('C2').font.italic
    assert style('C3').font.name == 'Arial'
    assert style('C3').font.size == 12
    assert style('C3').font.bold
    assert not style('C3').font.italic
    assert style('C4').font.name == 'Arial'
    assert style('C4').font.size == 14
    assert not style('C4').font.bold
    assert style('C4').font.italic
    assert style('C5').font.color.index == 'FF3300FF'
    assert style('C6').font.color.index == 'theme:9:'
    assert style('C7').fill.start_color.index == 'FFFFFF66'
    assert style('C8').fill.start_color.index == 'theme:8:'
    assert style('C9').alignment.horizontal == 'left'
    assert style('C10').alignment.horizontal == 'right'
    assert style('C11').alignment.horizontal == 'center'
    assert style('C12').alignment.vertical == 'top'
    assert style('C13').alignment.vertical == 'center'
    assert style('C14').alignment.vertical == 'bottom'
    assert style('C15').number_format == '0.00'
    assert style('C16').number_format == 'mm-dd-yy'
    assert style('C17').number_format == '0.00%'
    assert 'C18:D18' in ws._merged_cells
    assert ws.cell('D18').merged
    assert style('C19').borders.top.color.index == 'FF006600'
    assert style('C19').borders.bottom.color.index == 'FF006600'
    assert style('C19').borders.left.color.index == 'FF006600'
    assert style('C19').borders.right.color.index == 'FF006600'
    assert style('C21').borders.top.color.index == 'theme:7:'
    assert style('C21').borders.bottom.color.index == 'theme:7:'
    assert style('C21').borders.left.color.index == 'theme:7:'
    assert style('C21').borders.right.color.index == 'theme:7:'
    assert style('C23').fill.start_color.index == 'FFCCCCFF'
    assert style('C23').borders.top.color.index == 'theme:6:'
    assert 'C23:D24' in ws._merged_cells
    assert ws.cell('C24').merged
    assert ws.cell('D23').merged
    assert ws.cell('D24').merged
    assert style('C25').alignment.wrap_text
    assert style('C26').alignment.shrink_to_fit


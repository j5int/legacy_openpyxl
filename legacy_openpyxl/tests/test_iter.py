# file openpyxl/tests/test_iter.py

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

from nose.tools import eq_, raises, assert_raises
import os.path as osp
from legacy_openpyxl.tests.helper import DATADIR
from legacy_openpyxl.reader.iter_worksheet import read_worksheet, get_range_boundaries
from legacy_openpyxl.reader.excel import load_workbook

workbook_name = osp.join(DATADIR, 'genuine', 'empty.xlsx')
sheet_name = 'Sheet1 - Text'

expected = [['This is cell A1 in Sheet 1', '', '', '', '', '', ''],
            ['', '', '', '', '', '', ''],
            ['', '', '', '', '', '', ''],
            ['', '', '', '', '', '', ''],
            ['', '', '', '', '', '', 'This is cell G5'], ]

def test_read_fast():

    for row, expected_row in zip(read_worksheet(workbook_name, sheet_name), expected):

        row_values = [x.internal_value for x in row]

        eq_(row_values, expected_row)

def test_read_fast_integrated():

    wb = load_workbook(filename = workbook_name, use_iterators = True)
    ws = wb.get_sheet_by_name(name = sheet_name)

    for row, expected_row in zip(ws.iter_rows(), expected):

        row_values = [x.internal_value for x in row]

        eq_(row_values, expected_row)


def test_get_boundaries_range():
    
    eq_(get_range_boundaries('C1:C4'), (3,1,3,4))

def test_get_boundaries_one():


    eq_(get_range_boundaries('C1'), (3,1,4,1))

def test_read_single_cell_range():

    wb = load_workbook(filename = workbook_name, use_iterators = True)
    ws = wb.get_sheet_by_name(name = sheet_name)

    eq_('This is cell A1 in Sheet 1', list(ws.iter_rows('A1'))[0][0].internal_value)

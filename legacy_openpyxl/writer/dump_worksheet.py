from __future__ import absolute_import
from legacy_openpyxl.comments.comments import Comment
from legacy_openpyxl.writer.comments import CommentWriter
from legacy_openpyxl.writer.worksheet import write_worksheet_rels
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

"""Write worksheets to xml representations in an optimized way"""

import datetime
import os
from tempfile import NamedTemporaryFile

from legacy_openpyxl.compat import OrderedDict

from legacy_openpyxl.cell import  get_column_letter, Cell, TIME_TYPES
from legacy_openpyxl.worksheet import Worksheet
from legacy_openpyxl.xml.functions import (
    XMLGenerator,
    start_tag,
    end_tag, tag
)
from legacy_openpyxl.date_time import (
    to_excel,
    timedelta_to_days,
    time_to_days
)
from legacy_openpyxl.xml.constants import MAX_COLUMN, MAX_ROW, PACKAGE_XL
from legacy_openpyxl.units import NUMERIC_TYPES
from legacy_openpyxl.exceptions import WorkbookAlreadySaved
from legacy_openpyxl.writer.excel import ExcelWriter
from legacy_openpyxl.writer.strings import write_string_table
from legacy_openpyxl.writer.styles import StyleWriter
from legacy_openpyxl.styles import Style, NumberFormat, DEFAULTS

from legacy_openpyxl.xml.constants import (ARC_SHARED_STRINGS, PACKAGE_WORKSHEETS)

ITERABLES = (list, tuple)


DATETIME_STYLE = Style(number_format=NumberFormat(format_code=NumberFormat.FORMAT_DATE_YYYYMMDD2))
STYLES = {'datetime': {'type': Cell.TYPE_NUMERIC,
                       'style': DATETIME_STYLE},
          'string': {'type': Cell.TYPE_STRING,
                     'style': DEFAULTS},
          'numeric': {'type': Cell.TYPE_NUMERIC,
                      'style': DEFAULTS},
          'formula': {'type': Cell.TYPE_FORMULA,
                      'style': DEFAULTS},
          'boolean': {'type': Cell.TYPE_BOOL,
                    'style': DEFAULTS},
        }

DESCRIPTORS_CACHE_SIZE = 50
BOUNDING_BOX_PLACEHOLDER = 'A1:%s%d' % (get_column_letter(MAX_COLUMN), MAX_ROW)


class CommentParentCell(object):
    __slots__ = ('coordinate', 'row', 'column')

    def __init__(self, coordinate, row, column):
        self.coordinate = coordinate
        self.row = row
        self.column = column


def create_temporary_file(suffix=''):
    fobj = NamedTemporaryFile(mode='w+', suffix=suffix,
                              prefix='openpyxl.', delete=False)
    filename = fobj.name
    return filename


class DumpWorksheet(Worksheet):
    """
    .. warning::

        You shouldn't initialize this yourself, use :class:`openpyxl.workbook.Workbook` constructor instead,
        with `optimized_write = True`.
    """
    def __init__(self, parent_workbook, title):
        Worksheet.__init__(self, parent_workbook, title)

        self._max_col = 0
        self._max_row = 0
        self._parent = parent_workbook

        self._fileobj_header_name = create_temporary_file(suffix='.header')
        self._fileobj_content_name = create_temporary_file(suffix='.content')
        self._fileobj_name = create_temporary_file()

        self._strings = self._parent.shared_strings
        self._styles = self.parent.shared_styles
        self._comments = []

    def get_temporary_file(self, filename):
        if filename in self._descriptors_cache:
            fobj = self._descriptors_cache[filename]
            # re-insert the value so it does not get evicted
            # from cache soon
            del self._descriptors_cache[filename]
            self._descriptors_cache[filename] = fobj
            return fobj
        else:
            if filename is None:
                raise WorkbookAlreadySaved('this workbook has already been saved '
                        'and cannot be modified or saved anymore.')

            fobj = open(filename, 'r+')
            self._descriptors_cache[filename] = fobj
            if len(self._descriptors_cache) > DESCRIPTORS_CACHE_SIZE:
                filename, fileobj = self._descriptors_cache.popitem(last=False)
                fileobj.close()
            return fobj

    @property
    def _descriptors_cache(self):
        try:
            return self._parent._local_data.cache
        except AttributeError:
            self._parent._local_data.cache = OrderedDict()
            return self._parent._local_data.cache

    @property
    def filename(self):
        return self._fileobj_name

    @property
    def _temp_files(self):
        return (self._fileobj_content_name,
                self._fileobj_header_name,
                self._fileobj_name)

    def _unset_temp_files(self):
        self._fileobj_header_name = None
        self._fileobj_content_name = None
        self._fileobj_name = None

    def write_header(self):

        fobj = self.get_temporary_file(filename=self._fileobj_header_name)
        doc = XMLGenerator(fobj, 'utf-8')

        start_tag(doc, 'worksheet',
                {
                'xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'})
        start_tag(doc, 'sheetPr')
        tag(doc, 'outlinePr',
                {'summaryBelow': '1',
                'summaryRight': '1'})
        end_tag(doc, 'sheetPr')
        tag(doc, 'dimension', {'ref': 'A1:%s' % (self.get_dimensions())})
        start_tag(doc, 'sheetViews')
        start_tag(doc, 'sheetView', {'workbookViewId': '0'})
        tag(doc, 'selection', {'activeCell': 'A1',
                'sqref': 'A1'})
        end_tag(doc, 'sheetView')
        end_tag(doc, 'sheetViews')
        tag(doc, 'sheetFormatPr', {'defaultRowHeight': '15'})
        start_tag(doc, 'sheetData')

    def close(self):
        self._close_content()
        self._fileobj = self.get_temporary_file(filename=self._fileobj_name)
        self._write_fileobj(self._fileobj_header_name)
        self._write_fileobj(self._fileobj_content_name)
        self._fileobj.close()

    def _write_fileobj(self, fobj_name):
        fobj = self.get_temporary_file(filename=fobj_name)
        fobj.flush()
        fobj.seek(0)

        while True:
            chunk = fobj.read(4096)
            if not chunk:
                break
            self._fileobj.write(chunk)

        fobj.close()
        self._fileobj.flush()

    def _close_content(self):
        doc = self._get_content_generator()
        if self._comments:
            tag(doc, 'legacyDrawing', {'r:id': 'commentsvml'})
        end_tag(doc, 'sheetData')
        end_tag(doc, 'worksheet')

    def get_dimensions(self):
        if not self._max_col or not self._max_row:
            return 'A1'
        else:
            return '%s%d' % (get_column_letter(self._max_col), (self._max_row))

    def _get_content_generator(self):
        """ XXX: this is ugly, but it allows to resume writing the file
        even after the handle is closed"""

        # when I'll recreate the XMLGenerator, it will start writing at the
        # begining of the file, erasing previously entered rows, so we have
        # to move to the end of the file before adding new tags
        handle = self.get_temporary_file(filename=self._fileobj_content_name)
        handle.seek(0, 2)

        return XMLGenerator(out=handle)

    def append(self, row):
        """
        :param row: iterable containing values to append
        :type row: iterable
        """
        doc = self._get_content_generator()
        self._max_row += 1
        span = len(row)
        self._max_col = max(self._max_col, span)
        row_idx = self._max_row
        attrs = {'r': '%d' % row_idx,
                 'spans': '1:%d' % span}
        start_tag(doc, 'row', attrs)

        for col_idx, cell in enumerate(row):
            style = None
            comment = None
            if cell is None:
                continue
            elif isinstance(cell, dict):
                dct = cell
                cell = dct.get('value')
                if cell is None:
                    continue
                style = dct.get('style')
                comment = dct.get('comment')
                for ob, attr, cls in ((style, 'style', Style),
                                      (comment, 'comment', Comment)):
                    if ob is not None and not isinstance(ob, cls):
                        raise TypeError('%s should be a %s not a %s' %
                                        (attr,
                                         cls.__class__.__name__,
                                         ob.__class__.__name__))

            column = get_column_letter(col_idx + 1)
            coordinate = '%s%d' % (column, row_idx)
            attributes = {'r': coordinate}
            if comment is not None:
                comment._parent = CommentParentCell(coordinate,
                                                    row_idx,
                                                    column)
                self._comments.append(comment)
                self._comment_count += 1

            if isinstance(cell, bool):
                dtype = 'boolean'
            elif isinstance(cell, NUMERIC_TYPES):
                dtype = 'numeric'
            elif isinstance(cell, TIME_TYPES):
                dtype = 'datetime'
                if isinstance(cell, datetime.date):
                    cell = to_excel(cell)
                elif isinstance(cell, datetime.time):
                    cell = time_to_days(cell)
                elif isinstance(cell, datetime.timedelta):
                    cell = timedelta_to_days(cell)
                if style is None:
                    # allow user-defined style if needed
                    style = STYLES[dtype]['style']
            elif cell and cell[0] == '=':
                dtype = 'formula'
            else:
                dtype = 'string'
                cell = self._strings.add(cell)

            if style is not None:
                attributes['s'] = '%d' % self._styles.add(style)

            if dtype != 'formula':
                attributes['t'] = STYLES[dtype]['type']
            start_tag(doc, 'c', attributes)

            if dtype == 'formula':
                tag(doc, 'f', body='%s' % cell[1:])
                tag(doc, 'v')
            elif dtype == 'boolean':
                tag(doc, 'v', body='%d' % cell)
            else:
                tag(doc, 'v', body='%s' % cell)
            end_tag(doc, 'c')
        end_tag(doc, 'row')


def save_dump(workbook, filename):
    writer = ExcelDumpWriter(workbook)
    writer.save(filename)
    return True


class DumpCommentWriter(CommentWriter):
    def sheet_comments(self, sheet):
        return sheet._comments


class ExcelDumpWriter(ExcelWriter):
    def __init__(self, workbook):
        self.workbook = workbook
        self.style_writer = StyleWriter(workbook)

    def _write_string_table(self, archive):
        shared_strings = self.workbook.shared_strings
        archive.writestr(ARC_SHARED_STRINGS,
                write_string_table(shared_strings))

    def _write_worksheets(self, archive, style_writer):
        drawing_id = 1
        comments_id = 1

        for i, sheet in enumerate(self.workbook.worksheets):
            sheet.write_header()
            sheet.close()
            archive.write(sheet.filename, PACKAGE_WORKSHEETS + '/sheet%d.xml' % (i + 1))
            for filename in sheet._temp_files:
                del sheet._descriptors_cache[filename]
                os.remove(filename)
            sheet._unset_temp_files()

            # write comments
            if sheet._comments:
                archive.writestr(PACKAGE_WORKSHEETS +
                        '/_rels/sheet%d.xml.rels' % (i + 1),
                        write_worksheet_rels(sheet, drawing_id, comments_id))

                cw = DumpCommentWriter(sheet)
                archive.writestr(PACKAGE_XL + '/comments%d.xml' % comments_id,
                    cw.write_comments())
                archive.writestr(PACKAGE_XL + '/drawings/commentsDrawing%d.vml' % comments_id,
                    cw.write_comments_vml())
                comments_id += 1

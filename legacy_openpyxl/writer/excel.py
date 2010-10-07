# file openpyxl/writer/excel.py

"""Write a .xlsx file."""

# Python stdlib imports
from zipfile import ZipFile, ZIP_DEFLATED
from StringIO import StringIO

# package imports
from ..shared.ooxml import ARC_SHARED_STRINGS, ARC_CONTENT_TYPES, \
        ARC_ROOT_RELS, ARC_WORKBOOK_RELS, ARC_APP, ARC_CORE, ARC_THEME, \
        ARC_STYLE, ARC_WORKBOOK, PACKAGE_WORKSHEETS
from .strings import create_string_table, write_string_table
from .workbook import write_content_types, write_root_rels, \
        write_workbook_rels, write_properties_app, write_properties_core, \
        write_workbook
from .theme import write_theme
from .styles import create_style_table, write_style_table
from .worksheet import write_worksheet, write_worksheet_rels


class ExcelWriter(object):
    """Write a workbook object to an Excel file."""

    def __init__(self, workbook):
        self.workbook = workbook

    def write_data(self, archive):
        """Write the various xml files into the zip archive."""
        # cleanup all worksheets
        for ws in self.workbook.worksheets:
            ws.garbage_collect()
        shared_string_table = create_string_table(self.workbook)
        archive.writestr(ARC_SHARED_STRINGS,
                write_string_table(shared_string_table))
        shared_style_table = create_style_table(self.workbook)
        archive.writestr(ARC_CONTENT_TYPES, write_content_types(self.workbook))
        archive.writestr(ARC_ROOT_RELS, write_root_rels(self.workbook))
        archive.writestr(ARC_WORKBOOK_RELS, write_workbook_rels(self.workbook))
        archive.writestr(ARC_APP, write_properties_app(self.workbook))
        archive.writestr(ARC_CORE,
                write_properties_core(self.workbook.properties))
        archive.writestr(ARC_THEME, write_theme())
        archive.writestr(ARC_STYLE, write_style_table(shared_style_table))
        archive.writestr(ARC_WORKBOOK, write_workbook(self.workbook))
        style_id_by_hash = dict([(style.__crc__(), style_id) for
                style, style_id in shared_style_table.iteritems()])
        for i, sheet in enumerate(self.workbook.worksheets):
            archive.writestr(PACKAGE_WORKSHEETS + '/sheet%d.xml' % (i + 1),
                    write_worksheet(sheet, shared_string_table,
                            style_id_by_hash))
            if sheet.relationships:
                archive.writestr(PACKAGE_WORKSHEETS +
                        '/_rels/sheet%d.xml.rels' % (i + 1),
                        write_worksheet_rels(sheet))

    def save(self, filename):
        """Write data into the archive."""
        try:
            archive = ZipFile(filename, 'w', ZIP_DEFLATED)
            self.write_data(archive)
        finally:
            try:
                archive.close()
            except UnboundLocalError:
                pass


def save_workbook(workbook, filename):
    """Save the given workbook on the filesystem under the name filename.

    :param workbook: the workbook to save
    :type workbook: :class:`openpyxl.workbook.Workbook`

    :param filename: the path to which save the workbook
    :type filename: string

    :rtype: bool

    """
    writer = ExcelWriter(workbook)
    writer.save(filename)
    return True


def save_virtual_workbook(workbook):
    """Return an in-memory workbook, suitable for a Django response."""
    writer = ExcelWriter(workbook)
    temp_buffer = StringIO()
    try:
        archive = ZipFile(temp_buffer, 'w', ZIP_DEFLATED)
        writer.write_data(archive)
    finally:
        archive.close()
    virtual_workbook = temp_buffer.getvalue()
    temp_buffer.close()
    return virtual_workbook

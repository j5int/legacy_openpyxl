# file openpyxl/writer/workbook.py

from legacy_openpyxl.shared.xmltools import ElementTree, Element, SubElement

from legacy_openpyxl.cell import absolute_coordinate
from legacy_openpyxl.shared.xmltools import get_document_content
from legacy_openpyxl.shared.ooxml import NAMESPACES, ARC_CORE, ARC_WORKBOOK, ARC_APP, ARC_THEME, ARC_STYLE, ARC_SHARED_STRINGS
from legacy_openpyxl.shared.date_time import datetime_to_W3CDTF

def write_properties_core(properties):

    root = Element('cp:coreProperties', {'xmlns:cp': NAMESPACES['cp'],
                                         'xmlns:dc': NAMESPACES['dc'],
                                         'xmlns:dcterms': NAMESPACES['dcterms'],
                                         'xmlns:dcmitype': NAMESPACES['dcmitype'],
                                         'xmlns:xsi': NAMESPACES['xsi']})

    SubElement(root, 'dc:creator').text = properties.creator
    SubElement(root, 'cp:lastModifiedBy').text = properties.last_modified_by

    SubElement(root, 'dcterms:created', {'xsi:type': 'dcterms:W3CDTF'}).text = datetime_to_W3CDTF(properties.created)
    SubElement(root, 'dcterms:modified', {'xsi:type': 'dcterms:W3CDTF'}).text = datetime_to_W3CDTF(properties.modified)

    return get_document_content(root)


def write_content_types(workbook):

    root = Element('Types', {'xmlns' : "http://schemas.openxmlformats.org/package/2006/content-types"})

    SubElement(root, 'Override', {'PartName' : '/' + ARC_THEME,
                                  'ContentType' : 'application/vnd.openxmlformats-officedocument.theme+xml'})
    SubElement(root, 'Override', {'PartName' : '/' + ARC_STYLE,
                                  'ContentType' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'})
    SubElement(root, 'Default', {'Extension' : 'rels',
                                  'ContentType' : 'application/vnd.openxmlformats-package.relationships+xml'})
    SubElement(root, 'Default', {'Extension' : 'xml',
                                  'ContentType' : 'application/xml'})
    SubElement(root, 'Override', {'PartName' : '/' + ARC_WORKBOOK,
                                  'ContentType' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'})
    SubElement(root, 'Override', {'PartName' : '/' + ARC_APP,
                                  'ContentType' : 'application/vnd.openxmlformats-officedocument.extended-properties+xml'})
    SubElement(root, 'Override', {'PartName' : '/' + ARC_CORE,
                                  'ContentType' : 'application/vnd.openxmlformats-package.core-properties+xml'})
    SubElement(root, 'Override', {'PartName' : '/' + ARC_SHARED_STRINGS,
                                  'ContentType' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'})

    for sheet_id in xrange(len(workbook.worksheets)):
        SubElement(root, 'Override', {'PartName' : '/xl/worksheets/sheet%d.xml' % (sheet_id + 1),
                                      'ContentType' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'})

    return get_document_content(root)

def write_properties_app(workbook):

    worksheets_count = len(workbook.worksheets)


    root = Element('Properties', {'xmlns' : 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
                                  'xmlns:vt' : 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'})

    SubElement(root, 'Application').text = 'Microsoft Excel'
    SubElement(root, 'DocSecurity').text = '0'
    SubElement(root, 'ScaleCrop').text = 'false'
    SubElement(root, 'Company')

    SubElement(root, 'LinksUpToDate').text = 'false'
    SubElement(root, 'SharedDoc').text = 'false'
    SubElement(root, 'HyperlinksChanged').text = 'false'
    SubElement(root, 'AppVersion').text = '12.0000'

    # heading pairs part
    heading_pairs = SubElement(root, 'HeadingPairs')
    vector = SubElement(heading_pairs, 'vt:vector', {'size' : '2',
                                                     'baseType' : 'variant'})
    variant = SubElement(vector, 'vt:variant')
    SubElement(variant, 'vt:lpstr').text = 'Worksheets'

    variant = SubElement(vector, 'vt:variant')
    SubElement(variant, 'vt:i4').text = '%d' % worksheets_count

    # title of parts
    title_of_parts = SubElement(root, 'TitlesOfParts')
    vector = SubElement(title_of_parts, 'vt:vector', {'size' : '%d' % worksheets_count,
                                                     'baseType' : 'lpstr'})

    for ws in workbook.worksheets:
        SubElement(vector, 'vt:lpstr').text = '%s' % ws.title

    return get_document_content(root)

def write_root_rels(workbook):

    root = Element('Relationships', {'xmlns' : "http://schemas.openxmlformats.org/package/2006/relationships"})

    SubElement(root, 'Relationship', {'Id' : 'rId1',
                                      'Type' : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
                                      'Target' : ARC_WORKBOOK})
    SubElement(root, 'Relationship', {'Id' : 'rId2',
                                      'Type' : 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
                                      'Target' : ARC_CORE})
    SubElement(root, 'Relationship', {'Id' : 'rId3',
                                      'Type' : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
                                      'Target' : ARC_APP})

    return get_document_content(root)

def write_workbook(workbook):

    root = Element('workbook', {'xmlns' : 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                                'xml:space' : 'preserve',
                                'xmlns:r' : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'})

    # file version
    SubElement(root, 'fileVersion', {'appName' : 'xl',
                                     'lastEdited' : '4',
                                     'lowestEdited' : '4',
                                     'rupBuild' : '4505'})

    # workbook pr
    SubElement(root, 'workbookPr', {'defaultThemeVersion' : '124226',
                                    'codeName' : 'ThisWorkbook'})

    # book views
    book_views = SubElement(root, 'bookViews')
    SubElement(book_views, 'workbookView', {'activeTab' : '%d' % workbook.get_index(workbook.get_active_sheet()),
                                   'autoFilterDateGrouping' : '1',
                                   'firstSheet' : '0',
                                   'minimized' : '0',
                                   'showHorizontalScroll' : '1',
                                   'showSheetTabs' : '1',
                                   'showVerticalScroll' : '1',
                                   'tabRatio' : '600',
                                   'visibility' : 'visible'})

    # worksheets
    sheets = SubElement(root, 'sheets')
    for i, sheet in enumerate(workbook.worksheets):
        nd_sheet = SubElement(sheets, 'sheet', {'name' : sheet.title,
                                                'sheetId' : '%d' % (i + 1),
                                                'r:id' : 'rId%d' % (i + 1)})

        if not sheet.sheet_state == sheet.SHEETSTATE_VISIBLE:
            nd_sheet.set('state', sheet.sheet_state)

    # named ranges
    defined_names = SubElement(root, 'definedNames')
    for named_range in workbook.get_named_ranges():
        name = SubElement(defined_names, 'definedName', {'name' : named_range.name})
        if named_range.local_only:
            name.set('localSheetId', workbook.get_index(named_range.worksheet))

        name.text = "'%s'!%s" % (named_range.worksheet.title.replace("'", "''"),
                                 absolute_coordinate(named_range.range))



    # calc pr
    SubElement(root, 'calcPr', {'calcId' : '124519',
                                'calcMode' : 'auto',
                                'fullCalcOnLoad' : '1'})

    return get_document_content(root)

def write_workbook_rels(workbook):

    root = Element('Relationships', {'xmlns' : 'http://schemas.openxmlformats.org/package/2006/relationships'})

    for i, sheet in enumerate(workbook.worksheets):

        SubElement(root, 'Relationship', {'Id' : 'rId%d' % (i + 1),
                                          'Type' : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                                          'Target' : 'worksheets/sheet%s.xml' % (i + 1)})

    rid = len(workbook.worksheets) + 1

    SubElement(root, 'Relationship', {'Id' : 'rId%d' % rid,
                                     'Type' : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
                                     'Target' : 'sharedStrings.xml'})

    SubElement(root, 'Relationship', {'Id' : 'rId%d' % (rid + 1),
                                      'Type' : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
                                      'Target' : 'styles.xml'})

    SubElement(root, 'Relationship', {'Id' : 'rId%d' % (rid + 2),
                                      'Type' : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
                                      'Target' : 'theme/theme1.xml'})

    return get_document_content(root)
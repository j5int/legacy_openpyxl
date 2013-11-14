# coding=UTF-8

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

from legacy_openpyxl.shared.xmltools import Element, SubElement, get_document_content
from legacy_openpyxl.shared.ooxml import (
    SHEET_MAIN_NS,
    DRAWING_NS,
    SHEET_DRAWING_NS,
    CHART_NS,
    REL_NS,
    CHART_DRAWING_NS,
    PKG_REL_NS
)


class DrawingWriter(object):
    """ one main drawing file per sheet """

    def __init__(self, sheet):
        self._sheet = sheet

    def write(self):
        """ write drawings for one sheet in one file """

        root = Element("{%s}wsDr" % SHEET_DRAWING_NS)

        for idx, chart in enumerate(self._sheet._charts):
            self._write_chart(root, chart, idx)

        for idx, img in enumerate(self._sheet._images):
            self._write_image(root, img, idx)

        return get_document_content(root)

    def _write_chart(self, node, chart, idx):
        """Add a chart"""
        drawing = chart.drawing

        #anchor = SubElement(root, 'xdr:twoCellAnchor')
        #(start_row, start_col), (end_row, end_col) = drawing.coordinates
        ## anchor coordinates
        #_from = SubElement(anchor, 'xdr:from')
        #x = SubElement(_from, 'xdr:col').text = str(start_col)
        #x = SubElement(_from, 'xdr:colOff').text = '0'
        #x = SubElement(_from, 'xdr:row').text = str(start_row)
        #x = SubElement(_from, 'xdr:rowOff').text = '0'

        #_to = SubElement(anchor, 'xdr:to')
        #x = SubElement(_to, 'xdr:col').text = str(end_col)
        #x = SubElement(_to, 'xdr:colOff').text = '0'
        #x = SubElement(_to, 'xdr:row').text = str(end_row)
        #x = SubElement(_to, 'xdr:rowOff').text = '0'

        # we only support absolute anchor atm (TODO: oneCellAnchor, twoCellAnchor
        x, y, w, h = drawing.get_emu_dimensions()
        anchor = SubElement(node, '{%s}absoluteAnchor' % SHEET_DRAWING_NS)
        SubElement(anchor, '{%s}pos' % SHEET_DRAWING_NS, {'x':str(x), 'y':str(y)})
        SubElement(anchor, '{%s}ext' % SHEET_DRAWING_NS, {'cx':str(w), 'cy':str(h)})

        # graph frame
        frame = SubElement(anchor, '{%s}graphicFrame' % SHEET_DRAWING_NS, {'macro':''})

        name = SubElement(frame, '{%s}nvGraphicFramePr' % SHEET_DRAWING_NS)
        SubElement(name, '{%s}cNvPr'% SHEET_DRAWING_NS, {'id':'%s' % idx, 'name':'Graphique %s' % idx})
        SubElement(name, '{%s}cNvGraphicFramePr' % SHEET_DRAWING_NS)

        frm = SubElement(frame, '{%s}xfrm'  % SHEET_DRAWING_NS)
        # no transformation
        SubElement(frm, '{%s}off' % DRAWING_NS, {'x':'0', 'y':'0'})
        SubElement(frm, '{%s}ext' % DRAWING_NS, {'cx':'0', 'cy':'0'})

        graph = SubElement(frame, '{%s}graphic' % DRAWING_NS)
        data = SubElement(graph, '{%s}graphicData' % DRAWING_NS, {'uri':CHART_NS})
        SubElement(data, '{%s}chart' % CHART_NS, {'{%s}id' % REL_NS:'rId%s' % (idx + 1)})

        SubElement(anchor, '{%s}clientData' % SHEET_DRAWING_NS)
        return node

    def _write_image(self, node, img, idx):
        drawing = img.drawing

        x, y, w, h = drawing.get_emu_dimensions()
        anchor = SubElement(node, '{%s}absoluteAnchor' % SHEET_DRAWING_NS)
        SubElement(anchor, '{%s}pos' % SHEET_DRAWING_NS, {'x':str(x), 'y':str(y)})
        SubElement(anchor, '{%s}ext' % SHEET_DRAWING_NS, {'cx':str(w), 'cy':str(h)})

        pic = SubElement(anchor, '{%s}pic' % SHEET_DRAWING_NS)
        name = SubElement(pic, '{%s}nvPicPr' % SHEET_DRAWING_NS)
        SubElement(name, '{%s}cNvPr' % SHEET_DRAWING_NS, {'id':'%s' % idx, 'name':'Picture %s' % idx})
        SubElement(SubElement(name, '{%s}cNvPicPr' % SHEET_DRAWING_NS),
                   '{%s}picLocks' % DRAWING_NS, {'noChangeAspect':"1" if img.nochangeaspect\
                                    else '0','noChangeArrowheads':"1" if img.nochangearrowheads\
                                    else '0'})
        blipfill = SubElement(pic, '{%s}blipFill' % SHEET_DRAWING_NS)
        SubElement(blipfill, '{%s}blip' % DRAWING_NS, {
            '{%s}embed' % REL_NS: 'rId%s' % (idx + 1),
            'cstate':'print'
        })
        SubElement(blipfill, '{%s}srcRect' % DRAWING_NS)
        SubElement(
            SubElement(blipfill, '{%s}stretch' % DRAWING_NS),
            '{%s}fillRect' % DRAWING_NS)

        sppr = SubElement(pic, '{%s}spPr' % SHEET_DRAWING_NS, {'bwMode':'auto'})
        frm = SubElement(sppr, '{%s}xfrm' % DRAWING_NS)
        # no transformation
        SubElement(frm, '{%s}off' % DRAWING_NS, {'x':'0', 'y':'0'})
        SubElement(frm, '{%s}ext' % DRAWING_NS, {'cx':'0', 'cy':'0'})

        SubElement(
            SubElement(sppr, '{%s}prstGeom' % DRAWING_NS, {'prst':'rect'})
            , '{%s}avLst' % DRAWING_NS)

        SubElement(sppr, '{%s}noFill' % DRAWING_NS)

        ln = SubElement(sppr, '{%s}ln' % DRAWING_NS, {'w':'1'})
        SubElement(ln, '{%s}noFill' % DRAWING_NS)
        SubElement(ln, '{%s}miter' % DRAWING_NS, {'lim':'800000'})
        SubElement(ln, '{%s}headEnd' % DRAWING_NS)
        SubElement(ln, '{%s}tailEnd' % DRAWING_NS, {'type':'none', 'w':'med', 'len':'med'})
        SubElement(sppr, '{%s}effectLst' % DRAWING_NS)

        SubElement(anchor, '{%s}clientData' % SHEET_DRAWING_NS)

    def write_rels(self, chart_id, image_id):

        root = Element("{%s}Relationships" % PKG_REL_NS)
        for i, chart in enumerate(self._sheet._charts):
            attrs = {'Id' : 'rId%s' % (i + 1),
                'Type' : '%s/chart' % REL_NS,
                'Target' : '../charts/chart%s.xml' % (chart_id + i) }
            SubElement(root, '{%s}Relationship' % PKG_REL_NS, attrs)
        for j, img in enumerate(self._sheet._images):
            attrs = {'Id' : 'rId%s' % (i + j + 1),
                'Type' : '{%s/image' % REL_NS,
                'Target' : '../media/image%s.png' % (image_id + j) }
            SubElement(root, '{%s}Relationship' % PKG_REL_NS, attrs)
        return get_document_content(root)


class ShapeWriter(object):
    """ one file per shape """

    schema = "http://schemas.openxmlformats.org/drawingml/2006/main"

    def __init__(self, shapes):

        self._shapes = shapes

    def write(self, shape_id):

        root = Element('c:userShapes', {'xmlns:c' : 'http://schemas.openxmlformats.org/drawingml/2006/chart'})

        for shape in self._shapes:
            anchor = SubElement(root, 'cdr:relSizeAnchor',
                {'xmlns:cdr' : "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"})

            xstart, ystart, xend, yend = shape.coordinates

            _from = SubElement(anchor, 'cdr:from')
            SubElement(_from, 'cdr:x').text = str(xstart)
            SubElement(_from, 'cdr:y').text = str(ystart)

            _to = SubElement(anchor, 'cdr:to')
            SubElement(_to, 'cdr:x').text = str(xend)
            SubElement(_to, 'cdr:y').text = str(yend)

            sp = SubElement(anchor, 'cdr:sp', {'macro':'', 'textlink':''})
            nvspr = SubElement(sp, 'cdr:nvSpPr')
            SubElement(nvspr, 'cdr:cNvPr', {'id':str(shape_id), 'name':'shape %s' % shape_id})
            SubElement(nvspr, 'cdr:cNvSpPr')

            sppr = SubElement(sp, 'cdr:spPr')
            frm = SubElement(sppr, 'a:xfrm', {'xmlns:a':self.schema})
            # no transformation
            SubElement(frm, 'a:off', {'x':'0', 'y':'0'})
            SubElement(frm, 'a:ext', {'cx':'0', 'cy':'0'})

            prstgeom = SubElement(sppr, 'a:prstGeom', {'xmlns:a':self.schema, 'prst':str(shape.style)})
            SubElement(prstgeom, 'a:avLst')

            fill = SubElement(sppr, 'a:solidFill', {'xmlns:a':self.schema})
            SubElement(fill, 'a:srgbClr', {'val':shape.color})

            border = SubElement(sppr, 'a:ln', {'xmlns:a':self.schema, 'w':str(shape._border_width)})
            sf = SubElement(border, 'a:solidFill')
            SubElement(sf, 'a:srgbClr', {'val':shape.border_color})

            self._write_style(sp)
            self._write_text(sp, shape)

            shape_id += 1

        return get_document_content(root)

    def _write_text(self, node, shape):
        """ write text in the shape """

        tx_body = SubElement(node, 'cdr:txBody')
        SubElement(tx_body, 'a:bodyPr', {'xmlns:a':self.schema, 'vertOverflow':'clip'})
        SubElement(tx_body, 'a:lstStyle',
            {'xmlns:a':self.schema})
        p = SubElement(tx_body, 'a:p', {'xmlns:a':self.schema})
        if shape.text:
            r = SubElement(p, 'a:r')
            rpr = SubElement(r, 'a:rPr', {'lang':'en-US'})
            fill = SubElement(rpr, 'a:solidFill')
            SubElement(fill, 'a:srgbClr', {'val':shape.text_color})

            SubElement(r, 'a:t').text = shape.text
        else:
            SubElement(p, 'a:endParaRPr', {'lang':'en-US'})

    def _write_style(self, node):
        """ write style theme """

        style = SubElement(node, 'cdr:style')

        ln_ref = SubElement(style, 'a:lnRef', {'xmlns:a':self.schema, 'idx':'2'})
        scheme_clr = SubElement(ln_ref, 'a:schemeClr', {'val':'accent1'})
        SubElement(scheme_clr, 'a:shade', {'val':'50000'})

        fill_ref = SubElement(style, 'a:fillRef', {'xmlns:a':self.schema, 'idx':'1'})
        SubElement(fill_ref, 'a:schemeClr', {'val':'accent1'})

        effect_ref = SubElement(style, 'a:effectRef', {'xmlns:a':self.schema, 'idx':'0'})
        SubElement(effect_ref, 'a:schemeClr', {'val':'accent1'})

        font_ref = SubElement(style, 'a:fontRef', {'xmlns:a':self.schema, 'idx':'minor'})
        SubElement(font_ref, 'a:schemeClr', {'val':'lt1'})

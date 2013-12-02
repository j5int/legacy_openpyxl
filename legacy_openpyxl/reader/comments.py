# file openpyxl/reader/comments.py

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

from os import path

from legacy_openpyxl.comments import Comment
from legacy_openpyxl.shared.ooxml import PACKAGE_WORKSHEET_RELS, PACKAGE_WORKSHEETS
from legacy_openpyxl.shared.xmltools import fromstring

def _get_author_list(root):
    author_subtree = root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}authors')
    return [author.text for author in author_subtree]

def read_comments(xml_source):
    """Given the XML of a comments file, generates a list of the comments"""
    root = fromstring(xml_source)
    authors = _get_author_list(root)
    comment_nodes = root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}comment')
    comments = []
    for node in comment_nodes:
        author = authors[int(node.attrib['authorId'])]
        cell = node.attrib['ref']
        text_node = node.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}text')
        text = ''
        substrs = []
        for run in text_node.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}r'):
            runtext = ''.join([t.text for t in run.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')])
            substrs.append(runtext)
        comment_text = ''.join(substrs)
        comments.append(Comment(cell, comment_text, author))
    return comments

def get_worksheet_comment_dict(workbook, archive, valid_files):
    """Given a workbook, return a mapping of worksheet names to comment files"""
    mapping = {}
    for i, sheet_name in enumerate(workbook.worksheets):
        sheet_codename = 'sheet%d.xml' % (i + 1)
        rels_file = PACKAGE_WORKSHEET_RELS + '/' + sheet_codename + '.rels'
        print rels_file
        if rels_file not in valid_files:
            continue
        rels_source = archive.read(rels_file)
        root = fromstring(rels_source)
        for i in root:
            if i.attrib['Type'] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments":
                comments_file = path.normpath(PACKAGE_WORKSHEETS + '/' + i.attrib['Target'])
                mapping[sheet_codename] = comments_file

    return mapping




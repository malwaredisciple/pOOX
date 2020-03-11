"""
This tool was created by @malwaredisciple to automate the triaging
of malware samples of the OOXML file format.
"""

from zipfile import ZipFile
from tree import DisplayablePath
from pathlib import Path
import os
import re
import xml.dom.minidom
import hashlib


class OOXMLparser:
    """
    docstring
    """
    TYPE_OLE_OBJ = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject'
    TYPE_FRAME = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/frame'
    TYPE_TEMPLATE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate'
    TYPE_VBA_PROJ = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'
    TYPE_EXTERNAL = 'External'

    def __init__(self, file_path):
        self.file_path = file_path
        self.file_name = re.split(r'(/|\\)', file_path)[-1]
        self.new_dir = '{}_pOOX'.format(self.file_name)
        self.file_data = open(file_path, 'rb').read()
        self.md5 = hashlib.md5(self.file_data).hexdigest()
        self.sha1 = hashlib.sha1(self.file_data).hexdigest()
        self.sha256 = hashlib.sha256(self.file_data).hexdigest()
        self._is_xls = False
        self._is_doc = False
        self._is_ppt = False
        self.main_xml = False
        self.doc_dir = None
        self._has_embeddings = False
        self._has_remote_template = False
        self._has_remote_frame = False
        self._has_macros = False
        self.vba_bin = None
        self._has_ole = False
        self._has_dde = False
        self._remote_template = None
        self.remote_frame = None
        self.ole_object = None
        self.embeddings = None
        self.docs_rels = set()

    def unzip(self):
        with ZipFile(self.file_path) as zip:
            zip.extractall(self.new_dir)

    def set_type(self):
        dirs = os.listdir(self.new_dir)
        if 'xl' in dirs:
            self._is_xls = True
        elif 'word' in dirs:
            self._is_doc = True
        elif 'ppt' in dirs:
            self._is_ppt = True

    def set_doc_dir(self):
        if self._is_ppt:
            self.doc_dir = '{}/ppt'.format(self.new_dir)
        elif self._is_doc:
            self.doc_dir = '{}/word'.format(self.new_dir)
        elif self._is_xls:
            self.doc_dir = '{}/xl'.format(self.new_dir)

    def get_main_xml_data(self):
        if self._is_xls:
            self.main_xml = self.get_data('{}/workbook.xml'.format(self.doc_dir))
        elif self._is_doc:
            self.main_xml = self.get_data('{}/document.xml'.format(self.doc_dir))
        elif self._is_ppt:
            self.main_xml = self.get_data('{}/presentation.xml'.format(self.doc_dir))

    def set_embeddings(self):
        if 'embeddings' in os.listdir(self.new_dir) or 'embeddings' in os.listdir(self.doc_dir):
            self._has_embeddings = True
            self.embeddings = os.listdir('{}/embeddings'.format(self.doc_dir))

    @staticmethod
    def get_data(path):
        return open(path, 'r').read()

    def print_report(self):
        self.print_art()
        self.print_meta()
        self.print_tree()
        self.print_analysis()
        print('')

    def print_analysis(self):
        print('\nAnalysis:')
        if self._has_macros:
            print('[+] contains VBA macros -> {}'.format(self.vba_bin))
        if self._has_remote_template:
            print('[+] remote template injection -> {}'.format(self._remote_template))
        if self._has_remote_frame:
            print('[+] remote frame injection -> {}'.format(self.remote_frame))
        if self._has_embeddings:
            print('[+] contains embedded files -> {}'.format(self.embeddings))
        if self._has_ole:
            print('[+] contains oleobject -> {}'.format(self.ole_object))

    def print_tree(self):
        print('\nTree View of De-archived OOXML:\n{}'.format('-' * 31))
        paths = DisplayablePath.make_tree(Path(self.new_dir))
        for path in paths:
            print(path.displayable())

    def print_meta(self):
        print('\nMetadata: \n{}'.format('-' * 9))
        print('Sample: {}'.format(self.file_path))
        print('MD5: {}'.format(self.md5))
        print('SHA1: {}'.format(self.sha1))
        print('SHA256: {}'.format(self.sha256))

    @staticmethod
    def print_art():
        print('{}{}{}'.format('+', '-' * 58, '+'))
        print('|{}pOOX - Parse OOXML Samples{}|'.format(' ' * 17, ' ' * 15))
        print('+{}+'.format('-' * 58))
        print('|{}author: @malwaredisciple{}|'.format(' ' * 18, ' ' * 16))
        print('+{}+'.format('-' * 58))
        print('|{}https://github.com/malwaredisciple/pOOX{}|'.format(' ' * 11, ' ' * 8))
        print('+{}+'.format('-' * 58))
        print('\n+{}+\n|pOOX Report|\n+{}+'.format('-' * 11, '-' * 11))

    def set_doc_rels(self):
        if self._is_xls and 'sheet1.xml.rels' in os.listdir('{}/worksheets/_rels'.format(self.doc_dir)):
            self.docs_rels.add('{}/worksheets/_rels/sheet1.xml.rels'.format(self.doc_dir))
        if self._is_xls and 'workbook.xml.rels' in os.listdir('{}/_rels'.format(self.doc_dir)):
            self.docs_rels.add('{}/_rels/workbook.xml.rels'.format(self.doc_dir))
        if self._is_doc and 'settings.xml.rels' in os.listdir('{}/_rels'.format(self.doc_dir)):
            self.docs_rels.add('{}/_rels/settings.xml.rels'.format(self.doc_dir))
        if self._is_doc and 'document.xml.rels' in os.listdir('{}/_rels'.format(self.doc_dir)):
            self.docs_rels.add('{}/_rels/document.xml.rels'.format(self.doc_dir))
        if self._is_doc and 'webSettings.xml.rels' in os.listdir('{}/_rels'.format(self.doc_dir)):
            self.docs_rels.add('{}/_rels/webSettings.xml.rels'.format(self.doc_dir))

    def parse_rels(self, doc_rels):
        for rels in doc_rels:
            for rel in xml.dom.minidom.parse(rels).getElementsByTagName('Relationship'):
                if rel.getAttribute('Type') == self.TYPE_OLE_OBJ:
                    self._has_ole = True
                    self.ole_object = rel.getAttribute('Target')
                elif rel.getAttribute('Type') == self.TYPE_VBA_PROJ:
                    self._has_macros = True
                    self.vba_bin = rel.getAttribute('Target')
                elif (rel.getAttribute('Type') == self.TYPE_FRAME
                      and rel.getAttribute('TargetMode') == self.TYPE_EXTERNAL):
                    self._has_remote_frame = True
                    self.remote_frame = rel.getAttribute('Target')
                elif (rel.getAttribute('Type') == self.TYPE_TEMPLATE
                      and rel.getAttribute('TargetMode') == self.TYPE_EXTERNAL):
                    self._has_remote_frame = True
                    self.remote_frame = rel.getAttribute('Target')

    def start(self):
        try:
            os.mkdir(self.new_dir)
        except:
            pass
        self.unzip()
        self.set_type()
        self.set_doc_dir()
        self.set_embeddings()
        self.set_doc_rels()
        self.parse_rels(self.docs_rels)
        self.print_report()
        return 1


if __name__ == '__main__':
    parser = OOXMLparser('/Users/slayer/samples/PAYMENT_COPY_MT103.xlsx')
    parser.start()


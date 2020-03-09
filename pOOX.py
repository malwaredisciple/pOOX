"""
This tool was created by @malwaredisciple to automate the triaging
of malware samples of the OOXML file format.
"""

from zipfile import ZipFile
from pathlib import Path
import os
import re
import xml.dom.minidom


class OOXMLparser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.file_name = re.split(r'(/|\\)', file_path)[-1]
        self.new_dir = '{}_unzipped'.format(self.file_name)
        self._is_xls = False
        self._is_doc = False
        self._is_ppt = False
        self.main_xml = False
        self.doc_dir = None
        self._has_embeddings = False
        self._has_remote_template = False
        self._remote_template = None

    def print_tree(self):
        paths = DisplayablePath.make_tree(Path(self.new_dir))
        for path in paths:
            print(path.displayable())

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
        if 'embeddings' in os.listdir(self.new_dir):
            self._has_embeddings = True

    @staticmethod
    def get_data(path):
        return open(path, 'r').read()

    def check_remote_template(self):
        rels_xml = xml.dom.minidom.parse('{}/_rels/settings.xml.rels'.format(self.doc_dir))
        rels = rels_xml.getElementsByTagName('Relationship')
        for rel in rels:
            if (rel.getAttribute('Type') ==
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate'
                    and rel.getAttribute('TargetMode') == 'External'):
                self._has_remote_template = True
                self._remote_template = rel.getAttribute('Target')

    def start(self):
        try:
            os.mkdir(self.new_dir)
        except:
            print('[-] Directory -> {}/ already exists!'.format(self.new_dir))
            return
        self.unzip()
        self.print_tree()
        self.set_type()
        self.set_doc_dir()
        self.get_main_xml_data()
        self.set_embeddings()
        self.check_remote_template()
        return 1


"""
copypasta - https://stackoverflow.com/questions/9727673/list-directory-tree-structure-in-python
"""
class DisplayablePath(object):
    display_filename_prefix_middle = '├──'
    display_filename_prefix_last = '└──'
    display_parent_prefix_middle = '    '
    display_parent_prefix_last = '│   '

    def __init__(self, path, parent_path, is_last):
        self.path = Path(str(path))
        self.parent = parent_path
        self.is_last = is_last
        if self.parent:
            self.depth = self.parent.depth + 1
        else:
            self.depth = 0

    @property
    def displayname(self):
        if self.path.is_dir():
            return self.path.name + '/'
        return self.path.name

    @classmethod
    def make_tree(cls, root, parent=None, is_last=False, criteria=None):
        root = Path(str(root))
        criteria = criteria or cls._default_criteria

        displayable_root = cls(root, parent, is_last)
        yield displayable_root

        children = sorted(list(path
                               for path in root.iterdir()
                               if criteria(path)),
                          key=lambda s: str(s).lower())
        count = 1
        for path in children:
            is_last = count == len(children)
            if path.is_dir():
                yield from cls.make_tree(path,
                                         parent=displayable_root,
                                         is_last=is_last,
                                         criteria=criteria)
            else:
                yield cls(path, displayable_root, is_last)
            count += 1

    @classmethod
    def _default_criteria(cls, path):
        return True

    @property
    def displayname(self):
        if self.path.is_dir():
            return self.path.name + '/'
        return self.path.name

    def displayable(self):
        if self.parent is None:
            return self.displayname

        _filename_prefix = (self.display_filename_prefix_last
                            if self.is_last
                            else self.display_filename_prefix_middle)

        parts = ['{!s} {!s}'.format(_filename_prefix,
                                    self.displayname)]

        parent = self.parent
        while parent and parent.parent is not None:
            parts.append(self.display_parent_prefix_middle
                         if parent.is_last
                         else self.display_parent_prefix_last)
            parent = parent.parent

        return ''.join(reversed(parts))


if __name__ == '__main__':
    parser = OOXMLparser('/Users/slayer/samples/remote.docx')
    parser.start()


# -*- coding: utf-8 -*-
"""
    sphinxcontrib-docxbuilder
    ~~~~~~~~~~~~~~~~~~~~~~~~~~~

    OpenXML Document Sphinx builder.

    :copyright:
        Copyright 2010 by shimizukawa at gmail dot com (Sphinx-users.jp).
    :license: BSD, see LICENSE for details.
"""

import codecs
from os import path

from docutils import nodes
from docutils.io import StringOutput

from sphinx.builders import Builder
from sphinx.util.osutil import ensuredir, os_path
from sphinx.util.nodes import inline_all_toctrees
from sphinx.util.console import bold, darkgreen, brown
from writer import DocxWriter


class DocxBuilder(Builder):
    name = 'docx'
    format = 'docx'
    out_suffix = '.docx'

    def init(self):
        pass

    def get_outdated_docs(self):
        return 'pass'

    def get_target_uri(self, docname, typ=None):
        return ''

    def fix_refuris(self, tree):
        # fix refuris with double anchor
        fname = self.config.master_doc + self.out_suffix
        for refnode in tree.traverse(nodes.reference):
            if 'refuri' not in refnode:
                continue
            refuri = refnode['refuri']
            hashindex = refuri.find('#')
            if hashindex < 0:
                continue
            hashindex = refuri.find('#', hashindex + 1)
            if hashindex >= 0:
                refnode['refuri'] = fname + refuri[hashindex:]

    def prepare_writing(self, docnames):
        self.writer = DocxWriter(self)

    def assemble_doctree(self):
        master = self.config.master_doc
        tree = self.env.get_doctree(master)
        tree = inline_all_toctrees(self, set(), master, tree, darkgreen)
        tree['docname'] = master
        self.env.resolve_references(tree, master, self)
        self.fix_refuris(tree)
        return tree

    def write(self, *ignored):
        docnames = self.env.all_docs

        self.info(bold('preparing documents... '), nonl=True)
        self.prepare_writing(docnames)
        self.info('done')

        self.info(bold('assembling single document... '), nonl=True)
        doctree = self.assemble_doctree()
        self.info()
        self.info(bold('writing... '), nonl=True)
        docname = "%s-%s" % (self.config.project, self.config.version)
        self.write_doc(docname, doctree)
        self.info('done')

    def write_doc(self, docname, doctree):
        destination = StringOutput(encoding='utf-8')
        self.writer.write(doctree, destination)
        outfilename = path.join(
                self.outdir, os_path(docname) + self.out_suffix)
        ensuredir(path.dirname(outfilename))
        try:
            self.writer.save(outfilename)
        except (IOError, OSError), err:
            self.warn("error writing file %s: %s" % (outfilename, err))

    def finish(self):
        #self.warn("call finish")
        pass

# -*- coding: utf-8 -*-
"""
    sphinxcontrib-docxbuilder
    ~~~~~~~~~~~~~~~~~~~~~~~~~~~

    OpenXML Document Sphinx builder.

    :copyright:
        Copyright 2010 by shimizukawa at gmail dot com (Sphinx-users.jp).
    :license: BSD, see LICENSE for details.
"""

from os import path

from docutils.io import StringOutput

from sphinx.builders import Builder
from sphinx.util.osutil import ensuredir, os_path
from sphinx.util.nodes import inline_all_toctrees
from sphinx.util.console import bold, darkgreen
from docxbuilder.writer import DocxWriter


class DocxBuilder(Builder):
    name = 'docx'
    format = 'docx'
    out_suffix = '.docx'

    def init(self):
        pass

    def get_outdated_docs(self):
        return 'pass'

    def get_target_uri(self, docname, typ=None):
        return docname

    def prepare_writing(self, docnames):
        self.writer = DocxWriter(self)

    def assemble_doctree(self):
        master = self.config.master_doc
        tree = self.env.get_doctree(master)
        tree = inline_all_toctrees(self, set(), master, tree, darkgreen, [])
        tree['docname'] = master
        self.env.resolve_references(tree, master, self)
        return tree

    def make_numfig_map(self):
        numfig_map = {}
        for docname, item in self.env.toc_fignumbers.items():
            for figtype, info in item.items():
                prefix = self.config.numfig_format.get(figtype)
                if prefix is None:
                    continue
                _, num_map = numfig_map.setdefault(figtype, (prefix, {}))
                for id, num in info.items():
                    key = '%s/%s' % (docname, id)
                    num_map[key] = num
        return numfig_map

    def make_numsec_map(self):
        numsec_map = {}
        for docname, info in self.env.toc_secnumbers.items():
            for id, num in info.items():
                key = '%s/%s' % (docname, id)
                numsec_map[key] = num
        return numsec_map

    def write(self, *ignored):
        docnames = self.env.all_docs

        self.info(bold('preparing documents... '), nonl=True)
        self.prepare_writing(docnames)
        self.info('done')

        self.info(bold('assembling single document... '), nonl=True)
        doctree = self.assemble_doctree()
        self.writer.set_numsec_map(self.make_numsec_map())
        self.writer.set_numfig_map(self.make_numfig_map())
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
        except (IOError, OSError) as err:
            self.warn("error writing file %s: %s" % (outfilename, err))

    def finish(self):
        #self.warn("call finish")
        pass

# -*- coding: utf-8 -*-
"""
    sphinxcontrib-docxbuilder
    ~~~~~~~~~~~~~~~~~~~~~~~~~~~

    OpenXML Document Sphinx builder.

    :copyright:
        Copyright 2010 by shimizukawa at gmail dot com (Sphinx-users.jp).
    :license: BSD, see LICENSE for details.
"""

import os

from docutils import nodes
from docutils.io import BinaryFileOutput
from sphinx import addnodes
from sphinx.builders import Builder
from sphinx.util import logging
from sphinx.util.docutils import new_document
from sphinx.util.osutil import ensuredir

from docxbuilder.writer import DocxWriter, DocxTranslator


class DocxBuilder(Builder):
    # pylint: disable=attribute-defined-outside-init
    name = 'docx'
    format = 'docx'
    out_suffix = '.docx'
    default_translator_class = DocxTranslator

    def init(self):
        self.imagedir = '_images'
        self._logger = logging.getLogger('docxbuilder')
        self._docx_documents = []

    def get_outdated_docs(self):
        return 'pass'

    def get_target_uri(self, docname, typ=None):
        return docname

    def prepare_writing(self, docnames):
        for entry in self.config.docx_documents:
            if entry[0] not in self.env.all_docs:
                self._logger.warning(
                    'unknown document %s is found '
                    'in docx_documents' % entry[0])
                continue
            if not entry[1]:
                self._logger.warning(
                    'invalid filename %s s found for %s '
                    'in docx_documents' % (entry[1]. entry[0]))
                continue
            self._docx_documents.append(entry)
        if not self._docx_documents:
            self._logger.warning('no valid entry is found in docx_documents')
        self.writer = DocxWriter(self)

    def assemble_doctree(self, master, toctree_only):
        tree = self.env.get_doctree(master)
        if toctree_only:
            doc = new_document('docxbuilder/builder.py')
            for toctree in tree.traverse(addnodes.toctree):
                # ids is not assigned to toctree, but to the parent
                toctree.get('ids').extend(toctree.parent.get('ids'))
                doc.append(toctree)
            tree = doc
        tree = insert_all_toctrees(tree, master, self.env, [])
        tree['docname'] = master
        self._logger.info('')
        # TODO: Support cross references
        return tree

    def make_numfig_map(self):
        numfig_map = {}
        for docname, item in self.env.toc_fignumbers.items():
            for figtype, info in item.items():
                prefix = self.config.numfig_format.get(figtype)
                if prefix is None:
                    continue
                _, num_map = numfig_map.setdefault(figtype, (prefix, {}))
                for node_id, num in info.items():
                    key = '%s/%s' % (docname, node_id)
                    num_map[key] = num
        return numfig_map

    def make_numsec_map(self):
        numsec_map = {}
        for docname, info in self.env.toc_secnumbers.items():
            for node_id, num in info.items():
                key = '%s/%s' % (docname, node_id)
                numsec_map[key] = num
        return numsec_map

    def write(self, *_ignored): # pylint: disable=arguments-differ
        docnames = self.env.all_docs

        self._logger.info('preparing documents... ', nonl=True)
        self.prepare_writing(docnames)
        self._logger.info('done')

        for entry in self._docx_documents:
            start_doc, docname, props = entry[:3]
            toctree_only = entry[3] if len(entry) > 3 else False

            self._logger.info('processing %s... ' % docname, nonl=True)
            doctree = self.assemble_doctree(start_doc, toctree_only)
            self.doc_properties = props
            self._logger.info('writing... ', nonl=True)
            self.write_doc(docname, doctree)
            self._logger.info('done')

    def write_doc(self, docname, doctree):
        outfilename = os.path.join(self.outdir, docname)
        ensuredir(os.path.dirname(outfilename))
        destination = BinaryFileOutput(destination_path=outfilename)
        self.writer.write(doctree, destination)

    def finish(self):
        pass

def insert_all_toctrees(tree, docname, env, traversed):
    tree = tree.deepcopy()
    env.apply_post_transforms(tree, docname)
    for toctreenode in tree.traverse(addnodes.toctree):
        nodeid = 'docx_expanded_toctree%d' % id(toctreenode)
        newnodes = nodes.container(ids=[nodeid])
        toctreenode['docx_expanded_toctree_refid'] = nodeid
        includefiles = toctreenode['includefiles']
        for includefile in includefiles:
            if includefile in traversed:
                continue
            try:
                traversed.append(includefile)
                subtree = insert_all_toctrees(
                    env.get_doctree(includefile), includefile, env, traversed)
            except Exception: # pylint: disable=broad-except
                continue
            start_of_file = addnodes.start_of_file(docname=includefile)
            start_of_file.children = subtree.children
            newnodes.append(start_of_file)
        parent = toctreenode.parent
        index = parent.index(toctreenode)
        parent.insert(index + 1, newnodes)
    return tree

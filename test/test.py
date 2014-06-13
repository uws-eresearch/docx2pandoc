#!/usr/bin/env python2

from __future__ import print_function
import subprocess as SP
import unittest

DOCX2PANDOC = "../docx2pandoc"

def output_with_options(infile, expected_output_file, options=None):
    if options is None: options = []
    d2p_proc = SP.Popen([DOCX2PANDOC] + options + [infile],
                        stdin=SP.PIPE, stdout=SP.PIPE, stderr=SP.PIPE)
    pd_proc  = SP.Popen(["pandoc", "-f", "json", "-t", "native"],
                        stdin=SP.PIPE, stdout=SP.PIPE, stderr=SP.PIPE)
    d2p_out, d2p_err = d2p_proc.communicate()
    pd_out, pd_err  = pd_proc.communicate(d2p_out)
    with open(expected_output_file) as fp:
        return (pd_out.strip(), fp.read().strip())

class NativeOutput(unittest.TestCase):

    def test_inline_formatting(self):
        """Test inline formatting"""
        self.assertEqual(*output_with_options(
            "inline_formatting.docx", 
            "inline_formatting.native"
        ))

    def test_headers(self):
        """Test headers"""
        self.assertEqual(*output_with_options(
            "headers.docx", 
            "headers.native"
        ))

    def test_lists(self):
        """Test lists"""
        self.assertEqual(*output_with_options(
            "lists.docx", 
            "lists.native"
        ))

    def test_image(self):
        """Test image (without embedding)"""
        self.assertEqual(*output_with_options(
            "image.docx", 
            "image_no_embed.native"
        ))

    def test_notes(self):
        """Test footnotes and endnotes"""
        self.assertEqual(*output_with_options(
            "notes.docx", 
            "notes.native"
        ))

    def test_block_quotes(self):
        """Test block quotes (with automatic indent parsing)"""
        self.assertEqual(*output_with_options(
            "block_quotes.docx", 
            "block_quotes_parse_indent.native"
        ))

    def test_links(self):
        """Test links"""
        self.assertEqual(*output_with_options(
            "links.docx", 
            "links.native"
        ))

    def test_tables(self):
        """Test tables"""
        self.assertEqual(*output_with_options(
            "tables.docx", 
            "tables.native"
        ))

    

if __name__ == '__main__':
    unittest.main(verbosity=2)


    

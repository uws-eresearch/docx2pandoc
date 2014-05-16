from lxml import etree
import zipfile
import re

class DocxError(Exception):
    pass

class Docx(object):
    def __init__(self, body, notes, rels):
        self.body = body
        self.notes = notes
        self.rels = rels

    @classmethod
    def read_file(klass, filepath):
        zf = zipfile.ZipFile(filepath)
        
        doc_element = etree.parse(zf.open("word/document.xml")).getroot()
        body_element = doc_element.find('w:body', namespaces=doc_element.nsmap)
        
        try:
            footnotes_element = etree.parse(zf.open("word/footnotes.xml")).getroot()
        except KeyError:
            footnotes_element = None

        try:
            endnotes_element = etree.parse(zf.open("word/endnotes.xml")).getroot()
        except KeyError:
            endnotes_element = None

        rels = [zf.open(info.filename) for info in zf.filelist 
                if 
                re.match("^word/_rels/.*\.rel", info.filename)]
        rels_elements = [etree.parse(rel).getroot() for rel in rels]

        return klass(Body(body_element), 
                     Notes(footnotes_element, endnotes_element),
                     Rels(rels_elements))

        
class Notes(object):
    def __init__(self, footnotes_element, endnotes_element):
        self._footnotes = footnotes_element
        self._endnotes = endnotes_element

class Rels(object):

    def __init__(self, rels):
        self._rels = rels
    

class Body(object):
    def __init__(self, body_element):
        self._body = body_element

    def get_paragraphs(self):
        return [Paragraph(par_element) for par_element in 
                self._body.findall('w:p', namespaces=self._body.nsmap)]


class Paragraph(object):
    def __init__(self, p_element):
        self._p = p_element
        self._pPr = self._p.find('.//w:pPr', namespaces=self._p.nsmap)

    def get_run_containers(self):
        return [RunContainer.new(rc_elem) for rc_elem in 
                self._p.xpath('./w:r|./w:hyperlink', namespaces=self._p.nsmap)]

    @property
    def style(self):
        if self._pPr is None: return None
        result = self._pPr.find('w:pStyle', namespaces=self._p.nsmap)
        if result is None:
            return None
        else:
            return result.get("{%s}val" % result.nsmap["w"])


class RunContainer(object):

    @classmethod
    def new(self, element):
        if element.tag == ("{%s}r" % element.nsmap['w']):
            return Run(element)
        elif element.tag == ("{%s}hyperlink" % element.nsmap['w']):
            return HyperLink(element)
        else:
            raise DocxError("%r not supported run container type" % element)

    def get_runs(self):
        return [Run(elem) for elem 
                in 
                self._container.getiterator("{%s}r" % self._container.nsmap["w"])]

class HyperLink(RunContainer):
    def __init__(self, hyperlink_element):
        self._hyperlink = self._container = hyperlink_element
        
        

class Run(RunContainer):
    def __init__(self, run_element):
        self._run = self._container = run_element
        self._rPr = self._run.find('.//w:rPr', namespaces=self._run.nsmap)

    @property
    def bold(self):
        if self._rPr is None: return False
        return (self._rPr.find('w:b', namespaces=self._run.nsmap) is not None)

    @property
    def italic(self):
        if self._rPr is None: return False
        return (self._rPr.find('w:i', namespaces=self._run.nsmap) is not None)

    @property
    def smallCaps(self):
        if self._rPr is None: return False
        result = self._rPr.find('w:smallCaps', namespaces=self._run.nsmap)
        if result is None:
            return False
        elif (result.get("{%s}val" % result.nsmap["w"]) is None 
              or result.get("{%s}val" % result.nsmap["w"]) == "true"):
            return True
        else:
            return False

    @property
    def strike(self):
        if self._rPr is None: return False
        result = self._rPr.find('w:strike', namespaces=self._run.nsmap)
        if result is None:
            return False
        elif (result.get("{%s}val" % result.nsmap["w"]) is None 
              or result.get("{%s}val" % result.nsmap["w"]) == "true"):
            return True
        else:
            return False

    @property
    def style(self):
        if self._rPr is None: return None
        result = self._rPr.find('w:rStyle', namespaces=self._run.nsmap)
        if result is None:
            return None
        else:
            return result.get("{%s}val" % result.nsmap["w"])
        

            
        
        
        
        






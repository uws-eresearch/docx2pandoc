from lxml import etree
import zipfile
import re

class DocxError(Exception):
    pass

class Docx(object):

    def __init__(self):
        self.body = None
        self.notes = None
        self.relationships = None
        self.namespaces = None

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

        obj = klass()
        obj.body = Body(body_element, obj)
        obj.notes = Notes(footnotes_element, endnotes_element, obj)
        obj.relationships = Relationships(rels_elements, obj)
        obj.namespaces = doc_element.nsmap
        return obj

class DocxPart(object):

    def __init__(self, docx):
        self.docx = docx
    

class Notes(DocxPart):
    def __init__(self, footnotes_element, endnotes_element, docx):
        super(Notes, self).__init__(docx)
        self._footnotes = footnotes_element
        self._endnotes = endnotes_element
        self._endnote_table = None
        self._footnote_table = None

    @property
    def endnote_table(self):
        if self._endnote_table is None:
            self._endnote_table = {note.id:note for note in 
                                   EndNoteCollection(self._endnotes, self)}
        return self._endnote_table

    @property
    def footnote_table(self):
        if self._footnote_table is None:
            self._footnote_table = {note.id:note for note in 
                                    FootNoteCollection(self._footnotes, self)}
        return self._footnote_table

    def get_endnote(self, note_id):
        return self.endnote_table[note_id]

    def get_footnote(self, note_id):
        return self.footnote_table[note_id]



class NoteCollection(object):

    def __init__(self, notes_element, parent):
        self._notes = notes_element
        self.parent = parent

class FootNoteCollection(NoteCollection):

    def __iter__(self):
        if self._notes is None:
            raise StopIteration
        for elem in self._notes.getchildren():
            yield FootNote(elem, self.parent)

class EndNoteCollection(NoteCollection):

    def __iter__(self):
        if self._notes is None:
            raise StopIteration
        for elem in self._notes.getchildren():
            yield EndNote(elem, self.parent)

    
class Note(object):

    def __init__(self, note_element, parent):
        self._note = note_element
        self.parent = parent

    @property
    def id(self):
        return self._note.get("{%s}id" % self._note.nsmap["w"])

    def get_paragraphs(self):
        docx = self.parent.docx
        return [Paragraph(par_element, docx.body) for par_element in 
                self._note.findall('w:p', namespaces=self._note.nsmap)]

class EndNote(Note):

    @property
    def note_type(self):
        return "endnote"

class FootNote(Note):

    @property
    def note_type(self):
        return "footnote"


class Relationships(DocxPart):

    def __init__(self, rel_lst, docx):
        self._rel_lst = rel_lst
        self._relationship_table = None
        super(Relationships, self).__init__(docx)

    @property
    def relationship_table(self):
        if self._relationship_table is None:
            self._relationship_table = {r.id:r for r 
                                        in 
                                        self}
        return self._relationship_table

    def get_id(self, Id):
        return self.relationship_table[Id]

    def __iter__(self):
        for rel_lst in self._rel_lst:
            for rel_element in rel_lst:
                yield Relationship(rel_element, self)
        

class Relationship(object):

    def __init__(self, rel_element, parent):
        self._rel = rel_element
        self._attrib = self._rel.attrib
        self.parent = parent

    @property
    def id(self):
        return self._rel.get("Id")

    @property
    def target_mode(self):
        return self._rel.get("TargetMode")

    @property
    def target(self):
        return self._rel.get("Target")

    @property
    def type(self):
        nstype = self._rel.get("Type")
        ns = self.parent.docx.namespaces["r"]
        if '/'.join(nstype.split('/')[:-1]) == ns:
            return nstype.split('/')[-1]
        else:
            return nstype


class Body(DocxPart):
    def __init__(self, body_element, docx):
        self._body = body_element
        super(Body, self).__init__(docx)

    def get_paragraphs(self):
        return [Paragraph(par_element, self) for par_element in 
                self._body.findall('w:p', namespaces=self._body.nsmap)]


class Paragraph(object):
    def __init__(self, p_element, parent):
        self._p = p_element
        self._pPr = self._p.find('.//w:pPr', namespaces=self._p.nsmap)
        self.parent = parent

    def get_run_containers(self):
        return [RunContainer.new(rc_elem, self) for rc_elem in 
                self._p.xpath('./w:r|./w:hyperlink', namespaces=self._p.nsmap)]

    @property
    def indent(self):
        try:
            ind = self._pPr.find('w:ind', namespaces=self._p.nsmap)
            indent_amt = ind.get("{%s}left" % self._p.nsmap["w"])
            if indent_amt != "0":
                return indent_amt
            else:
                return None
        except AttributeError:
            return None


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
    def new(self, element, parent):
        if element.tag == ("{%s}r" % element.nsmap['w']):
            return Run(element, parent)
        elif element.tag == ("{%s}hyperlink" % element.nsmap['w']):
            if element.get("{%s}anchor" % element.nsmap["w"]) is not None:
                return InternalHyperLink(element, parent)
            else:
                return ExternalHyperLink(element, parent)
        else:
            raise DocxError("%r not supported run container type" % element)

    def get_runs(self):
        return [Run(elem, self.parent) for elem 
                in 
                self._container.getiterator("{%s}r" % self._container.nsmap["w"])]

class HyperLink(RunContainer):
    def __init__(self, hyperlink_element, parent):
        self._hyperlink = self._container = hyperlink_element
        self.parent = parent

class ExternalHyperLink(HyperLink):

    @property
    def id(self):
        return self._hyperlink.get('{%s}id' % self._hyperlink.nsmap["r"])

    @property
    def relationship(self):
        doc = self.parent.parent.docx
        return doc.relationships.get_id(self.id)

    @property
    def target(self):
        return self.relationship.target

class InternalHyperLink(HyperLink):

    @property
    def target(self):
        return '#%s' % self._hyperlink.get('{%s}anchor' % self._hyperlink.nsmap["w"])


class Run(RunContainer):
    def __init__(self, run_element, parent):
        self._run = self._container = run_element
        self._rPr = self._run.find('./w:rPr', namespaces=self._run.nsmap)
        self.parent = parent

    def get_footnote(self):
        ref = self._run.find("w:footnoteReference", namespaces=self._run.nsmap)
        if ref is not None:
            return ref.get("{%s}id" % self._run.nsmap["w"])
        else:
            return None

    def get_endnote(self):
        ref = self._run.find("w:endnoteReference", namespaces=self._run.nsmap)
        if ref is not None:
            return ref.get("{%s}id" % self._run.nsmap["w"])
        else:
            return None

    def get_text(self):
        t = self._run.find('./w:t', namespaces=self._run.nsmap)
        if t is None:
            return ''
        else:
            return t.text
    
    @property
    def runtype(self):
        if self.get_footnote() is not None:
            return "footnote"
        elif self.get_endnote() is not None:
            return "endnote"
        else:
            return "text"
        
    @property
    def is_bold(self):
        if self._rPr is None: return False
        return (self._rPr.find('w:b', namespaces=self._run.nsmap) is not None)

    @property
    def is_italic(self):
        if self._rPr is None: return False
        return (self._rPr.find('w:i', namespaces=self._run.nsmap) is not None)

    @property
    def is_smallCaps(self):
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
    def is_strike(self):
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
    def underline(self):
        if self._rPr is None: return False
        result = self._rPr.find('w:u', namespaces=self._run.nsmap)
        if result is None:
            return None
        elif result.get("{%s}val" % result.nsmap["w"]) is None:
            return "default"
        elif result.get("{%s}val" % result.nsmap["w"]) == "none":
            return None
        else:
            return result.get("{%s}val" % result.nsmap["w"])


    @property
    def style(self):
        if self._rPr is None: return None
        result = self._rPr.find('w:rStyle', namespaces=self._run.nsmap)
        if result is None:
            return None
        else:
            return result.get("{%s}val" % result.nsmap["w"])
        

            
        
        
        
        






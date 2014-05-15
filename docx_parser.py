from lxml import etree
import zipfile

class Docx(object):
    def __init__(self, body):
        self.body = body
        self.notes = notes
        self.rels = rels

class Body(object):
    def __init__(self, body_element):
        self._body = body_element

class Paragraph(object):
    def __init__(self, p_element):
        self._p = p_element

class RunContainer(object):

    def __init__(self, container_element):
        self._container = container_element

    def get_runs(self):
        return [Run(elem) for elem 
                in 
                self._container.getiterator("{%s}r" % self._container.nsmap["w"])]

class Run(object):
    def __init__(self, run_element):
        self._run = run_element
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
        elif result.get("val") is None or result.get("val") == "true":
            return True
        else:
            return False

    @property
    def strike(self):
        if self._rPr is None: return False
        result = self._rPr.find('w:strike', namespaces=self._run.nsmap)
        if result is None:
            return False
        elif result.get("val") is None or result.get("val") == "true":
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
            return result.get("val")
        

            
        
        
        
        






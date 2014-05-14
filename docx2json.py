import docx
from lxml import etree
import zipfile
import sys
import json
import pandocfilters as PF
import re

def get_indent(p):
    try:
        pPr =p._p.find('w:pPr', namespaces=p._p.nsmap)
        ind = pPr.find('w:ind', namespaces=p._p.nsmap)
        indent_amt = ind.get("{%s}left" % p._p.nsmap["w"])
        if indent_amt != "0":
            return indent_amt
        else:
            return None
    except AttributeError:
        return None

def get_extra_info(p):
    extra_info = []
    indent_amt = get_indent(p)
    if indent_amt:
        extra_info.append(("indent", indent_amt))
    return extra_info

def r2json(r):
    split_text = re.split(r'( +)', r.text)
    inlines = [PF.Space() if re.match(r'( +)', i) 
               else PF.Str(i) 
               for i in split_text]
    attr = []
    if r.bold:
        attr.append("strong")
    if r.italic:
        attr.append("emph")
    if r.small_caps:
        attr.append("smallcaps")
    if r.underline:
        attr.append("underline")
    if r.strike:
        attr.append("strikeout")

    if len(attr) > 0:
        return [PF.Span(("", attr, []), inlines)]
    else:
        return inlines

def p2json(p):
    styles  = [s for s in p.style.split() if s != "Normal"]
    kvs     = get_extra_info(p)
    inlines = reduce(list.__add__, [r2json(r) for r in p.runs], [])
    para    = PF.Para(inlines)
    if len(styles) > 0 or len(kvs) > 0:
        return PF.Div(("", styles, kvs), [para])
    else:
        return para

def doc2json(doc):
    return [p2json(p) for p in doc.paragraphs]

def doc2meta(doc):
    return {"unMeta": {}}


if __name__ == '__main__':
    docname = sys.argv[1]
    doc = docx.Document(docname)
    print json.dumps([doc2meta(doc), doc2json(doc)])
        

        

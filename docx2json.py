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

def get_part(doc, path):
    parts = doc._package.parts
    try:
        part = [p for p in parts if p.partname==path][0]
        return part.blob
    except IndexError:
        return None

def _xml_p_to_docx_p(p_elt):
    oxml_p = docx.oxml.shared.oxml_fromstring(etree.tostring(p_elt))
    return docx.text.Paragraph(oxml_p)
    

def get_notes(doc, notetype="foot"):
    notepath = "/word/%snotes.xml" % notetype
    notes_blob = get_part(doc, notepath)
    notes_tree = etree.fromstring(notes_blob)
    note_elts = notes_tree.findall("w:%snote" % notetype, namespaces=notes_tree.nsmap)
    note_dict = {note.get("{%s}id" % notes_tree.nsmap["w"]):
                 [_xml_p_to_docx_p(p) for p in 
                  note.findall("w:p", namespaces=notes_tree.nsmap)]
                 for note in notes_tree}
    return note_dict

def run_is_note(r, notetype="foot"):
    ref = r._r.find("w:%snoteReference" % notetype, namespaces=r._r.nsmap)
    if ref is not None:
        return ref.get("{%s}id" % r._r.nsmap["w"])
    else:
        return None

def handle_run_note(run_id, footnotes, endnotes, notetype="foot"):
    if notetype == "foot":
        ref = footnotes[run_id]
    elif notetype == "end":
        ref = endnotes[run_id]
    else:
        return None

    note = PF.Note([p2json(p, footnotes, endnotes) for p in ref])
    if notetype == "foot":
        return PF.Span(("", ["footnote"], []), [note])
    elif notetype == "end":
        return PF.Span(("", ["endnote"], []), [note])
    else:
        return None
        

def r2json(r, footnotes, endnotes):
    # First we check whether the run is a footnote or an endnote.
    footnote_id = run_is_note(r, "foot")
    if footnote_id:
        note = handle_run_note(footnote_id, footnotes, endnotes, "foot")
        if note:
            return [note]

    endnote_id = run_is_note(r, "end")
    if endnote_id:
        note = handle_run_note(endnote_id, footnotes, endnotes, "end")
        if note:
            return [note]

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

def p2json(p, footnotes, endnotes):
    styles  = [s for s in p.style.split() if s != "Normal"]
    kvs     = get_extra_info(p)
    inlines = reduce(list.__add__, [r2json(r, footnotes, endnotes) 
                                    for r in p.runs], [])
    para    = PF.Para(inlines)
    if len(styles) > 0 or len(kvs) > 0:
        return PF.Div(("", styles, kvs), [para])
    else:
        return para

def doc2json(doc):
    footnotes = get_notes(doc, "foot")
    endnotes  = get_notes(doc, "end")
    return [p2json(p, footnotes, endnotes) for p in doc.paragraphs]

def doc2meta(doc):
    return {"unMeta": {}}


if __name__ == '__main__':
    docname = sys.argv[1]
    doc = docx.Document(docname)
    print json.dumps([doc2meta(doc), doc2json(doc)])
        

        

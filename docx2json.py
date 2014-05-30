import docx_parser
from lxml import etree
import sys
import json
import pandocfilters as PF
import re
from span_div_reducer import reduce_elem_list
from span_div_replacer import tag_correct

def run_container_to_json(rc, doc):

    if isinstance(rc, docx_parser.HyperLink):
        ils = []
        for r in rc.get_runs():
            ils.extend(run_container_to_json(r, doc))
        return [PF.Link(ils, (rc.target, ""))]
    else:                       
    # it must be a run. Note that we're keeping this in a separate
    # else block for now, in case I add other sorts of run containers
    # later.
        if rc.runtype == "footnote":
            note_id = rc.get_footnote()
            note = doc.notes.get_footnote(note_id)
            pars = note.get_paragraphs()
            block = PF.Div(("", ["footnote"], []), [para_to_json(p, doc) for p in pars])
            return [PF.Note([block])]
        elif rc.runtype == "endnote":
            note_id = rc.get_endnote()
            note = doc.notes.get_endnote(note_id)
            pars = note.get_paragraphs()
            block = PF.Div(("", ["endnote"], []), [para_to_json(p, doc) for p in pars])
            return [PF.Note([block])]
        else:
            text = rc.get_text()
            try:
                split_text = re.split(r'( +)', text)
            except TypeError:
                return []

            
            inlines = [PF.Space() if re.match(r'( +)', i) 
                       else PF.Str(i) 
                       for i in split_text]
            attr = []
            kvs = []
            if rc.is_bold:
                attr.append("strong")
            if rc.is_italic:
                attr.append("emph")
            if rc.is_smallCaps:
                attr.append("smallcaps")
            if rc.is_strike:
                attr.append("strikeout")
            # if rc.underline is not None:
            #     attr.append("underline")
            #     kvs.append(("ul_style", rc.underline))

            if len(attr) == len(kvs) == 0:
                return inlines
            else:
                return [PF.Span(("", attr, []), inlines)]

def get_extra_para_info(p, doc):
    kvs = []
    if p.indent is not None:
        kvs.append(["indent", p.indent])
    return kvs

def para_to_json(p, doc):
    if p.style is None:
        styles = []
    else:
        styles  = [s for s in p.style.split() if s != "Normal"]
    kvs     = get_extra_para_info(p, doc)

    first_pass_inlines = reduce(list.__add__, 
                                [run_container_to_json(rc, doc) 
                                 for rc in p.get_run_containers()], 
                                [])
    #inlines = reduce_elem_list(first_pass_inlines)
    inlines = first_pass_inlines
    para    = PF.Para(inlines)
    if len(styles) > 0 or len(kvs) > 0:
        out = PF.Div(("", styles, kvs), [para])
    else:
        out = para

    # Finally, we have to deal with the list item stuff, which needs
    # its own div, that can't be stacked.
    if p.is_list_item:
        level = p.level
        num   = p.get_num()
        num_id = num.id
        frmt = num.get_abstract_num().get_level(level).format
        txt  = num.get_abstract_num().get_level(level).text
        start = num.get_abstract_num().get_level(level).start
        kvs = [["level", str(level)], 
               ["num-id", str(num_id)], 
               ["format", frmt], 
               ["text", txt]]
        if start is not None:
            kvs.append(["start", start])
        return PF.Div(("", ["list-item"], kvs), [out])
    else:
        return out
                    
    

def doc2json(doc):
    out = [para_to_json(p, doc) for p in doc.body.get_paragraphs()]

    # reduced = reduce_elem_list([para_to_json(p, doc) 
    #                             for p 
    #                             in doc.body.get_paragraphs()])
    return out

def doc2meta(doc):
    return {"unMeta": {}}

def filter(doc):
    whole_doc = [doc2meta(doc), doc2json(doc)]
    # Annoying, but we need it to be in proper json form, but I need
    # to use tuples for some of the reducing operations above.
    sanitized_output = json.loads(json.dumps(whole_doc))
    return sanitized_output
    #return PF.walk(sanitized_output, tag_correct, "", {})

if __name__ == '__main__':
    docname = sys.argv[1]
    doc = docx_parser.Docx.read_file(docname)
    filtered_doc = filter(doc)
    print json.dumps(filtered_doc)
        

        

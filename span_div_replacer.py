#!/usr/bin/env python2

import pandocfilters as PF

SPAN_CLASS_REPLACEMENTS = {"emph": PF.Emph,
                           "strong": PF.Strong,
                           "smallcaps": PF.SmallCaps,
                           "strikeout": PF.Strikeout}

DIV_CLASS_REPLACEMENTS = {"Quotations": PF.BlockQuote}

def replace_header_div(level, para):
    ils = para['c']
    return PF.Header(level, ['', [], []], ils)

DIV_CLASS_REPLACEMENTS['Heading1'] = lambda blks: replace_header_div(1, blks[0])
DIV_CLASS_REPLACEMENTS['Heading2'] = lambda blks: replace_header_div(2, blks[0])
DIV_CLASS_REPLACEMENTS['Heading3'] = lambda blks: replace_header_div(3, blks[0])
DIV_CLASS_REPLACEMENTS['Heading4'] = lambda blks: replace_header_div(4, blks[0])
DIV_CLASS_REPLACEMENTS['Heading5'] = lambda blks: replace_header_div(5, blks[0])
DIV_CLASS_REPLACEMENTS['Heading6'] = lambda blks: replace_header_div(5, blks[0])



def tag_correct(key, value, format, meta):
    if key == "Span":
        ident, classes, kvs = value[0]
        ils = value[1]
        for tag, func in SPAN_CLASS_REPLACEMENTS.items():
            if tag in classes:
                new_classes = [c for c in classes if c != tag]
                if len(ident) == len(new_classes) == len(kvs) == 0:
                    out = ils
                else:
                    out = [PF.Span((ident, new_classes, kvs), ils)]
                return func(PF.walk(out, tag_correct, "", {}))
    elif key == "Div":
        ident, classes, kvs = value[0]
        blks = value[1]
        for tag, func in DIV_CLASS_REPLACEMENTS.items():
            if tag in classes:
                new_classes = [c for c in classes if c != tag]
                if len(ident) == len(new_classes) == len(kvs) == 0:
                    out = blks
                else:
                    out = [PF.Div((ident, new_classes, kvs), blks)]
                return func(PF.walk(out, tag_correct, "", {}))
        
       
                # return func(PF.walk(out, tag_correct, format, meta))

if __name__ == '__main__':
    PF.toJSONFilter(tag_correct)




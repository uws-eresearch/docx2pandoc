import pandocfilters as PF

SPAN_CLASS_REPLACEMENTS = {"emph": PF.Emph,
                           "strong": PF.Strong,
                           "smallcaps": PF.SmallCaps,
                           "strikeout": PF.Strikeout} 

def tag_correct(key, value, format, meta):
    if key == "Span":
        ident, classes, kvs = value[0]
        ils = value[1]
        for tag, func in SPAN_CLASS_REPLACEMENTS.items():
            if tag in classes:
                new_classes = [c for c in classes if c != tag]
                if len(ident) == len(classes) == len(kvs) == 0:
                    out = ils
                else:
                    out = [PF.Span((ident, new_classes, kvs), ils)]
                return func(PF.walk(out, tag_correct, format, meta))




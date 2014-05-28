from __future__ import print_function
import pandocfilters as PF

class PFContainerError(Exception):
    pass

def get_attr(cont):
    try:
        return cont["c"][0]
    except (KeyError, IndexError):
        raise PFContainerError("Not a container or wrong container")

def get_elems(cont):
    try:
        return cont["c"][1]
    except (KeyError, IndexError):
        raise PFContainerError("Not a container or wrong container")

    
def combine_attr(attr1, attr2):
    kvs1 = [tuple(l) for l in attr1[2]]
    kvs2 = [tuple(l) for l in attr2[2]]
    shared_attr = ("", 
                   list(set(attr1[1]).intersection(attr2[1])),
                   list(set(kvs1).intersection(kvs2)))

    remaining_attr1 = (attr1[0],
                       list(set(attr1[1]).difference(shared_attr[1])),
                       list(set(kvs1).difference(shared_attr[2])))

    remaining_attr2 = (attr2[0],
                       list(set(attr2[1]).difference(shared_attr[1])),
                       list(set(kvs2).difference(shared_attr[2])))

    return (shared_attr, remaining_attr1, remaining_attr2)

def attr_is_null(attr):
    return (len(attr[0]) == len(attr[1]) == len(attr[2]) == 0)
     
    
def add_container(cont1, cont2):
    try:
        if cont1["t"] == cont2["t"] == "Div":
            cont_func = PF.Div
        elif cont1["t"] == cont2["t"] == "Span":
            cont_func = PF.Span
        else:
            raise PFContainerError("Container types don't match or are wrong types")
    except (KeyError, TypeError):
        raise PFContainerError("Not a pandoc container")

    shared, rem1, rem2 = combine_attr(get_attr(cont1), get_attr(cont2))
    

    if attr_is_null(shared):
        if attr_is_null(rem1):
            out1 = get_elems(cont1)
        else:
            out1 = [cont1]

        if attr_is_null(rem2):
            out2 = get_elems(cont2)
        else:
            out2 = [cont2]

        return (out1 + out2)
    else:
        contents1 = cont_func(rem1, get_elems(cont1))
        contents2 = cont_func(rem2, get_elems(cont2))
        return [cont_func(shared, add_container(contents1, contents2))]

def add_elems(elem1, elem2):
    try:
        return add_container(elem1, elem2)
    except PFContainerError:
        return [elem1, elem2]

def reduce_elems(elem_lst, elem):
    if len(elem_lst) == 0:
        return [elem]
    else:
        return elem_lst[:-1] + add_elems(elem_lst[-1], elem)

def reduce_elem_list(elem_lst):
    return reduce(reduce_elems, elem_lst, [])

    

    



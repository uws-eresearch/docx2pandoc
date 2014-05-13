import Text.Pandoc
import Text.Pandoc.JSON
import Data.List

spanReduce :: [Inline] -> [Inline]
spanReduce [] = []
spanReduce (s1@(Span (id1, classes1, kvs1) ils1) :
            s2@(Span (id2, classes2, kvs2) ils2) :
            ils) =
  let classes'  = classes1 `intersect` classes2
      kvs'      = kvs1 `intersect` kvs2
      classes1' = classes1 \\ classes'
      kvs1'     = kvs1 \\ kvs'
      classes2' = classes2 \\ classes'
      kvs2'     = kvs2 \\ kvs'
  in
   case null classes' && null kvs' of
     True -> s1 : (spanReduce (s2 : ils))
     False -> let attr'  = ("", classes', kvs')
                  attr1' = (id1, classes1', kvs1')
                  attr2' = (id2, classes2', kvs2')
              in
               spanReduce (Span attr' [(Span attr1' ils1), (Span attr2' ils2)] :
                           ils)
spanReduce (il:ils) = il : (spanReduce ils)

cleanUpEmpties :: [Inline] -> [Inline]
cleanUpEmpties [] = []
cleanUpEmpties ((Span ("", [], []) ils) : ils') = ils ++ (cleanUpEmpties ils')
cleanUpEmpties (il : ils) = il : (cleanUpEmpties ils)

foo :: Pandoc -> Pandoc
foo = (bottomUp cleanUpEmpties) . (bottomUp spanReduce)

main = toJSONFilter foo
       
      
  
  
  

import Text.Pandoc
import Text.Pandoc.JSON
import Text.Pandoc.Shared
import Text.Pandoc.DocX.Parser
import Data.Maybe
import Data.Char (isSpace)
import Data.List

runStyleToSpanAttr :: RunStyle -> (String, [String], [(String, String)])
runStyleToSpanAttr rPr = ("",
                          mapMaybe id [
                            if isBold rPr then (Just "strong") else Nothing,
                            if isItalic rPr then (Just "emph") else Nothing,
                            if isSmallCaps rPr then (Just "smallcaps") else Nothing,
                            if isStrike rPr then (Just "strike") else Nothing,
                            rStyle rPr],
                          case underline rPr of
                            Just fmt -> [("underline", fmt)]
                            _        -> []
                         )

parStyleToDivAttr :: ParagraphStyle -> (String, [String], [(String, String)])
parStyleToDivAttr pPr = ("",
                          mapMaybe id [pStyle pPr],
                          case indent pPr of
                            Just n  -> [("indent", (show n))]
                            Nothing -> []
                         )


strToInlines :: String -> [Inline]
strToInlines "" = []
strToInlines s  =
  let (v, w) = span (not . isSpace) s
      (x, y) = span isSpace w
  in
   (Str v) : Space : (strToInlines y)

runToInline :: DocX -> Run -> Inline
runToInline _ (Run rs s) =
  Span (runStyleToSpanAttr rs) (strToInlines s)
runToInline docx@(DocX _ notes _ _) (Footnote fnId) =
  case (getFootNote fnId notes) of
    Just bodyParts ->
      Note [Div ("", ["footnote"], []) (map (bodyPartToBlock docx) bodyParts)]
    Nothing        ->
      Note [Div ("", ["footnote"], []) []]
runToInline docx@(DocX _ notes _ _) (Endnote fnId) =
  case (getFootNote fnId notes) of
    Just bodyParts ->
      Note [Div ("", ["endnote"], []) (map (bodyPartToBlock docx) bodyParts)]
    Nothing        ->
      Note [Div ("", ["endnote"], []) []]

parPartToInline :: DocX -> ParPart -> Inline
parPartToInline docx (PlainRun r) = runToInline docx r
parPartToInline docx (InternalHyperLink anchor runs) =
  Link (map (runToInline docx) runs) (anchor, "")
parPartToInline docx@(DocX _ _ _ rels) (ExternalHyperLink relid runs) =
  case lookupRelationship relid rels of
    Just target ->
      Link (map (runToInline docx) runs) (target, "")
    Nothing ->
      Link (map (runToInline docx) runs) ("", "")

bodyPartToBlock :: DocX -> BodyPart -> Block
bodyPartToBlock docx (Paragraph pPr parparts) =
  Div (parStyleToDivAttr pPr) [Para (map (parPartToInline docx) parparts)]
bodyPartToBlock docx@(DocX _ _ numbering _) (ListItem pPr numId lvl parparts) =
  let fmt = ""
      txt = ""
  in
   Div
   ("",
    ["list-item"],
    [("level", lvl), ("num-id", numId), ("format", fmt), ("text", txt)])
   [bodyPartToBlock docx (Paragraph pPr parparts)]



spanReduce :: [Inline] -> [Inline]
spanReduce [] = []
spanReduce (s1@(Span (id1, classes1, kvs1) ils1) : ils)
  | (id1, classes1, kvs1) == ("", [], []) = ils1 ++ (spanReduce ils)
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

divReduce :: [Block] -> [Block]
divReduce [] = []
divReduce (d1@(Div (id1, classes1, kvs1) blks1) : blks)
  | (id1, classes1, kvs1) == ("", [], []) = blks1 ++ (divReduce blks)
divReduce (d1@(Div (id1, classes1, kvs1) blks1) :
            d2@(Div (id2, classes2, kvs2) blks2) :
            blks) =
  let classes'  = classes1 `intersect` classes2
      kvs'      = kvs1 `intersect` kvs2
      classes1' = classes1 \\ classes'
      kvs1'     = kvs1 \\ kvs'
      classes2' = classes2 \\ classes'
      kvs2'     = kvs2 \\ kvs'
  in
   case null classes' && null kvs' of
     True -> d1 : (divReduce (d2 : blks))
     False -> let attr'  = ("", classes', kvs')
                  attr1' = (id1, classes1', kvs1')
                  attr2' = (id2, classes2', kvs2')
              in
               divReduce (Div attr' [(Div attr1' blks1), (Div attr2' blks2)] :
                           blks)
divReduce (blk:blks) = blk : (divReduce blks)



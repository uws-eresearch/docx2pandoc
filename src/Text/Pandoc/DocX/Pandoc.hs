module Text.Pandoc.DocX.Pandoc
       ( archiveToBlocks
       ) where

import Codec.Archive.Zip
import Text.Pandoc.MIME
import Text.Pandoc.DocX.Parser
import Text.Pandoc.DocX.ItemLists (blocksToBullets,blocksToDefinitions)
import Data.Maybe
import Data.Char (isSpace)
import Data.List
import Text.Pandoc
import Text.Pandoc.UTF8 (toString)
import qualified Data.ByteString as BS
import qualified Data.ByteString.Lazy as B
import Data.ByteString.Base64
import System.FilePath (combine)

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
                          pStyle pPr,
                          case indent pPr of
                            Just n  -> [("indent", (show n))]
                            Nothing -> []
                         )


strToInlines :: String -> [Inline]
strToInlines "" = []
strToInlines s  =
  let (v, w) = span (not . isSpace) s
      (_, y) = span isSpace w
  in
      case null w of
        True  -> [Str v]
        False -> (Str v) : Space : (strToInlines y)

codeSpan :: String
codeSpan = "VerbatimChar"

blockQuoteDiv :: String
blockQuoteDiv = "BlockQuote"

codeDiv :: String
codeDiv = "SourceCode"

runToInline :: DocX -> Run -> Inline
runToInline _ (Run rs s) 
  | rStyle rs == Just codeSpan = Span (runStyleToSpanAttr rs) [Str s]
  | otherwise =  Span (runStyleToSpanAttr rs) (strToInlines s)
runToInline docx@(DocX _ notes _ _ _ ) (Footnote fnId) =
  case (getFootNote fnId notes) of
    Just bodyParts ->
      Note [Div ("", ["footnote"], []) (map (bodyPartToBlock docx) bodyParts)]
    Nothing        ->
      Note [Div ("", ["footnote"], []) []]
runToInline docx@(DocX _ notes _ _ _) (Endnote fnId) =
  case (getEndNote fnId notes) of
    Just bodyParts ->
      Note [Div ("", ["endnote"], []) (map (bodyPartToBlock docx) bodyParts)]
    Nothing        ->
      Note [Div ("", ["endnote"], []) []]

parPartToInline :: DocX -> ParPart -> Inline
parPartToInline docx (PlainRun r) = runToInline docx r
parPartToInline (DocX _ _ _ rels _) (Drawing relid) =
  case lookupRelationship relid rels of
    Just target -> Image [] (combine "word" target, "")
    Nothing     -> Image [] ("", "")
parPartToInline docx (InternalHyperLink anchor runs) =
  Link (map (runToInline docx) runs) (anchor, "")
parPartToInline docx@(DocX _ _ _ rels _) (ExternalHyperLink relid runs) =
  case lookupRelationship relid rels of
    Just target ->
      Link (map (runToInline docx) runs) (target, "")
    Nothing ->
      Link (map (runToInline docx) runs) ("", "")

parPartsToInlines :: DocX -> [ParPart] -> [Inline]
parPartsToInlines docx parparts =
  bottomUp spanRemove $
  bottomUp (makeImagesSelfContained docx) $ 
  bottomUp spanCorrect $
  bottomUp spanReduce $
  map (parPartToInline docx) parparts

cellToBlocks :: DocX -> Cell -> [Block]
cellToBlocks docx (Cell bps) = map (bodyPartToBlock docx) bps

rowToBlocksList :: DocX -> Row -> [[Block]]
rowToBlocksList docx (Row _ cells) = map (cellToBlocks docx) cells

-- rowToBlocks :: DocX -> Row -> [[Block]]
-- rowToBlocks docx (Row _ cells) =
--   map (\(Cell bps) -> map (bodyPartToBlock docx) bps) cells



bodyPartToBlock :: DocX -> BodyPart -> Block
bodyPartToBlock docx (Paragraph pPr parparts) =
  Div (parStyleToDivAttr pPr) [Para (parPartsToInlines docx parparts)]
bodyPartToBlock docx@(DocX _ _ numbering _ _) (ListItem pPr numId lvl parparts) =
  let
    kvs = case lookupLevel numId lvl numbering of
      Just (_, fmt, txt, Just start) -> [ ("level", lvl)
                                        , ("num-id", numId)
                                        , ("format", fmt)
                                        , ("text", txt)
                                        , ("start", (show start))
                                        ]
      
      Just (_, fmt, txt, Nothing)    -> [ ("level", lvl)
                                        , ("num-id", numId)
                                        , ("format", fmt)
                                        , ("text", txt)
                                        ]
      Nothing                        -> []
  in
   Div
   ("", ["list-item"], kvs)
   [bodyPartToBlock docx (Paragraph pPr parparts)]
bodyPartToBlock _ (Tbl _ _ []) =
  Para []
bodyPartToBlock docx (Tbl cap _ (r:rs)) =
  let caption = strToInlines cap
      (Row rstyle _) = r
      (hdr, rows) = case isTblHdrRow rstyle of
        True -> (Just r, rs)
        False -> (Nothing, r:rs)
      hdrCells = case hdr of
        Just r' -> rowToBlocksList docx r'
        Nothing -> []
      cells = map (rowToBlocksList docx) rows
      
      size = case null hdrCells of
        True -> length $ head cells
        False -> length $ hdrCells

      alignments = take size (repeat AlignDefault)
      widths = take size (repeat 0) :: [Double]
  in
   Table caption alignments widths hdrCells cells

makeImagesSelfContained :: DocX -> Inline -> Inline
makeImagesSelfContained (DocX _ _ _ _ media) i@(Image alt (uri, title)) =
  case lookup uri media of
    Just bs -> case getMimeType uri of
      Just mime ->  let data_uri =
                          "data:" ++ mime ++ ";base64," ++ toString (encode $ BS.concat $ B.toChunks bs)
                    in
                     Image alt (data_uri, title)
      Nothing  -> i
    Nothing -> i
makeImagesSelfContained _ inline = inline

bodyToBlocks :: DocX -> Body -> [Block]
bodyToBlocks docx (Body bps) =
  bottomUp divRemove $
  bottomUp divCorrect $
  bottomUp divReduce $
  bottomUp divCorrectPreReduce $
  bottomUp blocksToDefinitions $
  blocksToBullets $
  map (bodyPartToBlock docx) bps

docxToBlocks :: DocX -> [Block]
docxToBlocks d@(DocX (Document _ body) _ _ _ _) = bodyToBlocks d body

archiveToBlocks :: Archive -> Maybe [Block]
archiveToBlocks archive = do
  docx <- archiveToDocX archive
  return $ docxToBlocks docx

spanReduce :: [Inline] -> [Inline]
spanReduce [] = []
spanReduce ((Span (id1, classes1, kvs1) ils1) : ils)
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

ilToCode :: Inline -> String
ilToCode (Str s) = s
ilToCode _ = ""

spanRemove' :: Inline -> [Inline]
spanRemove' (Span (_, _, kvs) ils) =
  case lookup "underline" kvs of
    Just val -> [Span ("", [], [("underline", val)]) ils]
    Nothing  -> ils
spanRemove' il = [il]

spanRemove :: [Inline] -> [Inline]
spanRemove = concatMap spanRemove'

spanCorrect' :: Inline -> [Inline]
spanCorrect' (Span ("", [], []) ils) = ils
spanCorrect' (Span (ident, classes, kvs) ils)
  | "emph" `elem` classes =
    [Emph $ spanCorrect' $ Span (ident, (delete "emph" classes), kvs) ils]
  | "strong" `elem` classes =
      [Strong $ spanCorrect' $ Span (ident, (delete "strong" classes), kvs) ils]
  | "smallcaps" `elem` classes =
      [SmallCaps $ spanCorrect' $ Span (ident, (delete "smallcaps" classes), kvs) ils]
  | "strikeout" `elem` classes =
      [Strikeout $ spanCorrect' $ Span (ident, (delete "strikeout" classes), kvs) ils]
  | codeSpan `elem` classes =
         [Code (ident, (delete codeSpan classes), kvs) (init $ unlines $ map ilToCode ils)]
  | otherwise =
      [Span (ident, classes, kvs) ils]
spanCorrect' il = [il]

spanCorrect :: [Inline] -> [Inline]
spanCorrect = concatMap spanCorrect'

divReduce :: [Block] -> [Block]
divReduce [] = []
divReduce ((Div (id1, classes1, kvs1) blks1) : blks)
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

isHeaderClass :: String -> Maybe Int
isHeaderClass s | "Heading" `isPrefixOf` s =
  case reads (drop (length "Heading") s) :: [(Int, String)] of
    [] -> Nothing
    ((n, "") : []) -> Just n
    _       -> Nothing
isHeaderClass _ = Nothing

findHeaderClass :: [String] -> Maybe Int
findHeaderClass ss = case mapMaybe id $ map isHeaderClass ss of
  [] -> Nothing
  n : _ -> Just n

blksToInlines :: [Block] -> [Inline]
blksToInlines (Para ils : _) = ils
blksToInlines (Plain ils : _) = ils
blksToInlines _ = []

divCorrectPreReduce' :: Block -> [Block]
divCorrectPreReduce' (Div (ident, classes, kvs) blks)
  | isJust $ findHeaderClass classes =
    let n = fromJust $ findHeaderClass classes
    in
    [Header n (ident, delete ("Heading" ++ (show n)) classes, kvs) (blksToInlines blks)]
  | otherwise = [Div (ident, classes, kvs) blks]
divCorrectPreReduce' blk = [blk]

divCorrectPreReduce :: [Block] -> [Block]
divCorrectPreReduce = concatMap divCorrectPreReduce'

blkToCode :: Block -> String
blkToCode (Para []) = ""
blkToCode (Para ((Code _ s):ils)) = s ++ (blkToCode (Para ils))
blkToCode (Para ((Span (_, classes, _) ils'): ils))
  | codeSpan `elem` classes =
    (init $ unlines $ map ilToCode ils') ++ (blkToCode (Para ils))
blkToCode _ = ""

divRemove' :: Block -> [Block]
divRemove' (Div (_, _, kvs) blks) =
  case lookup "indent" kvs of
    Just val -> [Div ("", [], [("indent", val)]) blks]
    Nothing  -> blks
divRemove' blk = [blk]

divRemove :: [Block] -> [Block]
divRemove = concatMap divRemove'

                                                 
divCorrect' :: Block -> [Block]
divCorrect' (Div (ident, classes, kvs) blks)
  | blockQuoteDiv `elem` classes =
    [BlockQuote [Div (ident, delete blockQuoteDiv classes, kvs) blks]]
  | codeDiv `elem` classes =
    [CodeBlock (ident, (delete codeDiv classes), kvs) (init $ unlines $ map blkToCode blks)]
divCorrect' blk = [blk]

divCorrect :: [Block] -> [Block]
divCorrect = concatMap divCorrect'

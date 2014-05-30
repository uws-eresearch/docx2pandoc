import Text.Pandoc.JSON
import Text.Pandoc
import Data.List

correctSpan :: Inline -> [Inline]
correctSpan (Span ("", [], []) ils) = ils
correctSpan (Span (ident, classes, kvs) ils)
  | "emph" `elem` classes =
    [Emph $ correctSpan $ Span (ident, (delete "emph" classes), kvs) ils]
  | "strong" `elem` classes =
    [Strong $ correctSpan $ Span (ident, (delete "strong" classes), kvs) ils]
  | "smallcaps" `elem` classes =
    [SmallCaps $ correctSpan $ Span (ident, (delete "smallcaps" classes), kvs) ils]
  | "strikeout" `elem` classes =
    [Strikeout $ correctSpan $ Span (ident, (delete "strikeout" classes), kvs) ils]
  | otherwise =
      [Span (ident, classes, kvs) ils]
correctSpan il = [il]

correctDiv :: Block -> [Block]
correctDiv (Div ("", [], []) blks) = blks
correctDiv (Div (_, "list-item" : [], _) blks) = blks
correctDiv blk = [blk]


foo :: Pandoc -> Pandoc
foo = (bottomUp (concatMap correctDiv)) . (bottomUp (concatMap correctSpan))

main = toJSONFilter foo


module Text.Pandoc.DocX.ItemLists ( blocksToBullets
                                  , blocksToDefinitions) where

import Text.Pandoc.JSON
import Control.Monad
import Data.List
import Data.Maybe

isListItem :: Block -> Bool
isListItem (Div (_, classes, _) _) | "list-item" `elem` classes = True
isListItem _ = False

getLevel :: Block -> Maybe Integer
getLevel (Div (_, _, kvs) _) =  liftM read $ lookup "level" kvs
getLevel _ = Nothing

getLevelN :: Block -> Integer
getLevelN b = case getLevel b of
  Just n -> n
  Nothing -> -1

getNumId :: Block -> Maybe Integer
getNumId (Div (_, _, kvs) _) =  liftM read $ lookup "num-id" kvs
getNumId _ = Nothing

getNumIdN :: Block -> Integer
getNumIdN b = case getNumId b of
  Just n -> n
  Nothing -> -1

data ListType = Itemized | Enumerated ListAttributes

listStyleMap :: [(String, ListNumberStyle)]
listStyleMap = [("upperLetter", UpperAlpha),
                ("lowerLetter", LowerAlpha),
                ("upperRoman", UpperRoman),
                ("lowerRoman", LowerRoman),
                ("decimal", Decimal)]

listDelimMap :: [(String, ListNumberDelim)]
listDelimMap = [("%1)", OneParen),
                ("(%1)", TwoParens),
                ("%1.", Period)]

getListType :: Block -> Maybe ListType
getListType b@(Div (_, _, kvs) _) | isListItem b =
  let
    start = lookup "start" kvs
    frmt = lookup "format" kvs
    txt  = lookup "text" kvs
  in
   case frmt of
     Just "bullet" -> Just Itemized
     Just f        ->
       case txt of
         Just t -> Just $ Enumerated (
                  read (fromMaybe "1" start) :: Int,
                  fromMaybe DefaultStyle (lookup f listStyleMap),
                  fromMaybe DefaultDelim (lookup t listDelimMap))
         Nothing -> Nothing
     _ -> Nothing
getListType _ = Nothing
                  
separateBlocks' :: Block -> [[Block]] -> [[Block]]
separateBlocks' blk ([] : []) = [[blk]]
separateBlocks' b@(BulletList _) acc = (init acc) ++ [(last acc) ++ [b]]
separateBlocks' b@(OrderedList _ _) acc = (init acc) ++ [(last acc) ++ [b]]
-- separateBlocks' b acc | getNumIdN b == 1 = (init acc) ++ [(last acc) ++ [b]]
separateBlocks' b acc = acc ++ [[b]]

separateBlocks :: [Block] -> [[Block]]
separateBlocks blks = foldr separateBlocks' [[]] (reverse blks)

flatToBullets' :: Integer -> [(Integer, Block)] -> [Block]
flatToBullets' _ [] = []
flatToBullets' num xs@((n, b) : elems)
  | n == num = b : (flatToBullets' num elems)
  | otherwise = 
    let (children, remaining) = span (\(m, _) -> m > num) xs
    in
     case getListType b of
       Just (Enumerated attr) -> (OrderedList attr (separateBlocks $ flatToBullets' n children)) : (flatToBullets' num remaining)
       _ -> (BulletList (separateBlocks $ flatToBullets' n children)) : (flatToBullets' num remaining)

flatToBullets :: [(Integer, Block)] -> [Block]
flatToBullets elems = flatToBullets' (-1) elems

blocksToBullets :: [Block] -> [Block]
blocksToBullets blks =
  -- bottomUp removeListItemDivs $ 
  flatToBullets $ map (\b -> (getLevelN b, b)) blks


plainParaInlines :: Block -> [Inline]
plainParaInlines (Plain ils) = ils
plainParaInlines (Para ils) = ils
plainParaInlines _ = []

blocksToDefinitions' :: [([Inline], [[Block]])] -> [Block] -> [Block] -> [Block]
blocksToDefinitions' []     acc [] = reverse acc
blocksToDefinitions' defAcc acc [] =
  reverse $ (DefinitionList (reverse defAcc)) : acc
blocksToDefinitions' defAcc acc
  ((Div (_, classes1, _) blks1) : (Div (ident2, classes2, kvs2) blks2) : blks)
  | "DefinitionTerm" `elem` classes1 && "Definition"  `elem` classes2 =
    let remainingAttr2 = (ident2, delete "Definition" classes2, kvs2)
        pair = case remainingAttr2 == ("", [], []) of
          True -> (concatMap plainParaInlines blks1, [blks2])
          False -> (concatMap plainParaInlines blks1, [[Div remainingAttr2 blks2]])
    in
     blocksToDefinitions' (pair : defAcc) acc blks
blocksToDefinitions' defAcc acc
  ((Div (ident2, classes2, kvs2) blks2) : blks)
  | (not . null) defAcc && "Definition"  `elem` classes2 =
    let remainingAttr2 = (ident2, delete "Definition" classes2, kvs2)
        defItems2 = case remainingAttr2 == ("", [], []) of
          True -> blks2
          False -> [Div remainingAttr2 blks2]
        ((defTerm, defItems):defs) = defAcc
        defAcc' = case null defItems of
          True -> (defTerm, [defItems2]) : defs
          False -> (defTerm, init defItems ++ [last defItems ++ defItems2]) : defs
    in
     blocksToDefinitions' defAcc' acc blks
blocksToDefinitions' [] acc (b:blks) =
  blocksToDefinitions' [] (b:acc) blks
blocksToDefinitions' defAcc acc (b:blks) =
  blocksToDefinitions' [] (b : (DefinitionList (reverse defAcc)) : acc) blks


blocksToDefinitions :: [Block] -> [Block]
blocksToDefinitions = blocksToDefinitions' [] []

    
    
    

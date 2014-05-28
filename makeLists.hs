import Text.Pandoc.JSON
import Text.Pandoc
import Control.Monad.State
import Data.List
import Data.Maybe

isListItem :: Block -> Bool
isListItem (Div (_, classes, _) _) | "list-item" `elem` classes = True
isListItem blk = False

getLevel :: Block -> Maybe Int
getLevel b@(Div (_, _, kvs) _) =  liftM read $ lookup "level" kvs
getLevel b = Just (-1)

getLevelN :: Block -> Int
getLevelN b = case getLevel b of
  Just n -> n
  Nothing -> -1

getNumId :: Block -> Maybe Int
getNumId b@(Div (_, _, kvs) _) =  liftM read $ lookup "num-id" kvs

getNumIdN :: Block -> Int
getNumIdN b = case getNumId b of
  Just n -> n
  Nothing -> -1

getBlocks :: Block -> [Block]
getBlocks (Div (_, _, _) blks) = blks
getBlocks blk = [blk]

data ItemList a = Blah a | Item a [ItemList a]


listToTree :: [Block] -> [ItemList Block]
listToTree (blk:ls) | (not . isListItem) blk = (Blah blk) : (listToTree ls)
listToTree (blk:ls) = let
      (children, rest) = span (\b -> (getLevelN b) > (getLevelN blk)) ls
      subtrees = listToTree children
   in Item blk subtrees : listToTree rest
listToTree [] = []

itemToBullet :: ItemList Block -> Block
itemToBullet (Blah blk) = blk
itemToBullet (Item blk subLsts) = BulletList [[blk] ++ (map itemToBullet subLsts)]

foo :: [Block] -> [Block]
foo blks = map itemToBullet $ listToTree blks

bar :: Pandoc -> Pandoc
bar p = bottomUp foo p


main = toJSONFilter bar

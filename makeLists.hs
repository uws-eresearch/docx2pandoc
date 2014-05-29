import Text.Pandoc.JSON
import Text.Pandoc
import Control.Monad.State
import Data.List
import Data.Maybe

isListItem :: Block -> Bool
isListItem (Div (_, classes, _) _) | "list-item" `elem` classes = True
isListItem blk = False

getLevel :: Block -> Maybe Integer
getLevel b@(Div (_, _, kvs) _) =  liftM read $ lookup "level" kvs
getLevel b = Just (-1)

getLevelN :: Block -> Integer
getLevelN b = case getLevel b of
  Just n -> n
  Nothing -> -1

getNumId :: Block -> Maybe Integer
getNumId b@(Div (_, _, kvs) _) =  liftM read $ lookup "num-id" kvs

getNumIdN :: Block -> Integer
getNumIdN b = case getNumId b of
  Just n -> n
  Nothing -> -1

separateBlocks' :: Block -> [[Block]] -> [[Block]]
separateBlocks' blk ([] : []) = [[blk]]
separateBlocks' b@(BulletList _) acc = (init acc) ++ [(last acc) ++ [b]]
separateBlocks' b acc | getNumIdN b == 1 = (init acc) ++ [(last acc) ++ [b]]
separateBlocks' b acc = acc ++ [[b]]

separateBlocks :: [Block] -> [[Block]]
separateBlocks blks = foldr separateBlocks' [[]] blks 

flatToBullets' :: Integer -> [(Integer, Block)] -> [Block]
flatToBullets' _ [] = []
flatToBullets' num xs@((n, b) : elems)
  | n == num = b : (flatToBullets' num elems)
  | otherwise = 
    let (children, remaining) = span (\(m, _) -> m > num) xs
    in
     (BulletList (separateBlocks $ flatToBullets' n children)) : (flatToBullets' num remaining)

flatToBullets :: [(Integer, Block)] -> [Block]
flatToBullets elems = flatToBullets' (-1) elems

blocksToBullets :: [Block] -> [Block]
blocksToBullets blks = flatToBullets $ map (\b -> (getLevelN b, b)) blks
  
pandocToBullets :: Pandoc -> Pandoc
pandocToBullets (Pandoc m blks) = Pandoc m (blocksToBullets blks)


main = toJSONFilter pandocToBullets

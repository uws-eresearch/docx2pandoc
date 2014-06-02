import Codec.Archive.Zip
import Text.XML.Light
import Control.Monad.Reader
import Data.Maybe
import qualified Data.ByteString.Lazy as B

getNameSpace :: String -> Element -> Maybe String
getNameSpace s elem =
  let q = QName s Nothing (Just "xmlns")
  in
   findAttr q elem

getQName :: String -> String -> Element -> Maybe QName
getQName prefix name elem =
  case getNameSpace prefix elem of
    Just uri -> Just $ QName name (Just uri) (Just prefix)
    Nothing  -> Nothing

attrToNSPair :: Attr -> Maybe (String, String)
attrToNSPair (Attr (QName s _ (Just "xmlns")) val) = Just (s, val)
attrToNSPair attr = Nothing 

type NameSpaces = [(String, String)]

data Document = Document NameSpaces Body
          deriving Show

archiveToDocument :: Archive -> Maybe Document
archiveToDocument zf = do
  entry <- findEntryByPath "word/document.xml" zf
  docElem <- (parseXMLDoc . fromEntry) entry
  let namespaces = mapMaybe attrToNSPair (elAttribs docElem) 
  bodyElem <- findChild (QName "body" (lookup "w" namespaces) Nothing) docElem
  body <- elemToBody namespaces bodyElem
  return $ Document namespaces body

data Body = Body [BodyPart]
          deriving Show

elemToBody :: NameSpaces -> Element ->  Maybe Body
elemToBody ns element | qName (elName element) == "body" && qURI (elName element) == (lookup "w" ns) =
  Just $ Body
  $ map fromJust
  $ filter isJust
  $ map (elemToBodyPart ns) $ filterChildrenName isParOrTbl element
  where isParOrTbl :: QName -> Bool
        isParOrTbl q = qName q `elem` ["p", "tbl"] &&
                       qURI q == (lookup "w" ns)
elemToBody _ _ = Nothing

isRunOrLink :: NameSpaces -> QName ->  Bool
isRunOrLink ns q = qName q `elem` ["r", "hyperlink"] &&
                   qURI q == (lookup "w" ns)

isRow :: NameSpaces -> QName ->  Bool
isRow ns q = qName q `elem` ["tr"] &&
             qURI q == (lookup "w" ns)

-- isQNameNS :: [(String, String)] -> NameSpaces -> QName -> Bool
-- isQNameNS prefixPairs ns q = (qURI q, qName q) `elem`
--                        (map (\(s, t) -> (lookup s ns, t)) prefixPairs)

-- findChildrenNS :: [(String, String)] -> NameSpaces -> Element -> [Element]
-- findChildrenNS prefixPairs ns elem =
--   findChildren (isQNameNS prefixPairs ns) elem

elemToBodyPart :: NameSpaces -> Element ->  Maybe BodyPart
elemToBodyPart ns element
  | qName (elName element) == "p" &&
    qURI (elName element) == (lookup "w" ns) =
      Just 
      $ Paragraph []
      $ map fromJust
      $ filter isJust
      $ (map (elemToParPart ns)
         $ filterChildrenName (isRunOrLink ns) element)
  | qName (elName element) == "tbl" &&
    qURI (elName element) == (lookup "w" ns) =
      Just
      $ Table []
      $ map fromJust
      $ filter isJust
      $ (map (elemToRow ns)
         $ filterChildrenName (isRow ns) element)
  | otherwise = Nothing

data BodyPart = Paragraph Style [ParPart]
              | Table Style [Row]
              deriving Show

data Row = Row Style [Cell]
           deriving Show

elemToRow :: NameSpaces -> Element -> Maybe Row
elemToRow ns elem = Nothing

data Cell = Cell Style [BodyPart]
            deriving Show

data ParPart = PlainRun Run
             | InternalHyperLink Target [Run]
             | ExternalHyperLink Target [Run]
             deriving Show

elemToParPart :: NameSpaces -> Element -> Maybe ParPart
elemToParPart ns elem
  | qName (elName elem) == "r" &&
    qURI (elName elem) == (lookup "w" ns) =
      case findChild (QName "t" (lookup "w" ns) (Just "w")) elem of
        Just t -> Just $ PlainRun $ Run [] (strContent t)
        Nothing -> Just $ PlainRun $ Run [] ""
elemToParPart _ _ = Nothing



data Run = Run Style String
           deriving Show

type NameSpace = Reader Element

elemToRun :: Element -> NameSpace (Maybe Run)
elemToRun elem = do
  body <- ask
  let text = getQName "w" "t" body >>= (\q -> findChild q elem)
    in
   return $ case text of
     Nothing -> Nothing
     Just e  -> Just $ Run [] (strContent e)

type Target = String
type Style = [String]
               

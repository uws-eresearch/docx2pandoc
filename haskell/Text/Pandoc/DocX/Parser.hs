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

isQNameNS :: [(String, String)] -> NameSpaces -> QName -> Bool
isQNameNS prefixPairs ns q = (qURI q, qName q) `elem`
                       (map (\(s, t) -> (lookup s ns, t)) prefixPairs)

findChildrenNS :: [(String, String)] -> NameSpaces -> Element -> [Element]
findChildrenNS prefixPairs ns elem =
  filterChildrenName (isQNameNS prefixPairs ns) elem

type NameSpaces = [(String, String)]

data Document = Document NameSpaces Body
          deriving Show

data Notes = Notes NameSpaces [(String, [BodyPart])] [(String, [BodyPart])]
           deriving Show

getFootNote :: String -> Notes -> Maybe [BodyPart]
getFootNote s (Notes _ fns _) = lookup s fns

getEndNote :: String -> Notes -> Maybe [BodyPart]
getEndNote s (Notes _ _ ens) = lookup s ens

filePathToBPs :: FilePath -> IO (Maybe [BodyPart])
filePathToBPs fp = do
  f <- B.readFile fp
  let archive = toArchive f
  return $ do
    Document ns (Body bps) <- archiveToDocument archive
    return bps

isParagraph :: BodyPart -> Bool
isParagraph (Paragraph _ _) = True
isParagraph _ = False

findStyledP :: BodyPart -> Bool
findStyledP (Paragraph pPr _) = (isJust $ pStyle pPr) || (isJust $ indent pPr)
findStyledP bp = False

archiveToDocument :: Archive -> Maybe Document
archiveToDocument zf = do
  entry <- findEntryByPath "word/document.xml" zf
  docElem <- (parseXMLDoc . fromEntry) entry
  let namespaces = mapMaybe attrToNSPair (elAttribs docElem) 
  bodyElem <- findChild (QName "body" (lookup "w" namespaces) Nothing) docElem
  body <- elemToBody namespaces bodyElem
  return $ Document namespaces body

noteElemToNote :: NameSpaces -> Element -> Maybe (String, [BodyPart])
noteElemToNote ns element
  | qName (elName element) `elem` ["endnote", "footnote"] &&
    qURI (elName element) == (lookup "w" ns) =
      do
        id <- findAttr (QName "id" (lookup "w" ns) (Just "w")) element
        let bps = map fromJust
                  $ filter isJust
                  $ map (elemToBodyPart ns)
                  $ filterChildrenName (isParOrTbl ns) element
        return $ (id, bps)
noteElemToNote ns element = Nothing

elemToNotes :: NameSpaces -> String -> Element -> Maybe [(String, [BodyPart])]
elemToNotes ns notetype element
  | qName (elName element) == (notetype ++ "s") &&
    qURI (elName element) == (lookup "w" ns) =
      Just $ map fromJust
      $ filter isJust
      $ map (noteElemToNote ns)
      $ findChildren (QName notetype (lookup "w" ns) (Just "w")) element
elemToNotes ns notetype element = Nothing
  

archiveToNotes :: Archive -> Maybe Notes
archiveToNotes zf = do
  fn_entry <- findEntryByPath "word/footnotes.xml" zf
  en_entry <- findEntryByPath "word/endnotes.xml" zf
  fnElem <- (parseXMLDoc . fromEntry) fn_entry
  enElem <- (parseXMLDoc . fromEntry) en_entry
  let fn_namespaces = mapMaybe attrToNSPair (elAttribs fnElem)
      en_namespaces = mapMaybe attrToNSPair (elAttribs fnElem)
      namespaces = fn_namespaces ++ en_namespaces
  fn <- (elemToNotes namespaces "footnote" fnElem)
  en <- (elemToNotes namespaces "endnote" enElem)
  return
    $ Notes namespaces fn en
    
    
    



data Body = Body [BodyPart]
          deriving Show

isParOrTbl :: NameSpaces -> QName -> Bool
isParOrTbl ns q = qName q `elem` ["p", "tbl"] &&
                  qURI q == (lookup "w" ns)

elemToBody :: NameSpaces -> Element ->  Maybe Body
elemToBody ns element | qName (elName element) == "body" && qURI (elName element) == (lookup "w" ns) =
  Just $ Body
  $ map fromJust
  $ filter isJust
  $ map (elemToBodyPart ns) $ filterChildrenName (isParOrTbl ns) element
elemToBody _ _ = Nothing

isRunOrLink :: NameSpaces -> QName ->  Bool
isRunOrLink ns q = qName q `elem` ["r", "hyperlink"] &&
                   qURI q == (lookup "w" ns)

isRow :: NameSpaces -> QName ->  Bool
isRow ns q = qName q `elem` ["tr"] &&
             qURI q == (lookup "w" ns)


elemToBodyPart :: NameSpaces -> Element ->  Maybe BodyPart
elemToBodyPart ns element
  | qName (elName element) == "p" &&
    qURI (elName element) == (lookup "w" ns) =
      Just 
      $ Paragraph (elemToParagraphStyle ns element)
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

data ParagraphStyle = ParagraphStyle { pStyle :: Maybe String
                                     , indent :: Maybe Integer
                                     }
                      deriving Show

defaultParagraphStyle :: ParagraphStyle
defaultParagraphStyle = ParagraphStyle { pStyle = Nothing
                                       , indent = Nothing
                                       }

elemToParagraphStyle :: NameSpaces -> Element -> ParagraphStyle
elemToParagraphStyle ns elem =
  case findChild (QName "pPr" (lookup "w" ns) (Just "w")) elem of
    Just pPr ->
      ParagraphStyle
      {pStyle =
        findChild (QName "pStyle" (lookup "w" ns) (Just "w")) pPr >>=
        findAttr (QName "val" (lookup "w" ns) (Just "w"))
      , indent =
        findChild (QName "ind" (lookup "w" ns) (Just "w")) pPr >>=
        findAttr (QName "left" (lookup "w" ns) (Just "w")) >>=
        (\s -> listToMaybe (map fst (reads s :: [(Integer, String)])))
        }
    Nothing -> defaultParagraphStyle


data BodyPart = Paragraph ParagraphStyle [ParPart]
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

data Run = Run RunStyle String
         | Footnote String [BodyPart]
         | Endnote String [BodyPart]
           deriving Show

data RunStyle = RunStyle { isBold :: Bool
                         , isItalic :: Bool
                         , isSmallCaps :: Bool
                         , isStrike :: Bool
                         , underline :: Maybe String
                         , rStyle :: Maybe String }
                deriving Show

defaultRunStyle :: RunStyle
defaultRunStyle = RunStyle { isBold = False
                           , isItalic = False
                           , isSmallCaps = False
                           , isStrike = False
                           , underline = Nothing
                           , rStyle = Nothing
                           }

elemToRunStyle :: NameSpaces -> Element -> RunStyle
elemToRunStyle ns elem =
  case findChild (QName "rPr" (lookup "w" ns) (Just "w")) elem of
    Just rPr ->
      RunStyle
      {
        isBold = isJust $ findChild (QName "b" (lookup "w" ns) (Just "w")) rPr
      , isItalic = isJust $ findChild (QName "i" (lookup "w" ns) (Just "w")) rPr
      , isSmallCaps = isJust $ findChild (QName "smallCaps" (lookup "w" ns) (Just "w")) rPr
      , isStrike = isJust $ findChild (QName "strike" (lookup "w" ns) (Just "w")) rPr
      , underline =
        findChild (QName "u" (lookup "w" ns) (Just "w")) rPr >>=
        findAttr (QName "val" (lookup "w" ns) (Just "w"))
      , rStyle =
        findChild (QName "rStyle" (lookup "w" ns) (Just "w")) rPr >>=
        findAttr (QName "val" (lookup "w" ns) (Just "w"))
        }
    Nothing -> defaultRunStyle



elemToParPart :: NameSpaces -> Element -> Maybe ParPart
elemToParPart ns elem
  | qName (elName elem) == "r" &&
    qURI (elName elem) == (lookup "w" ns) =
      case
        findChild (QName "footnoteReference" (lookup "w" ns) (Just "w")) elem >>=
        findAttr (QName "id" (lookup "w" ns) (Just "w"))
      of
        Just s -> Just $ PlainRun $ Footnote s []
        Nothing ->
          case
            findChild (QName "endnoteReference" (lookup "w" ns) (Just "w")) elem >>=
            findAttr (QName "id" (lookup "w" ns) (Just "w"))
          of
            Just s -> Just $ PlainRun $ Endnote s []
            Nothing ->  case
              findChild (QName "t" (lookup "w" ns) (Just "w")) elem
              of
                Just t -> Just $ PlainRun $ Run (elemToRunStyle ns elem) (strContent t)
                Nothing -> Just $ PlainRun $ Run (elemToRunStyle ns elem) ""
elemToParPart _ _ = Nothing


type NameSpace = Reader Element

-- elemToRun :: Element -> NameSpace (Maybe Run)
-- elemToRun elem = do
--   body <- ask
--   let text = getQName "w" "t" body >>= (\q -> findChild q elem)
--     in
--    return $ case text of
--      Nothing -> Nothing
--      Just e  -> Just $ Run [] (strContent e)

type Target = String
type Style = [String]
               

import Codec.Archive.Zip
import Text.XML.Light
import Control.Monad.Reader
import Data.Maybe
import Data.List
import System.FilePath
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

data DocX = DocX Document Notes Numbering [Relationship]
          deriving Show

archiveToDocX :: Archive -> Maybe DocX
archiveToDocX archive = do
  doc <- archiveToDocument archive
  notes <- archiveToNotes archive
  numbering <- archiveToNumbering archive
  let rels = archiveToRelationships archive
  return $ DocX doc notes numbering rels

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

data Numbering = Numbering NameSpaces [Numb] [AbstractNumb]
                 deriving Show

data Numb = Numb String String           -- right now, only a key to an abstract num
            deriving Show

data AbstractNumb = AbstractNumb String [Level]
                    deriving Show

type Level = (String, String, String, Maybe Integer)

numElemToNum :: NameSpaces -> Element -> Maybe Numb
numElemToNum ns element |
  qName (elName element) == "num" &&
  qURI (elName element) == (lookup "w" ns) = do
    numId <- findAttr (QName "numId" (lookup "w" ns) (Just "w")) element
    absNumId <- findChild (QName "abstractNumId" (lookup "w" ns) (Just "w")) element
                >>= findAttr (QName "val" (lookup "w" ns) (Just "w"))
    return $ Numb numId absNumId
numElemToNum _ _ = Nothing

absNumElemToAbsNum :: NameSpaces -> Element -> Maybe AbstractNumb
absNumElemToAbsNum ns element |
  qName (elName element) == "abstractNum" &&
  qURI (elName element) == (lookup "w" ns) = do
    absNumId <- findAttr
                (QName "abstractNumId" (lookup "w" ns) (Just "w"))
                element
    let levelElems = findChildren
                 (QName "lvl" (lookup "w" ns) (Just "w"))
                 element
        levels = mapMaybe id $ map (levelElemToLevel ns) levelElems
    return $ AbstractNumb absNumId levels
absNumElemToNum _ _ = Nothing

levelElemToLevel :: NameSpaces -> Element -> Maybe Level
levelElemToLevel ns element |
    qName (elName element) == "lvl" &&
    qURI (elName element) == (lookup "w" ns) = do
      ilvl <- findAttr (QName "ilvl" (lookup "w" ns) (Just "w")) element
      fmt <- findChild (QName "numFmt" (lookup "w" ns) (Just "w")) element
             >>= findAttr (QName "val" (lookup "w" ns) (Just "w"))
      txt <- findChild (QName "lvlText" (lookup "w" ns) (Just "w")) element
             >>= findAttr (QName "val" (lookup "w" ns) (Just "w"))
      let start = findChild (QName "start" (lookup "w" ns) (Just "w")) element
                  >>= findAttr (QName "val" (lookup "w" ns) (Just "w"))
                  >>= (\s -> listToMaybe (map fst (reads s :: [(Integer, String)])))
      return (ilvl, fmt, txt, start)

archiveToNumbering :: Archive -> Maybe Numbering
archiveToNumbering zf = do
  entry <- findEntryByPath "word/numbering.xml" zf
  numberingElem <- (parseXMLDoc . fromEntry) entry
  let namespaces = mapMaybe attrToNSPair (elAttribs numberingElem)
      numElems = findChildren
                 (QName "num" (lookup "w" namespaces) (Just "w"))
                 numberingElem
      absNumElems = findChildren
                    (QName "abstracNum" (lookup "w" namespaces) (Just "w"))
                    numberingElem
      nums = mapMaybe id $ map (numElemToNum namespaces) numElems
      absNums = mapMaybe id $ map (absNumElemToNum namespaces) absNumElems
  return $ Numbering namespaces nums absNums

data Notes = Notes NameSpaces [(String, [BodyPart])] [(String, [BodyPart])]
           deriving Show

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

getFootNote :: String -> Notes -> Maybe [BodyPart]
getFootNote s (Notes _ fns _) = lookup s fns

getEndNote :: String -> Notes -> Maybe [BodyPart]
getEndNote s (Notes _ _ ens) = lookup s ens

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
      namespaces = unionBy (\x y -> fst x == fst y) fn_namespaces en_namespaces
  fn <- (elemToNotes namespaces "footnote" fnElem)
  en <- (elemToNotes namespaces "endnote" enElem)
  return
    $ Notes namespaces fn en

data Relationship = Relationship (String, String)
                  deriving Show

filePathIsRel :: FilePath -> Bool
filePathIsRel fp =
  let (dir, name) = splitFileName fp
  in
   (dir == "word/_rels") && ((takeExtension name) == ".rel")

relElemToRelationship :: Element -> Maybe Relationship
relElemToRelationship element | qName (elName element) == "Relationship" =
  do
    relId <- findAttr (QName "Id" Nothing Nothing) element
    target <- findAttr (QName "Target" Nothing Nothing) element
    return $ Relationship (relId, target)
relElemToRelationship _ = Nothing
  

archiveToRelationships :: Archive -> [Relationship]
archiveToRelationships archive = 
  let relPaths = filter filePathIsRel (filesInArchive archive)
      entries  = map fromJust $ filter isJust $ map (\f -> findEntryByPath f archive) relPaths
      relElems = map fromJust $ filter isJust $ map (parseXMLDoc . fromEntry) entries
      rels = map fromJust $ filter isJust $ map relElemToRelationship $ concatMap elChildren relElems
  in
   rels
   


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
             | InternalHyperLink Anchor [Run]
             | ExternalHyperLink RelId [Run]
             deriving Show

data Run = Run RunStyle String
         | Footnote String 
         | Endnote String 
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

elemToRun :: NameSpaces -> Element -> Maybe Run
elemToRun ns element
  | qName (elName element) == "r" &&
    qURI (elName element) == (lookup "w" ns) =
      case
        findChild (QName "footnoteReference" (lookup "w" ns) (Just "w")) element >>=
        findAttr (QName "id" (lookup "w" ns) (Just "w"))
      of
        Just s -> Just $ Footnote s
        Nothing ->
          case
            findChild (QName "endnoteReference" (lookup "w" ns) (Just "w")) element >>=
            findAttr (QName "id" (lookup "w" ns) (Just "w"))
          of
            Just s -> Just $ Endnote s 
            Nothing ->  case
              findChild (QName "t" (lookup "w" ns) (Just "w")) element
              of
                Just t -> Just $ Run (elemToRunStyle ns element) (strContent t)
                Nothing -> Just $ Run (elemToRunStyle ns element) ""
elemToRun _ _ = Nothing


elemToParPart :: NameSpaces -> Element -> Maybe ParPart
elemToParPart ns element
  | qName (elName element) == "r" &&
    qURI (elName element) == (lookup "w" ns) =
      do
        r <- elemToRun ns element
        return $ PlainRun r
elemToParPart ns element
  | qName (elName element) == "hyperlink" &&
    qURI (elName element) == (lookup "w" ns) =
      let runs = map fromJust $ filter isJust $ map (elemToRun ns)
                 $ findChildren (QName "r" (lookup "w" ns) (Just "w")) element
      in
       case findAttr (QName "anchor" (lookup "w" ns) (Just "w")) element of
         Just anchor ->
          Just $ InternalHyperLink anchor runs
         Nothing ->
           case findAttr (QName "id" (lookup "r" ns) (Just "r")) element of
             Just relId -> Just $ ExternalHyperLink relId runs
             Nothing    -> Nothing
elemToParPart _ _ = Nothing

type NameSpace = Reader Element

type Anchor = String
type RelId = String
type Style = [String]
               

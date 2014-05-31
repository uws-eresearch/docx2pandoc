import Codec.Archive.Zip
import Text.XML.Light
import qualified Data.ByteString.Lazy as B

data DocX = DocX Body 
          deriving Show

archiveToDocX :: Archive -> Maybe [Content]
archiveToDocX zf = do
  entry <- findEntryByPath "word/document.xml" zf
  return $ (parseXML . fromEntry) entry

data Body = Body [BodyPart]
          deriving Show

data BodyPart = Paragraph Style [ParPart]
              | Table Style [Row]
              deriving Show

data Row = Row Style [Cell]
           deriving Show

data Cell = Cell Style [BodyPart]
            deriving Show

data Run = Run Style String
           deriving Show

data ParPart = PlainRun Run
             | InternalHyperLink Target [Run]
             | ExternalHyperLink Target [Run]
             deriving Show

getNameSpace :: String -> Element -> Maybe String
getNameSpace s elem =
  let q = QName s Nothing (Just "xmlns")
  in
   findAttr q elem

type Target = String
type Style = [String]
               

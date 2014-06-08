import Text.Pandoc.DocX.Pandoc
import Codec.Archive.Zip
import qualified Data.ByteString.Lazy as B
import Text.Pandoc.JSON
import Text.Pandoc
import qualified Text.Pandoc.UTF8 as UTF8
import System.Environment

mkPandoc :: FilePath -> IO (Maybe [Block])
mkPandoc fp = do
  f <- B.readFile fp
  return $ archiveToBlocks (toArchive f) 

main :: IO ()
main = do
  args <- getArgs
  let fp = case args of
        [] -> error "A docx file is required"
        _  -> head args
  output <- mkPandoc fp
  case output of
    Just blks -> UTF8.putStr $ writeJSON def (Pandoc nullMeta blks)
    Nothing   -> error ("Couldn't parse docx file " ++ fp)

    

           



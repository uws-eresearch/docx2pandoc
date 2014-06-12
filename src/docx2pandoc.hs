import Text.Pandoc.Readers.DocX
import qualified Data.ByteString.Lazy as B
import Text.Pandoc.JSON
import Text.Pandoc
import qualified Text.Pandoc.UTF8 as UTF8
import System.Environment

mkPandoc :: FilePath -> IO Pandoc
mkPandoc fp = do
  f <- B.readFile fp
  return $ readDocX def f

main :: IO ()
main = do
  args <- getArgs
  let fp = case args of
        [] -> error "A docx file is required"
        _  -> head args
  pandoc <- mkPandoc fp
  UTF8.putStr $ writeJSON def pandoc

    

           



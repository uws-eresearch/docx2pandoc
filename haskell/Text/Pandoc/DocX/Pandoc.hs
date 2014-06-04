import Text.Pandoc
import Text.Pandoc.JSON
import Text.Pandoc.DocX.Parser
import Data.Maybe

runStyleToSpanAttr :: RunStyle -> (String, [String], [(String, String)])
runStyleToSpanAttr rPr = ("",
                          mapMaybe id [
                            if isBold rPr then (Just "strong") else Nothing,
                            if isItalic rPr then (Just "emph") else Nothing,
                            if isSmallCaps rPr then (Just "smallcaps") else Nothing,
                            if isStrike rPr then (Just "strike") else Nothing,
                            rStyle],
                          case underline rPr of
                            Just fmt -> [("underline", fmt)]
                            _        -> []
                         )



FILES = docx2pandoc.hs \
	Text/Pandoc/DocX/Parser.hs \
	Text/Pandoc/DocX/Pandoc.hs
GHC_FLAGS = -O2 -Wall -fno-warn-unused-do-bind

docx2pandoc: $(FILES)
	ghc --make $(GHC_FLAGS) $<

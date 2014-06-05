DOCX2PANDOC = docx2pandoc.hs
LIB_FILES = Text/Pandoc/DocX/Parser.hs \
	    Text/Pandoc/DocX/Pandoc.hs
GHC_FLAGS = -O2 -Wall -fno-warn-unused-do-bind

.PHONY: all
all: docx2pandoc


.PHONY: clean
clean:
	rm $(DOCX2PANDOC:.hs=) $(DOCX2PANDOC:.hs=.hi) $(DOCX2PANDOC:.hs=.o) $(LIB_FILES:.hs=.hi) $(LIB_FILES:.hs=.o)

docx2pandoc: $(DOCX2PANDOC) $(LIBFILES)
	ghc --make $(GHC_FLAGS) $<

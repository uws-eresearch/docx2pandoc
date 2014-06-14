DOCX2PANDOC = docx2pandoc.hs
LIB_FILES = Text/Pandoc/MIME.hs \
            Text/Pandoc/Readers/DocX/Parse.hs \
	    Text/Pandoc/Readers/DocX/Lists.hs \
            Text/Pandoc/Readers/DocX.hs
HS_FILES = $(DOCX2PANDOC) $(LIB_FILES)
SRC_HS_FILES = $(foreach f,$(HS_FILES), src/$(f))
GHC_FLAGS = -O2 -Wall -fno-warn-unused-do-bind

.PHONY: all
all: docx2pandoc


.PHONY: clean
clean:
	rm -f $(DOCX2PANDOC:.hs=) $(SRC_HS_FILES:.hs=.hi) $(SRC_HS_FILES:.hs=.o)

.PHONY: tests
tests: docx2pandoc
	cd test; python2 test.py

docx2pandoc: $(SRC_HS_FILES)
	ghc -isrc src/$(DOCX2PANDOC) --make $(GHC_FLAGS)
	mv src/$(DOCX2PANDOC:.hs=) .

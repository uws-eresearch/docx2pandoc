docx2pandoc
===========

This package provides an executable and the beginnings of a library
for translating from Microsoft Word OOXML to pandoc output
formats. Though the source is still a bit messy and uncommented, the
goal is for this to be an official reader for pandoc, enabling users
to translate from MS Word files.

Usage
-----

`docx2pandoc` outputs pandoc JSON, so to use it, just pipe it to
`pandoc` with the desired output format. For example:

~~~
docx2pandoc input.docx | pandoc -Ss -f json -t html 
~~~

For the time being, docx2pandoc does not embed images into the output
files, since this can result in size and memory problems. It refers
instead to an image in the `word/media` file in the docx archive. To
make this available to your output file, run something like the
following, in your output directory

~~~
unzip input.docx "word/media/*"
~~~

Building
--------

Just `cd` into the `docx2pandoc` dir and run `make`. It doesn't
require any libraries not required by `pandoc`, so if you can build
that you should be fine.

To run the test suite, just run `make tests`.

Some words on how it works
--------------------------

OOXML is weird, and it doesn't really nest. It splits documents into
"paragraphs," and paragraphs into "runs." Each can have a body of
text, with its own style. So one run will be bold, the next will be
bold and italic, and the next will be bold. This maps rather awkwardly
onto pandoc's AST.

So the reader works in a few steps (glossing over some subtleties):

 1. Reads the xml and converts it into a haskell `DocX` type that
    mirrors the docx structure (this is defined in the file
    `Parser.hs`, if you're interested). It doesn't keep anywhere close
    to all the info, but it tries to keep most of the stuff that will
    be useful for pandoc, or could be useful down the line.

 2. Convert that into a legal, but weird pandoc format. This will have
    a *lot* of `div`s and `span`s, mirroring the run styles. These
    will be non-nesting.

 3. Deal with lists, which are completely bizarre in docx.

 4. Reduce the adjacent `span`s and `div`s to produce proper nesting.

 5. Translate classes like "emph" and "strong" to pandoc inlines. Same
    with divs like "SourceCode" and "BlockQuote."

There is some further information that is kept for the purpose of
scripting, but which the reader doesn't do anything with. Notably, it
keeps an `indent` attribute on divs that specify it, due to the
fact that something like 98% of docx files in the wild use indentation
to define block quotes.

Contributions
-------------

Please! Pull requests or patches welcome. I'll comment up the code
soon to make it easier.





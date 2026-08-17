Real HTML pages from the govdocs1 corpus, kept as regression fixtures.

WHY THESE ARE HERE. Every other DOCX/XLSX/PPTX/HTML fixture in this suite is built
by the code under test, so the suite could only ever find defects in things it
already knew how to produce. That is not a hypothetical limitation: 1,284 tests
were green while HTML to PDF succeeded on 58.6% of real pages, and every defect
behind that gap was invisible here.

Each file is the SMALLEST real page in the corpus that exposed one specific
defect. They are kept whole - not reduced - because a reduced page is a page this
project wrote, which is the property being avoided.

  rowspan-past-last-row.html   a rowspan reaching past the last row (govdocs1 000436)
  old-style-name-anchor.html   internal links whose targets are <a name> (000814)
  spacer-cell.html             a whitespace-only cell beside long text (000415)
  image-only-link.html         a link wrapping only an image, in a cell (000061)

PROVENANCE AND LICENCE. govdocs1 is a public corpus of documents crawled from the
.gov domain, published by Digital Corpora for exactly this kind of format and
converter testing. As works of the United States government they are in the public
domain, which is why they can sit in a public repository.

  https://digitalcorpora.org/corpora/file-corpora/files/

DO NOT EDIT THEM. Fixing the markup would destroy the only thing they are for. If
one of these stops exposing its defect, that is news - the test says so.

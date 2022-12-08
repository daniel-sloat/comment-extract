from functools import cached_property
import re

# from docx.ooxml_ns import ns
from lxml.etree import QName
from reprlib import Repr

limit = Repr()


class DOCXElement:
    def __init__(self, element):
        self.element = element

    def __repr__(self):
        return f"{self.__class__.__name__}(tag='{(self.tag)}',text={limit.repr(self.text)})"

    def __iter__(self):
        return iter(self.element)

    def __len__(self):
        return len(self.element)

    def __getitem__(self, key):
        return self.element.xpath(f"{key}", **ns)

    @property
    def tag(self):
        _tag = QName(self.element).localname
        return _tag if _tag != "id" else "_id"

    @property
    def text(self):
        return self.element.xpath("string(.)")


from lxml.builder import ElementMaker

ns_maker = {
    "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "nsmap": {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
}

ns = {"namespaces": ns_maker["nsmap"]}

E = ElementMaker(**ns_maker)

RUN = E.r
PARA = E.p
BODY = E.body
DOC = E.document
COMMENTSTART = E.commentRangeStart
COMMENTEND = E.commentRangeEnd
TEXT = E.t
RUNPROPS = E.rPr
SCRIPTHEIGHT = E.vertAlign

run_text = "Run text that is getting pretty long!"

different_level_starts_and_ends = DOC(
    BODY(
        COMMENTSTART(val="1"),
        PARA(
            E.pPr(SCRIPTHEIGHT(val="2")),
            RUN(RUNPROPS(SCRIPTHEIGHT(val="2")), TEXT(run_text)),
            RUN(TEXT(run_text)),
            COMMENTEND(val="1"),
        ),
        PARA(
            RUN(TEXT(run_text)),
            COMMENTSTART(val="2"),
            RUN(RUNPROPS(SCRIPTHEIGHT(val="2")), TEXT(run_text)),
            RUN(TEXT(run_text)),
        ),
        COMMENTEND(val="2"),
    ),
)


for el in different_level_starts_and_ends.xpath("//w:p", **ns):
    prop = DOCXElement(el)
    if prop:
        print(prop.text)

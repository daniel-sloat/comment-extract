from collections import ChainMap
from functools import cache, cached_property

from docx.elements.prop_decode import PropDecode
from docx.ooxml_ns import ns


@cache
class Properties:
    def __init__(self, run):
        self._parent = run
        self._element = run.element
        self._styles = self._parent._parent._doc.styles.styles
        self.rstyle = self._element.xpath("string(w:rPr/w:rStyle/@w:val)", **ns)

    @cached_property
    def pstyle(self):
        return self._element.xpath("string(parent::w:p/w:pPr/w:pStyle/@w:val)", **ns)

    @cached_property
    def chain(self):
        rstyle_run_props = self._styles[self.rstyle].run if self.rstyle else {}
        pstyle_run_props = self._styles[self.pstyle].run if self.pstyle else {}
        return ChainMap(self._parent._props, rstyle_run_props, pstyle_run_props)

    @property
    def decode(self):
        return PropDecode(self.chain)

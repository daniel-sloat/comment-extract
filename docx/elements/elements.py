from functools import cached_property
import re
from docx.ooxml_ns import ns


class BaseDOCXElement:
    def __init__(self, element):
        self.element = element

    def __repr__(self):
        return f"{self.__class__.__name__}(tag='{self.tag}')"

    @property
    def tag(self):
        return self._strip_uri(self.element.tag)

    @cached_property
    def attrib(self):
        return {self._strip_uri(tag): el for tag, el in self.element.items()}

    @staticmethod
    def _strip_uri(tag):
        fix_id = re.sub("id", "_id", tag if tag else "")
        return re.sub("{.*}", "", fix_id)


class TextElement(BaseDOCXElement):
    def __repr__(self):
        disp = 32
        return (
            f"{self.__class__.__name__}("
            f"text='{self.text[0:disp]}"
            f"{'...' if self.text and len(self.text) > disp else ''}"
            f"')"
        )

    @property
    def text(self):
        return self.element.xpath("string(.)")


class PropElement(BaseDOCXElement):
    pass

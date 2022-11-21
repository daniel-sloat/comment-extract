from functools import cached_property
from docx.baseelement import BaseDOCXElement
from ..ooxml_ns import ns


class CommentsExt:
    def __init__(self, document):
        self._doc = document

    # @cached_property
    @property
    def exts(self):
        if (xml := self._doc.xml.get("word/commentsExtended.xml")) is not None:
            return [CommentExtElement(el) for el in xml.xpath("./w15:commentEx", **ns)]
        return []

    @cached_property
    def ancestors(self):
        d = {}
        for ext in self.exts:
            for _id, data in self._doc.comments.metadata().items():
                if data["para_id"] == ext.attrib.get("paraId"):
                    for _id2, data2 in self._doc.comments.metadata().items():
                        if data2["para_id"] == ext.attrib.get("paraIdParent"):
                            d[_id] = _id2

        e = {}
        for k, v in d.items():
            e[k] = v  # TODO LOOK UP COMMENT FOR ANCESTOR

        return e

    @property
    def para_ids(self):
        if (xml := self._doc.xml.get("word/commentsExtended.xml")) is not None:
            return xml.xpath(
                "/w15:commentsEx/w15:commentEx/@w15:paraId[../@w15:paraIdParent]",
                **ns,
            )
        return []


class CommentExtElement(BaseDOCXElement):
    def __init__(self, element):
        super().__init__(element)

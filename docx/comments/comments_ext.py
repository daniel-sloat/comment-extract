from ..ooxml_ns import ns


class CommentsExt:
    def __init__(self, document):
        self._doc = document

    @property
    def _comment_ext_xml(self):
        return self._doc.xml.get("word/commentsExtended.xml")

    @property
    def _parents(self):
        if self._comment_ext_xml is not None:
            return [
                para_id
                for para_id in self._comment_ext_xml.xpath(
                    "/w15:commentsEx/w15:commentEx/@w15:paraId[../@w15:paraIdParent]",
                    **ns
                )
            ]
        return []

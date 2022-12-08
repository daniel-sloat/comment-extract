from functools import cached_property
from docx.ooxml_ns import ns
from docx.comments.comment_data import CommentBounds


class CommentsDocument:
    def __init__(self, document):
        self._doc = document
        self.xml = self._doc.xml["word/document.xml"]
        self.comment_range_elements = self.xml.xpath(
            "./w:body/w:p/w:commentRangeStart|"
            "./w:body/w:commentRangeStart|"
            "./w:body/w:p/w:commentRangeEnd|"
            "./w:body/w:commentRangeEnd",
            **ns,
        )

    # @cached_property
    # def comment_bounds(self):
    #     comment_ids = [
    #         id
    #         for id in self.xml.xpath(
    #             "//w:commentRangeStart/@w:id",
    #             **ns,
    #         )
    #     ]
    #     comment_bounds = {}
    #     for _id in comment_ids:
    #         start_and_end_nodes = self.xml.xpath(
    #             f"//w:commentRangeStart[@w:id={_id}]|//w:commentRangeEnd[@w:id={_id}]",
    #             **ns,
    #         )
    #         comment_bounds[_id] = start_and_end_nodes
    #     return {_id: CommentBounds(comment_bounds[_id][0], comment_bounds[_id][1]).asdict() for _id in comment_ids}

    @cached_property
    def comment_bounds(self):
        starts = {
            element.xpath("string(@w:id)", **ns): element
            for element in self.comment_range_elements
            if element.xpath("self::w:commentRangeStart", **ns)
        }
        ends = {
            element.xpath("string(@w:id)", **ns): element
            for element in self.comment_range_elements
            if element.xpath("self::w:commentRangeEnd", **ns)
        }
        return {_id: CommentBounds(starts[_id], ends[_id]).asdict() for _id in starts}

    # @property
    # def comment_runs_inside(self):
    #     # This works for comments inside paragraphs
    #     return self._doc.xml["word/document.xml"].xpath(
    #         "w:body/w:p/w:commentRangeStart[@w:id=$id]/following-sibling::*[not(self::w:commentRangeEnd[@w:id=$id])]",
    #         id="0",
    #         **ns,
    #     )

    # @property
    # def comment_runs_outside(self):
    #     # This works for comments outside paragraphs
    #     return self._doc.xml["word/document.xml"].xpath(
    #         "w:body/w:commentRangeStart[@w:id=$id]/following-sibling::*[not(self::w:commentRangeEnd[@w:id=$id] or w:commentRangeEnd[@w:id=$id])]",
    #         id="1",
    #         **ns,
    #     )

    # @property
    # def comment_runs_outside2(self):
    #     # This works for comments outside paragraphs
    #     return self._doc.xml["word/document.xml"].xpath(
    #         "//w:commentRangeStart[@w:id=$id]/../following::*[node(self::w:r) and (not(self::w:commentRangeEnd[@w:id=$id] or preceding-sibling::w:commentRangeEnd[@w:id=$id]))]",
    #         id="7",
    #         **ns,
    #     )

    # # @property
    # # def comment_runs_outside3(self):
    # #     # This works for comments outside paragraphs
    # #     return self._doc.xml["word/document.xml"].xpath(
    # #         "w:body/w:p[preceding-sibling::w:commentRangeStart[@w:id=$id] or w:commentRangeStart[@w:id=$id]]|following-sibling::*)[not(self::w:commentRangeEnd[@w:id=$id] or preceding-sibling::w:p/w:commentRangeEnd[@w:id=$id])]",
    # #         id="5",
    # #         **ns
    # #     )

    # @property
    # def comments_inside_para(self):
    #     return self._doc.xml["word/document.xml"].xpath(
    #         "//w:commentRangeEnd[parent::w:p]/@w:id", **ns
    #     )

    # @property
    # def comments_outside_para(self):
    #     return self._doc.xml["word/document.xml"].xpath(
    #         "//w:commentRangeEnd[parent::w:body]/@w:id", **ns
    #     )

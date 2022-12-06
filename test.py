from lxml.builder import ElementMaker
from lxml import etree


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


# no = 10
# footnote_number_run = RUN(
#     RUNPROPS(
#         SCRIPTHEIGHT(val="2"),
#     ),
#     TEXT(str(no)),
# )

# for para in paras:
#     para.insert(0, footnote_number_run)

run_text = "Run. "

different_level_starts_and_ends = DOC(
    BODY(
        COMMENTSTART(val="1"),
        PARA(
            RUN(TEXT(run_text)),
            RUN(TEXT(run_text)),
            COMMENTEND(val="1"),
        ),
        PARA(
            RUN(TEXT(run_text)),
            COMMENTSTART(val="2"),
            RUN(TEXT(run_text)),
            RUN(TEXT(run_text)),
        ),
        COMMENTEND(val="2"),
    ),
)

overlapping_comments = DOC(
    BODY(
        PARA(
            COMMENTSTART(val="1"),
            RUN(TEXT(run_text)),
            COMMENTSTART(val="2"),
            RUN(TEXT(run_text)),
            COMMENTEND(val="2"),
        ),
        PARA(
            RUN(TEXT(run_text)),
            RUN(TEXT(run_text)),
            COMMENTSTART(val="3"),
            COMMENTEND(val="1"),
            RUN(TEXT(run_text)),
            COMMENTEND(val="3"),
        ),
    ),
)

multi_paragraph_comment = DOC(
    BODY(
        COMMENTSTART(val="1"),
        PARA(
            RUN(TEXT(run_text)),
            RUN(TEXT(run_text)),
        ),
        PARA(
            RUN(TEXT(run_text)),
            RUN(TEXT(run_text)),
        ),
        PARA(
            RUN(TEXT(run_text)),
            RUN(TEXT(run_text)),
        ),
        COMMENTEND(val="1"),
    ),
)

comment_inside_paragraph = DOC(
    BODY(
        PARA(
            RUN(TEXT(run_text)),
            COMMENTSTART(val="1"),
            RUN(TEXT(run_text)),
            RUN(TEXT(run_text)),
            COMMENTEND(val="1"),
            RUN(TEXT(run_text)),
        ),
        PARA(
            COMMENTSTART(val="2"),
            RUN(TEXT(run_text)),
            RUN(TEXT(run_text)),
            COMMENTEND(val="2"),
        ),
    ),
)

empty_paragraphs_and_runs = DOC(
    BODY(
        PARA(),
        PARA(
            RUN(TEXT(run_text)),
            RUN(),
            RUN(TEXT(run_text)),
        ),
        PARA(),
        PARA(
            RUN(),
            RUN(TEXT(run_text)),
            RUN(),
        ),
        PARA(
            RUN(),
        ),
        PARA(
            RUN(TEXT(run_text)),
        ),
    ),
)

documents = (
    empty_paragraphs_and_runs,
    comment_inside_paragraph,
    multi_paragraph_comment,
    overlapping_comments,
    different_level_starts_and_ends,
)


class Tests:
    def __init__(self):
        self.store = self.make_list()

    def make_list(self):
        return ["1", "2", "3"]


r = Tests()
q = Tests()
print(r.store, q.store)
r.store.append("i")
print(r.store, q.store)

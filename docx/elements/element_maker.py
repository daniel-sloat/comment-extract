from lxml.builder import ElementMaker

from docx.elements.properties import PropDecode


class DocxElementMaker:
    ns_maker = {
        "namespace": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "nsmap": {
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "xml": "http://www.w3.org/XML/1998/namespace",
        },
    }

    def __init__(self):
        self.E = ElementMaker(**self.ns_maker)

    def RUN(self, *elements, **kwargs):
        return self.E.r(*elements, **kwargs)

    def PARA(self, *elements, **kwargs):
        return self.E.p(*elements, **kwargs)

    def TEXT(self, text, *args, **kwargs):
        return self.E.t(text, *args, **kwargs)

    def RUNPROPS(self, *props, **kwargs):
        return self.E.rPr(*props, **kwargs)

    def PARAPROPS(self, *props, **kwargs):
        return self.E.pPr(*props, **kwargs)

    def BOLD(self):
        return self.E.b()

    def ITALIC(self):
        return self.E.i()

    def UNDERLINE(self):
        return self.E.u(val="single")

    def D_UNDERLINE(self):
        return self.E.u(val="double")

    def STRIKE(self):
        return self.E.strike()

    def D_STRIKE(self):
        return self.E.dstrike()

    def SUBSCRIPT(self):
        return self.E.vertAlign({"val": "subscript"})

    def SUPERSCRIPT(self):
        return self.E.vertAlign({"val": "superscript"})


def property_converter(E, props):
    prop_check = PropDecode(props)
    if prop_check.bold:
        yield E.BOLD
    if prop_check.italic:
        yield E.ITALIC
    if prop_check.underline:
        yield E.UNDERLINE
    if prop_check.d_underline:
        yield E.D_UNDERLINE
    if prop_check.strike:
        yield E.STRIKE
    if prop_check.d_strike:
        yield E.D_STRIKE
    if prop_check.subscript:
        yield E.SUBSCRIPT
    if prop_check.superscript:
        yield E.SUPERSCRIPT


def paragraph_maker(*runs):
    preserve = {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}
    E = DocxElementMaker()
    run_data = []
    for run in runs:
        props, text = run
        properties = property_converter(E, props)
        leading_trailing_whitespace = text[0] == " " or text[-1] == " "
        if leading_trailing_whitespace:
            text_e = E.TEXT(text, **preserve)
        else:
            text_e = E.TEXT(text)
        run_data.append(E.RUN(*(E.RUNPROPS(*properties), text_e)))
    return E.PARA(*run_data)


# from lxml import etree


# runs = [
#     ({"vertAlign": {"val": "superscript"}}, "1"),
#     ({"vertAlign": {"val": "superscript"}}, "1"),
#     ({}, " "),
#     ({}, "Footnote text."),
#     ({}, "Footnote text."),
#     ({"b": {}}, " Bold text."),
#     ({"strike": {}}, " Strike text."),
#     ({"u": {}}, " Underline text."),
#     ({"i": {}}, " Italic text."),
#     ({"u": {"val": "double"}}, " Double underlined text."),
#     ({"dstrike": {}}, " Double strike text."),
#     ({"vertAlign": {"val": "subscript"}}, "Subscript!"),
# ]

# paragraph = paragraph_maker(*runs)
# # print(paragraph)
# # print(etree.tostring(paragraph, encoding="unicode", pretty_print=True))

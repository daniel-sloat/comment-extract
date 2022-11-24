from docx.ooxml_ns import ns
from docx.baseelement import BaseDOCXElement


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


class Bubble(TextElement):
    @property
    def text(self):
        return "\n".join(z for z in ("".join(run.text for run in para.runs) for para in self.paragraphs))
        
    @property
    def paragraphs(self):
        return [Paragraph(el) for el in self.element.xpath("w:p", **ns)]


class Paragraph(TextElement):
    def __iter__(self):
        return iter(self.runs)

    @property
    def runs(self):
        return [Run(element) for element in self.element.xpath("w:r", **ns)]

    @property
    def props(self):
        return {
            (element := PropElement(el)).tag: element.attrib 
            for el in self.element.xpath("w:pPr/*", **ns)
            if not el.xpath("self::w:rPr", **ns)
        }

    @property
    def glyph_props(self):
        return {
            (element := PropElement(el)).tag: element.attrib
            for el in self.element.xpath("w:pPr/w:rPr/*", **ns)
        }

    @property
    def style(self):
        return self.element.xpath("string(w:pPr/w:pStyle/@w:val)", **ns)


class ParagraphStyled(Paragraph):
    def __init__(self, element, styles):
        super().__init__(element)
        self._styles = styles

    @property
    def style_props(self):
        return self._styles.styles_map.get(self.style,{}).get("para",{})


class CommentParagraph(ParagraphStyled):
    def __init__(self, element, comment):
        super().__init__(element,comment._comments._doc.styles)
        self._comment = comment

    @property
    def text(self):
        return "".join(run.text for run in self.runs)

    @property
    def runs(self):
        def get_runs(children):
            select_runs = []
            for child in children:
                if child.xpath("self::w:r", **ns):
                    select_runs.append(CommentRun(child,self))
                if child.xpath(f"self::w:commentRangeEnd[@w:id={self._comment._id}]", **ns):
                    break
            return select_runs

        para_elements = (el for el in self.element.xpath("*", **ns))
        comment_start_paragraph = self.element.xpath(
            f"w:commentRangeStart[@w:id={self._comment._id}]", **ns
        )

        if comment_start_paragraph:
            for element in para_elements:
                if element.xpath(
                    f"self::w:commentRangeStart[@w:id={self._comment._id}]", **ns
                ):
                    return get_runs(para_elements)
        else:
            return get_runs(para_elements)


class Run(TextElement):
    @property
    def props(self):
        return {
            (element := PropElement(el)).tag: element.attrib 
            for el in self.element.xpath("w:rPr/*", **ns)
        }

    @property
    def style(self):
        return self.element.xpath("string(w:rPr/w:rStyle/@w:val)", **ns)

    @property
    def footnote(self):
        return self.element.xpath("string(w:footnoteReference/@w:id)", **ns)

    @property
    def endnote(self):
        return self.element.xpath("string(w:endnoteReference/@w:id)", **ns)


class RunStyled(Run):
    def __init__(self, element, styles):
        super().__init__(element)
        self._styles = styles

    @property
    def display_props(self):
        pstyle = self.element.xpath("string(parent::w:p/w:pPr/w:pStyle/@w:val)", **ns)
        pstyle_run_props = self._styles.styles_map.get(pstyle,{}).get("run",{}) # Returns ChainMap
        return (
            pstyle_run_props
                .new_child(self._styles.styles_map.get(self.style,{}).get("run",{}))
                .new_child(self.props)
        )


class CommentRun(RunStyled):
    def __init__(self, element, para):
        super().__init__(element,para._styles)
        self._notes = para._comment._comments._doc.notes

    @property
    def fn(self):
        return self._notes.footnotes.get(self.footnote,"")


class PropElement(BaseDOCXElement):
    pass

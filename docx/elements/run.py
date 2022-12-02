from collections import ChainMap
from functools import cached_property
import re
from docx.elements.elements import PropElement, TextElement
from docx.ooxml_ns import ns


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

    @cached_property
    def display_props(self):
        pstyle = self.element.xpath("string(parent::w:p/w:pPr/w:pStyle/@w:val)", **ns)
        if pstyle:
            pstyle_run_props = self._styles.styles_map.get(pstyle, {}).get(
                "run", {}
            )  # Returns ChainMap
            return pstyle_run_props.new_child(
                self._styles.styles_map.get(self.style, {}).get("run", {})
            ).new_child(self.props)
        return ChainMap(
            self._styles.styles_map.get(self.style, {}).get("run", {}), self.props
        )

    def simple_props(self):
        return "".join(
            [
                "b" if self.bold else "",
                "i" if self.italic else "",
                "u" if self.underline else "",
                "s" if self.strike else "",
                "w" if self.d_underline else "",
                "z" if self.d_strike else "",
                "v" if self.subscript else "",
                "x" if self.superscript else "",
            ]
        )

    def asdict(self):
        d = {}
        if self.bold:
            d["bold"] = True
        if self.italic:
            d["italic"] = True
        if self.underline:
            d["underline"] = 1
        elif self.d_underline:
            d["underline"] = 2
        if self.strike:
            d["font_strikeout"] = True
        elif self.d_strike:
            # No double strikethrough in Excel
            d["font_strikeout"] = True
            d["font_color"] = "#FF0000"  # #FF0000 = red
        if self.subscript:
            d["font_script"] = 2
        elif self.superscript:
            d["font_script"] = 1
        return d

    def _toggled(self, prop):
        # b, i, strike, dstrike are 'toggle' formats.
        return prop in self.display_props and (
            not self.display_props.get(prop, {})
            or self.display_props.get(prop, {}).get("val") in ("1", "on", "true")
        )

    @property
    def bold(self):
        return self._toggled("b")

    @property
    def italic(self):
        return self._toggled("i")

    @property
    def underline(self):
        return "u" in self.display_props and not re.search(
            "[D|d]ouble|^none$", self.display_props.get("u", {}).get("val", "")
        )

    @property
    def strike(self):
        return self._toggled("strike")

    @property
    def d_underline(self):
        return "u" in self.display_props and re.search(
            "[D|d]ouble", self.display_props.get("u", {}).get("val", "")
        )

    @property
    def d_strike(self):
        return self._toggled("dstrike")

    @property
    def subscript(self):
        return (
            "vertAlign" in self.display_props
            and self.display_props.get("vertAlign", {}).get("val", "") == "subscript"
        )

    @property
    def superscript(self):
        return (
            "vertAlign" in self.display_props
            and self.display_props.get("vertAlign", {}).get("val", "") == "superscript"
        )

    @property
    def caps(self):
        return self._toggled("caps")


class CommentRun(RunStyled):
    def __init__(self, element, para):
        super().__init__(element, para._styles)
        self._notes = para._comment._comments._doc.notes

    @property
    def fn(self):
        return self._notes.footnotes.get(self.footnote, "")

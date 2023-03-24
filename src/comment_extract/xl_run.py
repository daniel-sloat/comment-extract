class XlRun:
    """Adaptor between DOCX Run object and XLSXWriter Format object."""

    def __init__(self, run):
        self._docxrun = run

    @property
    def text(self):
        return self._docxrun.text

    @property
    def props(self):
        fmt = {}
        if self._docxrun.props.decode.bold:
            fmt["bold"] = True
        if self._docxrun.props.decode.italic:
            fmt["italic"] = True
        if self._docxrun.props.decode.underline:
            fmt["underline"] = 1
        elif self._docxrun.props.decode.d_underline:
            fmt["underline"] = 2
        if self._docxrun.props.decode.strike:
            fmt["font_strikeout"] = True
        elif self._docxrun.props.decode.d_strike:
            # No double strikethrough in Excel
            fmt["font_strikeout"] = True
            fmt["font_color"] = "#FF0000"  # #FF0000 = red
        if self._docxrun.props.decode.subscript:
            fmt["font_script"] = 2
        elif self._docxrun.props.decode.superscript:
            fmt["font_script"] = 1
        return fmt

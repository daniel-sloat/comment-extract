import re


class PropDecode:
    """Decodes OOXML format properties."""

    def __init__(self, props):
        self._props = props

    def _toggled(self, prop):
        # b, i, strike, dstrike (among others) are 'toggled'
        try:
            toggle1 = not self._props[prop]  # {"b": {}}
            toggle2 = self._props[prop].get("val", "") in ("1", "on", "true")
        except KeyError:
            return False
        else:
            return toggle1 or toggle2

    @property
    def bold(self):
        return self._toggled("b")

    @property
    def italic(self):
        return self._toggled("i")

    @property
    def underline(self):
        return "u" in self._props and not re.search(
            "[D|d]ouble|^none$", self._props.get("u", {}).get("val", "")
        )

    @property
    def strike(self):
        return self._toggled("strike")

    @property
    def d_underline(self):
        return re.search("[D|d]ouble", self._props.get("u", {}).get("val", ""))

    @property
    def d_strike(self):
        return self._toggled("dstrike")

    @property
    def subscript(self):
        return self._props.get("vertAlign", {}).get("val", "") == "subscript"

    @property
    def superscript(self):
        return self._props.get("vertAlign", {}).get("val", "") == "superscript"

    @property
    def caps(self):
        return self._toggled("caps")

    @property
    def color(self):
        return self._props.get("color", {}).get("val", "")

    @property
    def emboss(self):
        return self._toggled("emboss")

    @property
    def imprint(self):
        return self._toggled("imprint")

    @property
    def outline(self):
        return self._toggled("outline")

    @property
    def shadow(self):
        return self._toggled("shadow")

    @property
    def smallcaps(self):
        return self._toggled("smallCaps")

    @property
    def size(self):
        return self._props.get("sz", {}).get("val", "")

    @property
    def vanish(self):
        return self._toggled("vanish")

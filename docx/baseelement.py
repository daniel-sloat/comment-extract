import re


class BaseDOCXElement:
    def __init__(self, element):
        self.element = element

    def __repr__(self):
        return f"{self.__class__.__name__}(tag='{self.tag}')"

    @property
    def tag(self):
        return self._strip_uri(self.element.tag)

    @property
    def attrib(self):
        return {self._strip_uri(tag): el for tag, el in self.element.items()}

    @staticmethod
    def _strip_uri(tag):
        fix_id = re.sub("id", "_id", tag if tag else "")
        return re.sub("{.*}", "", fix_id)

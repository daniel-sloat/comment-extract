from reprlib import Repr

limit_repr_text = Repr()


class ParagraphGroup:
    def __init__(self, paragraphs):
        self.paragraphs = paragraphs

    def __repr__(self):
        return f"{self.__class__.__name__}(text='{limit_repr_text.repr(self.text)}')"

    def __str__(self):
        return self.text

    def __iter__(self):
        return iter(self.paragraphs)

    def __len__(self):
        return len(self.paragraphs)

    def __getitem__(self, key):
        return self.paragraphs[key]

    def __bool__(self):
        return bool(self.paragraphs)

    @property
    def text(self):
        return "\n".join(para.text for para in self.paragraphs)

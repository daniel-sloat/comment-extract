DELIMITER = "-"


class FileNameParser:
    def __init__(self, filename):
        self.path = filename
        self.doc_number, self.commenter_code = self._partition_filename("-")

    def _partition_filename(self, delimiter):
        doc_number, _, commenter_code = self.path.stem.partition(delimiter)
        if doc_number.isnumeric():
            return int(doc_number), commenter_code.strip()
        elif not doc_number:
            return None, self.path.stem.strip()
        else:
            return doc_number, self.path.stem.strip()

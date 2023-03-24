class FileNameParser:
    def __init__(self, filename, filename_delimiter=""):
        self.path = filename
        self.doc_number, self.commenter_code = self._partition_filename(filename_delimiter)

    def _partition_filename(self, delimiter):
        if delimiter:
            doc_number, _, commenter_code = self.path.stem.partition(delimiter)
            if doc_number.isnumeric():
                return int(doc_number), commenter_code.strip()
            return doc_number, self.path.stem.strip()
        return None, self.path.stem.strip()

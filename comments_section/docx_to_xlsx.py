from itertools import groupby


class DOCX_XLSXAdapter:
    def __init__(self, comment_record):
        self.comment_record = comment_record

    def header(self, add_columns):
        column_names = ["Comment Number", "Runs"]
        self.comment_record[0]

    def prepared_data(self):
        p = []
        count = 0
        for comments in self.comment_record:
            path = comments._doc.file
            doc_comment_count = 0
            for comment in comments:
                q = {}

                count += 1
                doc_comment_count += 1

                q["filename"] = path.name
                q["folder"] = path.parent.name
                q["comment_count"] = count
                q["doc_comment_count"] = doc_comment_count
                q["author"] = comment.author
                q["initials"] = comment.initials
                q["date"] = comment.date
                q["bubble"] = comment.bubble.text

                runs = []
                for paragraph in (o := comment.paragraphs):
                    newline = "\n" if paragraph is not o[-1] else ""
                    for key_format, group in groupby(
                        paragraph.runs, lambda x: x.asdict()
                    ):
                        text_list = []
                        for run in group:
                            text_list.append(run.text)
                        text = "".join(text_list).strip()
                        if text:
                            g = (key_format, text)
                            runs.extend(g)
                        if newline:
                            runs.append(newline)
                q["runs"] = runs
                if runs:
                    yield q

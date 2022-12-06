"""Main document processing script"""

from pathlib import Path

from comments_section.get_comments import get_comments
from comments_section.public_record import CommentRecord
from config import toml_config
from docx.docx import Document
from logger.logger import logger_init


def docx_in_folder(folder_path: str):
    """Returns all docx files in folder."""
    return (x for x in Path(folder_path).glob("*.docx") if x.is_file())


@logger_init
def main():
    config = toml_config.load()
    comment_record = CommentRecord()
    # documents = docx_in_folder(config["folder_path"])
    # comments = get_comments(documents)
    for document in docx_in_folder(config["folder_path"]):
        doc = Document(document)
        comments = doc.comments
        # for comment in comments:
        #     for paragraph in comment:
        #         for run in paragraph:
        #             print(run.text)
        comment_record.append(comments)
    comment_record.to_excel(**config)


import cProfile

if __name__ == "__main__":
    main()
    # cProfile.run("main()", sort="tottime")

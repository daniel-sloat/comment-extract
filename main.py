"""Document processing script"""

from pathlib import Path

import tomllib
from docx_comments import Document

from comment_extract import CommentRecord
from comment_extract.logger.logger import logger_init


def docx_in_folder(folder_path: str):
    """Returns all docx files in folder."""
    return (x for x in Path(folder_path).glob("*.docx") if x.is_file())


@logger_init()
def main():
    with open("config.toml", "rb") as toml:
        config = tomllib.load(toml)

    all_comments = []
    for document in docx_in_folder(config["folder_path"]):
        doc = Document(document)
        comments = doc.comments
        all_comments.append(comments)

    comment_record = CommentRecord(all_comments)
    comment_record.to_excel(**config)


if __name__ == "__main__":
    main()

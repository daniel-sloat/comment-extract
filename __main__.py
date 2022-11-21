"""Main document processing script"""

from pathlib import Path
from config import toml_config
from logger.logger import logger_init

from docx.document import Document
from comments_section.public_record import CommentRecord


def documents_in_folder(folder_path: str):
    """Returns all docx files in folder."""
    return (x for x in Path(folder_path).glob("*.docx") if x.is_file())


@logger_init
def main():
    config = toml_config.load()
    comment_record = CommentRecord()
    for document in documents_in_folder(config["folder_path"]):
        doc = Document(document)
        # comments = doc.comments
        styles = doc.styles
        for style, props in styles.inherited_styles.items():
            for k,v in props["run"].items():
                if "" in k:
                    print(k)
    #     print(styles["Heading1"]._para_and_run)
    #     comment_record.append(comments)
    # comment_record.to_excel("output/comments.xlsx", **config)


if __name__ == "__main__":
    main()

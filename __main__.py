from . import docx_dataclasses as docx

# import xlsx

from pathlib import Path
from config import toml_config
from logger.logger import logger_init

from docx.document import Document


def documents_in_folder(folder_path: str):
    """Returns all docx files in folder."""
    return (x for x in Path(folder_path).glob("*.docx") if x.is_file())

@logger_init
def main():
    config = toml_config.load()
    public_record = []
    for document in documents_in_folder("input/45-Day"):
        print(document)
        doc = Document(document)
        comments = doc.comments.comments
        # styles = doc.styles.styles
        for comment in comments:
            for para in comment.paragraphs:
                # print(para.props)
                for run in para.comment_runs:
                    for prop in run.props:
                        if prop.tag == "rStyle":
                            print(prop.attrib)
        # break
        comments = docx.Comments(doc, **config)
        public_record.append(comments)
    # xlsx.Workbook(public_record, **config)


if __name__ == "__main__":
    main()

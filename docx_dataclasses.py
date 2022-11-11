"""Main document processing script"""

from pathlib import Path

from docx.document import Document


def documents_in_folder(folder_path: str):
    """Returns all docx files in folder."""
    return (x for x in Path(folder_path).glob("*.docx") if x.is_file())


def main():
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


if __name__ == "__main__":
    main()

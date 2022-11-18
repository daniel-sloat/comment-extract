from pathlib import Path

from docx.document import Document


def documents_in_folder(folder_path: str):
    """Returns all docx files in folder."""
    return (x for x in Path(folder_path).glob("*.docx") if x.is_file())


def retrieve_comments(folder_path):
    for document in documents_in_folder(folder_path):
        doc = Document(document)
        comments = doc.comments
        yield comments
